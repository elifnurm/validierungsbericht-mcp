#!/usr/bin/env python3
import base64
import copy
import hashlib
import io
import json
import os
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
ET.register_namespace("w", W_NS)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
MASTER_TEMPLATE = PROJECT_ROOT / "skills" / "validierungsbericht" / "references" / "Validierungsbericht_Template_Plugin_Dynamisch.docx"
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
COLOR_MAP = {
    "Grün": "00B050",
    "Grüngelb": "92D050",
    "Gelb": "FFFF00",
    "Orange": "FFC000",
    "Rot": "FF0000",
}

def _text(el):
    return "".join((t.text or "") for t in el.iter(W + "t"))

def _replace_text(el, mapping):
    for t in el.iter(W + "t"):
        if not t.text:
            continue
        value = t.text
        for key, replacement in mapping.items():
            value = value.replace(key, "" if replacement is None else str(replacement))
        t.text = value

def _parent_map(root):
    return {c: p for p in root.iter() for c in p}

def _set_cell_fill_for_marker(root, marker, color_name=None, fallback_text=None):
    color = COLOR_MAP.get(color_name) if color_name else None
    parents = _parent_map(root)
    for t in list(root.iter(W + "t")):
        if marker not in (t.text or ""):
            continue
        cur = t
        cell = None
        while cur in parents:
            cur = parents[cur]
            if cur.tag == W + "tc":
                cell = cur
                break
        if cell is not None:
            tc_pr = cell.find(W + "tcPr")
            if tc_pr is None:
                tc_pr = ET.Element(W + "tcPr")
                cell.insert(0, tc_pr)
            shd = tc_pr.find(W + "shd")
            if shd is None:
                shd = ET.SubElement(tc_pr, W + "shd")
            shd.set(W + "val", "clear")
            shd.set(W + "color", "auto")
            shd.set(W + "fill", color or "FFFFFF")
        t.text = (t.text or "").replace(marker, fallback_text or "")

def _find_row_with_marker(table, marker):
    for row in table.findall(W + "tr"):
        if marker in _text(row):
            return row
    return None

def _duplicate_template_row(table, marker, records, populate):
    row = _find_row_with_marker(table, marker)
    if row is None:
        return
    children = list(table)
    idx = children.index(row)
    for rec in records:
        new_row = copy.deepcopy(row)
        populate(new_row, rec)
        table.insert(idx, new_row)
        idx += 1
    table.remove(row)

def _set_last_cell_fill(row, color_name):
    cells = row.findall(W + "tc")
    if not cells:
        return
    cell = cells[-1]
    tc_pr = cell.find(W + "tcPr")
    if tc_pr is None:
        tc_pr = ET.Element(W + "tcPr")
        cell.insert(0, tc_pr)
    shd = tc_pr.find(W + "shd")
    if shd is None:
        shd = ET.SubElement(tc_pr, W + "shd")
    shd.set(W + "val", "clear")
    shd.set(W + "color", "auto")
    shd.set(W + "fill", COLOR_MAP[color_name])

def _populate_finding_row(row, finding, section_value=None):
    had_ampel_marker = "{{FESTSTELLUNG_AMPEL}}" in _text(row)
    _replace_text(row, {
        "{{ABSCHNITT}}": section_value or finding.get("abschnitt", ""),
        "{{FESTSTELLUNG_ID}}": finding.get("id", ""),
        "{{FESTSTELLUNG_TEXT}}": finding.get("beschreibung", ""),
        "{{MASSNAHME_ID}}": finding.get("massnahme_id", ""),
        "{{MASSNAHME_TEXT}}": finding.get("massnahme", ""),
        "{{MASSNAHME_FRIST}}": finding.get("erledigung_bis", ""),
        "{{MASSNAHME_ZUSTAENDIGKEIT}}": finding.get("zustaendigkeit", ""),
    })
    if had_ampel_marker:
        _set_cell_fill_for_marker(row, "{{FESTSTELLUNG_AMPEL}}", finding.get("ampelfarbe"))
    else:
        _set_last_cell_fill(row, finding.get("ampelfarbe"))

def _populate_anomaly_row(row, anomaly, section_value=None, marker="{{ABSCHNITT}}"):
    _replace_text(row, {
        marker: section_value or anomaly.get("abschnitt", ""),
        "{{AUFFAELLIGKEIT_ID}}": anomaly.get("id", ""),
        "{{AUFFAELLIGKEIT_TEXT}}": anomaly.get("beschreibung", ""),
        "{{BEGRUENDUNG_ANGEMESSENHEIT}}": anomaly.get("begruendung_angemessenheit", ""),
    })
    _set_cell_fill_for_marker(row, "{{AUFFAELLIGKEIT_AMPEL}}", anomaly.get("ampelfarbe"))

def _populate_section_summary_row(row, section):
    _set_cell_fill_for_marker(
        row, "{{AMPEL_{{ABSCHNITT_NUMMER}}}}", section.get("ampelfarbe"),
        fallback_text="Offener Prüfpunkt" if not section.get("ampelfarbe") else "",
    )
    _replace_text(row, {
        "{{ABSCHNITT_NUMMER}}": section.get("abschnitt", ""),
        "{{ABSCHNITT_NAME}}": section.get("validierungsgegenstand", ""),
    })

def _replace_global_assessment(body, payload):
    _replace_text(body, {"{{GESAMTBEURTEILUNG}}": payload.get("gesamtbeurteilung", "")})
    if payload.get("gesamt_ampelfarbe"):
        _set_cell_fill_for_marker(body, "{{GESAMTBEURTEILUNG_AMPEL}}", payload.get("gesamt_ampelfarbe"))
    else:
        _replace_text(body, {"{{GESAMTBEURTEILUNG_AMPEL}}": ""})

def _process_chapter_32(body, sections):
    findings, anomalies = [], []
    for sec in sections:
        for f in sec.get("feststellungen", []):
            item = dict(f); item["abschnitt"] = sec.get("abschnitt", ""); findings.append(item)
        for a in sec.get("auffaelligkeiten", []):
            item = dict(a); item["abschnitt"] = sec.get("abschnitt", ""); anomalies.append(item)
    children = list(body)
    begin_idx = next((i for i, ch in enumerate(children) if "BEGIN_REPEAT_VALIDIERUNGSGEGENSTAND" in _text(ch)), len(children))
    for table in [ch for ch in children[:begin_idx] if ch.tag == W + "tbl"]:
        txt = _text(table)
        if "{{ABSCHNITT_NUMMER}}" in txt and "{{ABSCHNITT_NAME}}" in txt:
            _duplicate_template_row(table, "{{ABSCHNITT_NUMMER}}", sections, _populate_section_summary_row)
        elif "{{FESTSTELLUNG_ID}}" in txt:
            _duplicate_template_row(table, "{{FESTSTELLUNG_ID}}", findings, lambda row, rec: _populate_finding_row(row, rec))
        elif "{{AUFFAELLIGKEIT_ID}}" in txt:
            _duplicate_template_row(table, "{{AUFFAELLIGKEIT_ID}}", anomalies, lambda row, rec: _populate_anomaly_row(row, rec))

def _process_repeat_block(body, sections):
    children = list(body)
    start = next((i for i, ch in enumerate(children) if "BEGIN_REPEAT_VALIDIERUNGSGEGENSTAND" in _text(ch)), None)
    end = next((i for i, ch in enumerate(children) if "END_REPEAT_VALIDIERUNGSGEGENSTAND" in _text(ch)), None)
    if start is None or end is None or end <= start:
        raise ValueError("Repeat-Block im Mastertemplate nicht gefunden.")
    template_nodes = children[start + 1:end]
    insert_at = start
    for ch in children[start:end + 1]:
        body.remove(ch)
    for sec in sections:
        block = [copy.deepcopy(ch) for ch in template_nodes]
        for node in block:
            _replace_text(node, {"{{ABSCHNITT_NAME}}": sec.get("validierungsgegenstand", "")})
            if node.tag == W + "tbl" and "{{TEILBEURTEILUNG}}" in _text(node):
                _replace_text(node, {"{{TEILBEURTEILUNG}}": sec.get("teilbeurteilung", "")})
                _set_cell_fill_for_marker(node, "{{TEILBEURTEILUNG_AMPEL}}", sec.get("ampelfarbe"),
                    fallback_text="Offener Prüfpunkt" if not sec.get("ampelfarbe") else "")
            elif node.tag == W + "tbl" and "{{FESTSTELLUNG_ID}}" in _text(node):
                _duplicate_template_row(node, "{{FESTSTELLUNG_ID}}", sec.get("feststellungen", []),
                    lambda row, rec, s=sec: _populate_finding_row(row, rec, s.get("abschnitt", "")))
            elif node.tag == W + "tbl" and "{{AUFFAELLIGKEIT_ID}}" in _text(node):
                _duplicate_template_row(node, "{{AUFFAELLIGKEIT_ID}}", sec.get("auffaelligkeiten", []),
                    lambda row, rec, s=sec: _populate_anomaly_row(row, rec, s.get("abschnitt", ""), marker="{{ABSCHNITT_NUMMER}}"))
        for node in block:
            body.insert(insert_at, node); insert_at += 1
    for ch in list(body):
        if "{{HINWEIS: Den folgenden Block einmal pro eindeutigem Abschnitt der Eingabetabelle duplizieren." in _text(ch):
            body.remove(ch)

def _validate_payload(payload):
    sections = payload.get("sections")
    if not isinstance(sections, list) or not sections:
        raise ValueError("sections muss eine nichtleere Liste sein.")
    seen = {}
    for sec in sections:
        abschnitt = str(sec.get("abschnitt", "")).strip()
        subject = str(sec.get("validierungsgegenstand", "")).strip()
        if not abschnitt or not subject:
            raise ValueError("Jeder Abschnitt benötigt abschnitt und validierungsgegenstand.")
        if abschnitt in seen and seen[abschnitt] != subject:
            raise ValueError(f"Widersprüchliche Validierungsgegenstände für Abschnitt {abschnitt}.")
        seen[abschnitt] = subject
        color = sec.get("ampelfarbe")
        if color is not None and color not in COLOR_MAP:
            raise ValueError(f"Unbekannte Ampelfarbe für Abschnitt {abschnitt}: {color}")
        for kind in ("feststellungen", "auffaelligkeiten"):
            items = sec.get(kind, []) or []
            if not isinstance(items, list):
                raise ValueError(f"{kind} in Abschnitt {abschnitt} muss eine Liste sein.")
            for item in items:
                if not str(item.get("id", "")).strip():
                    raise ValueError(f"Fehlende ID in {kind} von Abschnitt {abschnitt}; ID-Dialog muss vor dem Rendern abgeschlossen sein.")
                if not str(item.get("beschreibung", "")).strip():
                    raise ValueError(f"Fehlende Beschreibung in {kind} von Abschnitt {abschnitt}.")
                if item.get("ampelfarbe") not in COLOR_MAP:
                    raise ValueError(f"Unbekannte Ampelfarbe in {kind} von Abschnitt {abschnitt}: {item.get('ampelfarbe')}")

def render_docx(payload):
    _validate_payload(payload)
    if not MASTER_TEMPLATE.is_file():
        raise FileNotFoundError(f"Mastertemplate fehlt: {MASTER_TEMPLATE}")
    with zipfile.ZipFile(MASTER_TEMPLATE, "r") as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        body = root.find(W + "body")
        if body is None:
            raise ValueError("Ungültiges DOCX: word/document.xml enthält keinen Body.")
        sections = payload["sections"]
        _replace_global_assessment(body, payload)
        _process_chapter_32(body, sections)
        _process_repeat_block(body, sections)
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                zout.writestr(item, data)
    return out.getvalue()

def _tool_schema():
    finding = {"type":"object","required":["id","beschreibung","ampelfarbe"],"properties":{
        "id":{"type":"string"},"beschreibung":{"type":"string"},
        "ampelfarbe":{"type":"string","enum":list(COLOR_MAP)},
        "massnahme_id":{"type":"string"},"massnahme":{"type":"string"},
        "erledigung_bis":{"type":"string"},"zustaendigkeit":{"type":"string"}},"additionalProperties":False}
    anomaly = {"type":"object","required":["id","beschreibung","ampelfarbe"],"properties":{
        "id":{"type":"string"},"beschreibung":{"type":"string"},
        "ampelfarbe":{"type":"string","enum":list(COLOR_MAP)},
        "begruendung_angemessenheit":{"type":"string"}},"additionalProperties":False}
    section = {"type":"object","required":["abschnitt","validierungsgegenstand","teilbeurteilung","feststellungen","auffaelligkeiten"],"properties":{
        "abschnitt":{"type":"string"},"validierungsgegenstand":{"type":"string"},"teilbeurteilung":{"type":"string"},
        "ampelfarbe":{"type":["string","null"],"enum":list(COLOR_MAP)+[None]},
        "feststellungen":{"type":"array","items":finding},"auffaelligkeiten":{"type":"array","items":anomaly}},"additionalProperties":False}
    return {"type":"object","required":["sections","gesamtbeurteilung"],"properties":{
        "sections":{"type":"array","minItems":1,"items":section},"gesamtbeurteilung":{"type":"string"},
        "gesamt_ampelfarbe":{"type":["string","null"],"enum":list(COLOR_MAP)+[None]},
        "output_filename":{"type":"string","default":"Validierungsbericht_Entwurf.docx"}},"additionalProperties":False}

def _sanitize_filename(name):
    name = os.path.basename(name or "Validierungsbericht_Entwurf.docx")
    if not name.lower().endswith(".docx"): name += ".docx"
    return re.sub(r"[^A-Za-z0-9ÄÖÜäöüß._ -]", "_", name)

def handle_mcp(req):
    method = req.get("method")
    if method == "initialize":
        return {"protocolVersion":"2025-06-18","capabilities":{"tools":{}},"serverInfo":{"name":"validierungsbericht-docx","version":"0.2.0"}}
    if method == "tools/list":
        return {"tools":[{"name":"render_validierungsbericht_docx",
            "description":"Erzeugt deterministisch einen DOCX-Entwurf aus dem verbindlichen Mastertemplate. Erwartet bereits fachlich aufbereitete Abschnitte, Teil-/Gesamtbeurteilungen, IDs und Ampelfarben; es erfindet keine fachlichen Inhalte.",
            "inputSchema":_tool_schema()}]}
    if method == "tools/call":
        params=req.get("params",{})
        if params.get("name")!="render_validierungsbericht_docx": raise ValueError("Unbekanntes Tool.")
        args=params.get("arguments",{})
        data=render_docx(args)
        filename=_sanitize_filename(args.get("output_filename"))
        digest=hashlib.sha256(data).hexdigest()
        return {"content":[
            {"type":"text","text":f"DOCX erzeugt: {filename} ({len(data)} Bytes, SHA-256 {digest})"},
            {"type":"resource","resource":{"uri":f"file:///{filename}","mimeType":MIME_DOCX,"blob":base64.b64encode(data).decode("ascii")}}
        ],"structuredContent":{"filename":filename,"size_bytes":len(data),"sha256":digest}}
    if method and method.startswith("notifications/"): return None
    raise ValueError(f"Nicht unterstützte MCP-Methode: {method}")
