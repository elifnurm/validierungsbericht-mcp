import json
from http.server import BaseHTTPRequestHandler
from server.mcp_server import handle_mcp

class handler(BaseHTTPRequestHandler):
    def _headers(self, status=200, content_type="application/json"):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "content-type, accept, mcp-session-id")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.end_headers()

    def do_OPTIONS(self):
        self._headers(204)

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            req = json.loads(self.rfile.read(length) or b"{}")
            result = handle_mcp(req)
            if "id" not in req or result is None:
                self._headers(202)
                self.wfile.write(b"{}")
                return
            body = json.dumps({"jsonrpc":"2.0","id":req["id"],"result":result}, ensure_ascii=False).encode("utf-8")
            self._headers(200)
            self.wfile.write(body)
        except Exception as exc:
            rid = None
            try: rid = req.get("id")
            except Exception: pass
            body = json.dumps({"jsonrpc":"2.0","id":rid,"error":{"code":-32000,"message":str(exc)}}, ensure_ascii=False).encode("utf-8")
            self._headers(200)
            self.wfile.write(body)
