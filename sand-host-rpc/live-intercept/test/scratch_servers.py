#!/usr/bin/env python3
"""scratch_servers.py — faux serveurs pour tester l'interception SANS toucher host-main.
  plain <port>           : serveur HTTP/1.1 EN CLAIR (simule host-main sur 1341)
  tls <port> <cert> <key>: serveur HTTPS (simule un backend)
Chaque réponse inclut VOLONTAIREMENT des secrets factices (JWT, Bearer, e-mail,
clé privée, token GH) pour prouver que la rédaction les supprime du log.
"""
import http.server, ssl, sys, json

FAKE_JWT = "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJ0ZXN0LXNlY3JldC1wYXlsb2FkLXh4eCJ9.Zm9vYmFyYmF6cXV4c2VjcmV0c2ln"
SECRETS_BODY = {
    "note": "reponse scratch",
    "token": FAKE_JWT,
    "contact": "victime@example.com",
    "gh": "ghp_"+"abcdef0123456789ABCDEFghij0123456789"  # fixture factice, scindée pour éviter push-protection,
}

class H(http.server.BaseHTTPRequestHandler):
    def _h(self):
        self.send_header("Authorization", f"Bearer {FAKE_JWT}")
        self.send_header("Set-Cookie", "session=supersecretcookievalue123456")
        self.send_header("Content-Type", "application/json")
    def do_POST(self): self._reply()
    def do_GET(self): self._reply()
    def _reply(self):
        n = int(self.headers.get("content-length", 0) or 0)
        _ = self.rfile.read(n) if n else b""
        body = json.dumps(SECRETS_BODY).encode()
        self.send_response(200); self._h()
        self.send_header("Content-Length", str(len(body))); self.end_headers()
        self.wfile.write(body)
    def log_message(self, *a): pass

def main():
    mode, port = sys.argv[1], int(sys.argv[2])
    srv = http.server.HTTPServer(("127.0.0.1", port), H)
    if mode == "tls":
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(sys.argv[3], sys.argv[4])
        srv.socket = ctx.wrap_socket(srv.socket, server_side=True)
    srv.serve_forever()

if __name__ == "__main__":
    main()
