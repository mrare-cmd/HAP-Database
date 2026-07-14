#!/usr/bin/env python3
"""
HAP DATABASE - local helper.
Starts a tiny local web server and opens the HAP Database page in your browser.
The page's "Update" button pulls the current month's HUD data and refreshes.
Nothing leaves your computer; this only runs while the window is open.
"""
import os, sys, json, threading, datetime, webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import hap_pipeline as hp
import hap_writer as hw
import hap_html as hh

PORT = 8765
STATE = os.path.join(HERE, "hap_state.json")     # cached data payload (JSON)
PAGE  = os.path.join(HERE, "_hap_page.html")      # served page (no embedded data)
_lock = threading.Lock()
_busy = False

def build_page():
    hh.write_html(None, PAGE, asof="", embed=False)

def read_state_bytes():
    if os.path.exists(STATE):
        with open(STATE, "rb") as f:
            return f.read()
    return json.dumps({"cols": [], "types": [], "rows": [], "asof": "(no data yet)"}).encode()

def do_update():
    msgs = []
    master = hp.fetch_and_build(progress=lambda m: msgs.append(m))
    asof = datetime.date.today().isoformat()
    payload = hh.payload_json(master, asof)
    with open(STATE, "w", encoding="utf-8") as f:
        f.write(payload)
    # also drop a plain spreadsheet copy
    try:
        hw.write_master(master, os.path.join(HERE, f"HAP_Database_{asof}.xlsx"))
    except Exception:
        pass
    return asof, len(master)

class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def _send(self, code, body, ctype="application/json"):
        if isinstance(body, str): body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)
    def do_GET(self):
        if self.path in ("/", "/index.html"):
            with open(PAGE, "rb") as f: self._send(200, f.read(), "text/html; charset=utf-8")
        elif self.path.startswith("/data"):
            self._send(200, read_state_bytes())
        else:
            self._send(404, "not found", "text/plain")
    def do_POST(self):
        global _busy
        if self.path.startswith("/update"):
            with _lock:
                if _busy:
                    self._send(200, json.dumps({"ok": False, "error": "an update is already running"})); return
                _busy = True
            try:
                asof, rows = do_update()
                self._send(200, json.dumps({"ok": True, "asof": asof, "rows": rows}))
            except Exception as e:
                self._send(200, json.dumps({"ok": False, "error": str(e)}))
            finally:
                _busy = False
        else:
            self._send(404, "not found", "text/plain")

def main():
    build_page()
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), H)
    url = f"http://127.0.0.1:{PORT}/"
    print("HAP Database helper running.")
    print("Your database is open in the browser at", url)
    print("Leave this window open while you use it. Close it when you're done.")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
