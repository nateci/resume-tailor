#!/usr/bin/env python3
"""
Small local HTTP server so the tracker HTMLs' "Update" button can actually
do something. A static file opened via file:// has no way to run a local
script (browser sandboxing forbids it entirely) -- this serves output/ over
http://localhost instead, and the button's fetch('/update') call hits a
POST endpoint here that re-runs the Outlook scan + Sheets import in-process
before the page reloads.

Run this, then open http://localhost:8765/applied_tracker.html (or
intern_applied_tracker.html / dashboard.html / intern_dashboard.html)
instead of double-clicking the file directly. Ctrl+C to stop.

/update always refreshes both pipelines (new-grad + intern) in one shot,
since scan_email_status.py and import_tracker_sheets.py already process
both every run regardless of which tracker's button was clicked.
"""
import http.server
import os
import subprocess
import sys

PORT = 8765
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(os.path.dirname(SCRIPTS_DIR), "output")


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=OUTPUT_DIR, **kwargs)

    def do_POST(self):
        if self.path != "/update":
            self.send_error(404)
            return
        ok, log = run_update()
        body = log.encode("utf-8")
        self.send_response(200 if ok else 500)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))


def run_update():
    steps = [
        [sys.executable, os.path.join(SCRIPTS_DIR, "scan_email_status.py")],
        [sys.executable, os.path.join(SCRIPTS_DIR, "import_tracker_sheets.py")],
    ]
    log = []
    ok = True
    for cmd in steps:
        result = subprocess.run(cmd, capture_output=True, text=True)
        log.append(f"$ {os.path.basename(cmd[1])}\n{result.stdout}{result.stderr}")
        if result.returncode != 0:
            ok = False
    return ok, "\n".join(log)


def main():
    server = http.server.ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"Serving {OUTPUT_DIR} at http://localhost:{PORT}/ (Ctrl+C to stop)", file=sys.stderr)
    print(f"Open http://localhost:{PORT}/applied_tracker.html or "
          f"http://localhost:{PORT}/intern_applied_tracker.html", file=sys.stderr)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
