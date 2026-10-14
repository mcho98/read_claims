"""A small local stand-in for the real claim site, for testing the extension offline.

It copies the behaviours seen in the saved real page (not its look):
  * link-style buttons (<a role="button">), the Add claim label split by an accesskey span
  * changing the date or the service makes the form refresh itself after a delay and the
    refresh wipes hours and cost (the "auto-submit" race)
  * Add claim does a full page POST + reload (the page script is destroyed)
  * client-side validation (date window, hours, cost) shows pink fields and messages
  * the claims table ends with a bare total such as "$416.26"
  * Predetermine raises a native alert()
Run:  python3 test/real_like_site.py [port]      (prints the address)
"""
import datetime as dt
import html
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs

REFRESH_DELAY_MS = 900   # how long the simulated server refresh takes
RELOAD_DELAY_S = 0.6     # how long the simulated POST takes
ROWS = []                # (date, service, hours, cost)
LOCK = threading.Lock()
SERVICES = ["Personal care", "Respite", "Homemaking"]


def page():
    today = dt.date.today()
    lo, hi = (today - dt.timedelta(days=60)).isoformat(), today.isoformat()
    with LOCK:
        rows = list(ROWS)
    total = sum(float(r[3]) for r in rows)
    body_rows = "".join(
        f"<tr><td>{html.escape(d)}</td><td>{html.escape(s)}</td><td>{html.escape(h)}</td><td>{html.escape(c)}</td>"
        f"<td><a href='#'>Modify</a></td><td><a href='#'>Remove</a></td></tr>" for d, s, h, c in rows)
    table = ""
    if rows:
        table = (f"<table border=1><thead><tr><th>Service date (YYYY-MM-DD)</th><th>Service</th>"
                 f"<th>Number of hours</th><th>Total cost ($)</th><th></th><th></th></tr></thead><tbody>{body_rows}</tbody>"
                 f"<tfoot><tr><td></td><td></td><td></td><td><b>${total:.2f}</b></td><td></td><td></td></tr></tfoot></table>")
    options = "".join(f"<option value='{i}'>{s}</option>" for i, s in enumerate(SERVICES))
    return f"""<!doctype html><html><head><meta charset=utf-8><title>Real-like claim site</title></head><body>
<h3>Claim details</h3>
<form id="f" method="post" action="/add" onsubmit="return false">
 <table><tr><td><label for="p:serviceDate::content">Service date (YYYY-MM-DD)</label></td>
  <td><input id="p:serviceDate::content" name="serviceDate" size="10" maxlength="10"><span id="p:serviceDateErrorMsg"></span></td></tr>
 <tr><td><label for="p:sproducts::content">Service</label></td>
  <td><select id="p:sproducts::content" name="service"><option value="">- Choose one -</option>{options}</select>
  <span id="p:productsErrorMsg"></span></td></tr></table>
 <div id="refreshable"><table>
  <tr><td><label for="p:cdcHours::content">Number of hours</label></td>
   <td><input id="p:cdcHours::content" name="hours" size="6" value="0"><span id="p:cdcHoursErrorMsg"></span></td></tr>
  <tr><td><label for="p:price::content">Total cost ($)</label></td>
   <td><input id="p:price::content" name="price" size="6"><span id="p:PriceErrorMsg"></span></td></tr></table></div>
 <p><a id="add" href="#" role="button" onclick="return addClaim()">Add cl<span>a</span>im</a>
    <a href="#" role="button" onclick="return false">Cancel</a>
    <a id="pre" href="#" role="button" onclick="return pre()"><span>P</span>redetermine</a></p>
</form>
{table}
<script>
var lo="{lo}", hi="{hi}";
function $(id){{return document.getElementById(id);}}
function bad(el,msgId,msg){{ el.style.backgroundColor="pink"; $(msgId).textContent=msg; return false; }}
function good(el,msgId){{ el.style.backgroundColor=""; $(msgId).textContent=""; return true; }}
// "auto-submit": after the date or service changes the server re-renders the hours/cost block
var pending=0;
function refresh(){{ pending++; setTimeout(function(){{
  $("refreshable").innerHTML = '<table><tr><td><label for="p:cdcHours::content">Number of hours</label></td>' +
   '<td><input id="p:cdcHours::content" name="hours" size="6" value="0"><span id="p:cdcHoursErrorMsg"></span></td></tr>' +
   '<tr><td><label for="p:price::content">Total cost ($)</label></td>' +
   '<td><input id="p:price::content" name="price" size="6"><span id="p:PriceErrorMsg"></span></td></tr></table>';
  pending--; }}, {REFRESH_DELAY_MS}); }}
$("p:serviceDate::content").addEventListener("change", refresh);
$("p:sproducts::content").addEventListener("change", refresh);
function addClaim(){{
  var d=$("p:serviceDate::content"), s=$("p:sproducts::content"), h=$("p:cdcHours::content"), c=$("p:price::content");
  var ok=true;
  if(!/^\\d{{4}}-\\d{{2}}-\\d{{2}}$/.test(d.value)) ok=bad(d,"p:serviceDateErrorMsg","Invalid date");
  else if(d.value<lo||d.value>hi) ok=bad(d,"p:serviceDateErrorMsg","Date must be between "+lo+" and "+hi); else good(d,"p:serviceDateErrorMsg");
  if(!s.value) ok=bad(s,"p:productsErrorMsg","Please select a service."); else good(s,"p:productsErrorMsg");
  if(!/^\\d*(\\.\\d{{0,2}})?$/.test(h.value)||!(+h.value>0&&+h.value<99)) ok=bad(h,"p:cdcHoursErrorMsg","Please enter the number of hours."); else good(h,"p:cdcHoursErrorMsg");
  if(!/^\\d+(\\.\\d{{1,2}})?$/.test(c.value)) ok=bad(c,"p:PriceErrorMsg","Please enter the total cost of the service."); else good(c,"p:PriceErrorMsg");
  if(!ok) return false;
  var f=$("f"); f.onsubmit=null; f.submit(); return false; }}
function pre(){{ alert("Your claim is being processed."); return false; }}
</script></body></html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):  # quiet
        pass

    def do_GET(self):
        data = page().encode()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        f = parse_qs(self.rfile.read(n).decode())
        time.sleep(RELOAD_DELAY_S)
        service = SERVICES[int(f.get("service", ["0"])[0])]
        with LOCK:
            ROWS.append((f["serviceDate"][0], service, f["hours"][0], f["price"][0]))
        self.send_response(303)
        self.send_header("Location", "/")
        self.end_headers()


def serve(port=0):
    server = HTTPServer(("127.0.0.1", port), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    s = serve(port)
    print(f"http://127.0.0.1:{s.server_address[1]}/")
    threading.Event().wait()
