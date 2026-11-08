"""Build the demo site: a landing page and each
engine's own demo. Every page is a static file; the folder can be hosted anywhere as it is.

    python3 build.py

Copies each engine's built demo (demo/index.html in its repository) beside the pages here.
"""

from __future__ import annotations

import pathlib
import shutil

HERE = pathlib.Path(__file__).resolve().parent
ASSETS = HERE.parent
SITE = HERE / "site"
ENGINES = {
    "drafting.html": ASSETS / "contract-drafting-engine" / "demo" / "index.html",
    "search.html": ASSETS / "document-search-engine" / "demo" / "index.html",
}

def main() -> None:
    SITE.mkdir(exist_ok=True)
    for name, source in ENGINES.items():
        shutil.copyfile(source, SITE / name)
    for page in ("index.html",):
        shutil.copyfile(HERE / "pages" / page, SITE / page)
    one_file(SITE / "contract-ai-demos.html")
    print("built", sorted(p.name for p in SITE.iterdir()))


def one_file(out: pathlib.Path) -> None:
    """Every page in one HTML file with no frames, for a host that runs pages in a locked-down
    frame: such a frame lets a page run its scripts but not reach into frames of its
    own, open pop-ups or start downloads. Each page becomes a view isolated in a shadow root, so
    the pages' styles never meet, and its script runs against that root. A link to another page
    switches the view in memory, never by changing the URL."""
    import json
    import re

    def parts(page: str) -> dict:
        css = "\n".join(re.findall(r"<style>(.*?)</style>", page, re.S))
        css = css.replace(":root", ":host")
        css = re.sub(r"(^|[}\s,])body(\s*[{,])", r"\1.page-body\2", css)
        body = re.search(r"<body>(.*)</body>", page, re.S).group(1)
        code = "\n".join(re.findall(r"<script>(.*?)</script>", body, re.S))
        body = re.sub(r"<script>.*?</script>", "", body, flags=re.S)
        return {"css": css, "body": body, "code": code}

    pages = {
        "index": (SITE / "index.html").read_text(),
        "drafting": (SITE / "drafting.html").read_text(),
        "search": (SITE / "search.html").read_text(),
    }
    views = json.dumps({k: parts(v) for k, v in pages.items()}, ensure_ascii=False).replace("</", "<\\/")
    out.write_text(
        """<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Contract AI demos</title>
<style>body{margin:0;background:#f6f8fa} .view{display:none} .view.on{display:block}
#bar{position:sticky;top:0;z-index:50;display:none;gap:10px;align-items:center;padding:8px 16px;background:#14202e;color:#fff;
font:14px -apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
#bar button{background:#22344a;color:#fff;border:0;border-radius:6px;padding:6px 12px;font:inherit;cursor:pointer}
#bar button:hover{background:#2d4560} #bar span{color:#9fb2c6}</style></head>
<body><div id="bar"><button id="back">← Back</button><button id="home">All demos</button><span id="where"></span></div><div id="app"></div>
<script id="views" type="application/json">"""
        + views
        + """</script>
<script>
const VIEWS = JSON.parse(document.getElementById("views").textContent);
function mount(host, name) {
  const view = VIEWS[name], root = host.attachShadow({ mode: "open" });
  root.innerHTML = `<style>${view.css}</style><div class="page-body">${view.body}</div>`;
  // a link to another page switches the view
  root.addEventListener("click", (e) => {
    const a = e.target.closest && e.target.closest('a[href$=".html"]');
    if (a) { e.preventDefault(); go(a.getAttribute("href").replace(/\\.html$/, "")); }
  });
  const doc = {
    querySelector: (s) => root.querySelector(s), querySelectorAll: (s) => root.querySelectorAll(s),
    getElementById: (id) => root.getElementById(id), createElement: (t) => document.createElement(t),
    addEventListener: (...a) => root.addEventListener(...a), body: root.querySelector(".page-body"),
  };
  new Function("document", view.code)(doc);
}
window.__mount = mount;
window.__go = (name) => go(name);
const app = document.getElementById("app"), shown = {};
// views switch in memory, and Back walks the page's own trail of them
const trail = [], NAMES = { index: "All demos", drafting: "Contract drafting engine", search: "Document search engine" };
function go(asked, going_back) {
  const name = VIEWS[asked] ? asked : "index";
  if (!going_back && trail[trail.length - 1] !== name) {
    trail.push(name);

  }
  document.getElementById("bar").style.display = name === "index" ? "none" : "flex";
  document.getElementById("where").textContent = NAMES[name] || "";
  if (!shown[name]) {
    const host = document.createElement("div"); host.className = "view";
    // what a page pins to the top of the screen sits below the bar
    host.style.setProperty("--top", document.getElementById("bar").offsetHeight + "px");
    app.appendChild(host); mount(host, name); shown[name] = host;
  }
  Object.entries(shown).forEach(([n, h]) => h.classList.toggle("on", n === name));
  window.scrollTo(0, 0);
}
function back() { if (trail.length > 1) { trail.pop(); go(trail[trail.length - 1], true); } else go("index", true); }
// never the browser's history: a host that runs the page in its own frame takes any
// history step for the page replacing itself
document.getElementById("back").onclick = () => back();
document.getElementById("home").onclick = () => go("index");
trail.push("index"); go("index", true);
</script></body></html>"""
    )

if __name__ == "__main__":
    main()
