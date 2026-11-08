# Demo site

Static pages for showing the two engines: a landing page and each engine's own click-through
demo. No server, no model key, nothing to install: host the `site/` folder on any static host
and share its link.

    python3 build.py      # copies each engine's demo/index.html in, writes site/

`site/contract-ai-demos.html` is the whole site in one file, for a host that takes single files
or runs pages inside a locked-down frame: open it and every page is there.

Rebuild after an engine's demo changes. All parties, documents and figures are invented.
