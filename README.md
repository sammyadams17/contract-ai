# Contract AI

Two agentic engines for legal work on contracts, each usable on its own.

## Live demo

**[Open the click-through demo →](https://sammyadams17.github.io/contract-ai/demo-site/site/contract-ai-demos.html)**

It replays real runs of both engines on invented documents. No setup, no model key, and it
works offline. To run it locally, open `demo-site/site/contract-ai-demos.html` in a browser.

- **Contract drafting engine:** an agreed term sheet and the client's template (or a precedent
  from an earlier deal) go in. A first-draft contract comes out as a Word redline, with every
  change tracked and every open point raised as a comment that gives its reason.
- **Document search engine:** ask questions of a library of signed contracts. The agent reads
  every document that could change the answer (amendments, novations, terminations, side
  letters) and cites the page.

## What's here

| Folder | What it is |
|---|---|
| [`contract-drafting-engine/`](contract-drafting-engine/) | Term sheet + template or precedent in, Word redline out. Tens of agents in parallel; a full 150-page contract in about 20–25 minutes |
| [`document-search-engine/`](document-search-engine/) | Questions over a fixed contract library, answered with page citations. 13 of 13 checked questions right |
| [`demo-site/`](demo-site/) | The click-through demo above, built from each engine's own replay |

Each engine has its own README with architecture, tests, example data and setup.

All parties, documents and figures in the examples are invented. Every output is a first draft or
an answer for a lawyer to check, never legal advice.
