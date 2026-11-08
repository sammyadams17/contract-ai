"""The loop: a question in, an answer with its citations out.

The model chooses what to ask of the library; it never chooses what the documents say. Every tool
answers from the stored pages and clauses and returns its own evidence, with the link that opens
the page, and the brief requires every fact to cite one.
"""

from __future__ import annotations

import json
import os
import pathlib
import re
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from urllib.parse import quote, unquote

from document_search import llm
from document_search.library import Hit, Library

BRIEF = (pathlib.Path(__file__).with_name("brief.md")).read_text()
LINK_BASE = os.environ.get("SEARCH_LINK_BASE", "library/")
_ROUNDS = 16
_CITED = re.compile(r"\[([^\]]+)\]\(([^)\s#]+)(?:#page=(\d+))?\)")


@dataclass
class Answer:
    text: str
    citations: list[dict] = field(default_factory=list)
    steps: list[str] = field(default_factory=list)


def _link(path: str, page: int) -> str:
    """A link that opens the page: the path encoded, since folder names have spaces."""
    return f"{LINK_BASE}{quote(path)}#page={page}"


def _tool(name: str, description: str, properties: dict, required: list[str]) -> dict:
    return {
        "type": "function",
        "function": {
            "name": name,
            "description": description,
            "parameters": {"type": "object", "properties": properties, "required": required},
        },
    }


_S = {"type": "string"}
TOOLS = [
    _tool("folders", "Every folder of the library and how many documents each holds.", {}, []),
    _tool(
        "documents",
        "A folder's documents: each one's name, pages and opening lines (its title, date and"
        " parties as it states them), with the link that opens it.",
        {"folder": _S},
        ["folder"],
    ),
    _tool(
        "find_clauses",
        "Search clauses by words and meaning, in one folder or the whole library. Each hit gives"
        " its link.",
        {"query": _S, "folder": _S},
        ["query"],
    ),
    _tool(
        "read_clause",
        "A clause of one document by its number ('5.3', 'Schedule 2') or its heading, whole, with"
        " its sub-clauses. `document` is the document's name as the tools gave it.",
        {"document": _S, "naming": _S},
        ["document", "naming"],
    ),
    _tool(
        "read_pages",
        "The text of up to five pages of one document, for what no clause holds (a cover, a"
        " signature page, a schedule table).",
        {"document": _S, "first": {"type": "integer"}, "last": {"type": "integer"}},
        ["document", "first"],
    ),
]
DEEP = _tool(
    "deep_search",
    "Ask one question of every folder (or the folders named), each folder read in full by itself;"
    " for questions across the library ('which agreements are governed by German law?'). Slower:"
    " use it only for a question about many folders.",
    {"question": _S, "folders": {"type": "array", "items": _S}},
    ["question"],
)


class Session:
    """One answer's tools."""

    def __init__(self, library: Library, say: Callable[[str], None]) -> None:
        self.library = library
        self.say = say

    def _shown(self, hit: Hit | dict, path: str) -> dict:
        first, last = hit["first_page"] if isinstance(hit, dict) else hit.first_page, (
            hit["last_page"] if isinstance(hit, dict) else hit.last_page
        )
        return {"link": _link(path, first), "pages": f"{first}" if first == last else f"{first}-{last}"}

    def folders(self) -> str:
        return json.dumps([{"folder": f, "documents": n} for f, n in self.library.folders()])

    def documents(self, folder: str) -> str:
        docs = self.library.documents(folder)
        if not docs:
            return f"no folder named {folder!r}; the folders are {[f for f, _ in self.library.folders()]}"
        return json.dumps(
            [{k: d[k] for k in ("name", "pages", "opening")} | {"link": _link(d["path"], 1)} for d in docs]
        )

    def find_clauses(self, query: str, folder: str = "") -> str:
        hits = self.library.search(query, folder)
        if not hits:
            return "no clause found for those words" + (f" in {folder}" if folder else "")
        return json.dumps(
            [
                {"document": h.name, "folder": h.folder, "clause": h.number, "heading": h.heading,
                 "text": h.text[:1500]} | self._shown(h, h.path)
                for h in hits
            ]
        )

    def read_clause(self, document: str, naming: str) -> str:
        rows = self.library.clause(document, naming)
        if not rows:
            return f"no clause {naming!r} in {document!r}; use find_clauses or read_pages"
        path = self.library.path_of(rows[0]["document"])
        return json.dumps(
            [{"document": r["name"], "clause": r["number"], "heading": r["heading"], "text": r["text"]}
             | self._shown(r, path) for r in rows]
        )

    def read_pages(self, document: str, first: int, last: int | None = None) -> str:
        pages = self.library.pages(document, first, last or first)
        if not pages:
            return f"no pages {first}-{last or first} in {document!r}"
        row = self.library._document(document)
        path = self.library.path_of(row["id"]) if row else ""
        return json.dumps([{"page": n, "link": _link(path, n), "text": t} for n, t in pages])

    def deep_search(self, question: str, folders: list[str] | None = None) -> str:
        names = folders or [f for f, _ in self.library.folders()]
        self.say(f"reading {len(names)} folders, each by itself")

        def one(folder: str) -> tuple[str, Answer]:
            return folder, answer(
                f"In folder {folder!r} only: {question}\nStart your answer with yes, no or not stated"
                " where the question allows, then the answer in two sentences with its citation.",
                self.library,
                say=lambda note, f=folder: self.say(f"{f}: {note}"),
                folder=folder,
            )

        with ThreadPoolExecutor(max_workers=4) as pool:
            done = list(pool.map(one, names))
        # the list is written here, not by the model: handed many folders' answers, a model
        # writes a narrative and never gives the list that was asked for
        groups: dict[str, list[str]] = {"Yes": [], "No": [], "Not stated": []}
        for folder, a in done:
            first = a.text.strip().split(None, 1)[0].strip(".,:;*").casefold() if a.text.strip() else ""
            key = "Yes" if first == "yes" else "No" if first == "no" else "Not stated"
            line = " ".join(a.text.split("**Sources**")[0].split("### Sources")[0].split())
            groups[key].append(f"- {folder}: {line}")
        listing = "\n".join(
            f"{name} ({len(lines)}):\n" + "\n".join(lines) for name, lines in groups.items() if lines
        )
        return f"Read {len(done)} folders, each by itself.\n\n{listing}\n\nGive the reader this list as written, grouped as it is, with its links; add at most one line before it."

    def run(self, name: str, args: dict) -> str:
        method = getattr(self, name, None)
        if method is None or name.startswith("_") or name == "run":
            return f"unknown tool {name!r}"
        try:
            return method(**args)
        except Exception as error:  # a tool failure is an observation, said as words
            return f"that tool failed: {type(error).__name__}: {error}"


def citations(text: str) -> list[dict]:
    """The pages an answer links to, read out of the answer itself."""
    out = []
    for label, url, page in _CITED.findall(text):
        path = unquote(url.removeprefix(LINK_BASE))
        entry = {"label": label, "path": path, "page": int(page) if page else 0, "link": f"{url}#page={page}" if page else url}
        if entry not in out:
            out.append(entry)
    return out


def answer(
    question: str,
    library: Library,
    history: list[dict] | None = None,
    say: Callable[[str], None] = lambda _n: None,
    folder: str = "",
) -> Answer:
    """One question answered. ``folder`` confines a deep search's reader to its folder."""
    session = Session(library, say)
    tools = TOOLS if folder else [*TOOLS, DEEP]
    shelf = "\n".join(f"- {f} ({n} documents)" for f, n in library.folders())
    messages: list[dict] = [
        {"role": "system", "content": BRIEF + f"\n\nTHE LIBRARY'S FOLDERS:\n{shelf}"},
        *(history or []),
        {"role": "user", "content": question},
    ]
    steps: list[str] = []
    for _ in range(_ROUNDS):
        message = llm.chat(messages, tools, say)
        calls = message.tool_calls or []
        if not calls:
            text = message.content or ""
            return Answer(text, citations(text), steps)
        messages.append(
            {
                "role": "assistant",
                "content": message.content or "",
                "tool_calls": [
                    {"id": c.id, "type": "function", "function": {"name": c.function.name, "arguments": c.function.arguments}}
                    for c in calls
                ],
            }
        )
        for c in calls:
            args = json.loads(c.function.arguments or "{}")
            step = f"{c.function.name}({', '.join(f'{k}={v!r}' for k, v in args.items())})"
            steps.append(step)
            say(step)
            messages.append({"role": "tool", "tool_call_id": c.id, "content": session.run(c.function.name, args)})
    return Answer("I could not finish this answer in the time allowed; ask again more narrowly.", [], steps)
