"""The library: a fixed set of folders of documents, read once into one SQLite file.

    folder/                  a matter, a client, an asset: however the documents are filed
      document.pdf           read into pages, cut into clauses, indexed by words (and meaning)

No model reads the documents at ingest and nothing is extracted from them: answers are read from
the pages and clauses when a question is asked. A document is known by a digest of its content, so
the same file filed in two folders is one document filed twice, and re-ingesting an unchanged
library reads nothing again.
"""

from __future__ import annotations

import hashlib
import json
import math
import pathlib
import re
import sqlite3
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from dataclasses import dataclass

from document_search import llm
from document_search.read import clauses_of, pages_of

READABLE = (".pdf", ".docx", ".txt", ".md")

_SCHEMA = """
create table if not exists document (id text primary key, name text, pages integer, opening text);
create table if not exists filing (document text, folder text, path text, primary key (folder, path));
create table if not exists page (document text, page integer, text text, primary key (document, page));
create table if not exists clause (id integer primary key, document text, ordinal integer,
  number text, heading text, text text, first_page integer, last_page integer, vector text);
create virtual table if not exists clause_words using fts5(number, heading, text,
  content='clause', content_rowid='id');
"""

@dataclass(frozen=True)
class Hit:
    clause: int
    document: str
    name: str
    folder: str
    path: str
    number: str
    heading: str
    text: str
    first_page: int
    last_page: int
    score: float


class Library:
    def __init__(self, db: pathlib.Path) -> None:
        self.db = db
        db.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as c:
            c.executescript(_SCHEMA)

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        c = sqlite3.connect(self.db)
        c.row_factory = sqlite3.Row
        try:
            yield c
            c.commit()
        finally:
            c.close()

    # ------------------------------------------------------------------ ingest

    def ingest(self, root: pathlib.Path, say: Callable[[str], None] = lambda _n: None) -> dict[str, int]:
        """Every readable file under ``root``, each top-level folder one folder of the library.
        A document already read (same content) is filed again, never read again."""
        files = sorted(p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in READABLE)
        with self._connect() as c:
            known = {row["id"] for row in c.execute("select id from document")}
            c.execute("delete from filing")
            fresh = 0
            for path in files:
                digest = hashlib.sha256(path.read_bytes()).hexdigest()[:16]
                relative = path.relative_to(root)
                folder = relative.parts[0] if len(relative.parts) > 1 else ""
                c.execute("insert or replace into filing values (?,?,?)", (digest, folder, str(relative)))
                if digest in known:
                    continue
                known.add(digest)
                fresh += 1
                say(f"reading {path.name}")
                pages = pages_of(path)
                c.execute("insert into document values (?,?,?,?)", (digest, path.name, len(pages), _opening(pages)))
                c.executemany("insert into page values (?,?,?)", [(digest, n, t) for n, t in enumerate(pages, 1)])
                for k, clause in enumerate(clauses_of(pages)):
                    c.execute(
                        "insert into clause (document, ordinal, number, heading, text, first_page, last_page)"
                        " values (?,?,?,?,?,?,?)",
                        (digest, k, clause.number, clause.heading, clause.text, clause.first_page, clause.last_page),
                    )
            c.execute("insert into clause_words(clause_words) values ('rebuild')")
            missing = c.execute("select id, heading, text from clause where vector is null").fetchall()
        vectors = llm.embed([f"{r['heading']}\n{r['text']}" for r in missing])
        if vectors:
            with self._connect() as c:
                c.executemany(
                    "update clause set vector = ? where id = ?",
                    [(json.dumps(v), r["id"]) for v, r in zip(vectors, missing, strict=True)],
                )
        with self._connect() as c:
            return {
                "files": len(files),
                "documents": c.execute("select count(*) from document").fetchone()[0],
                "read now": fresh,
                "clauses": c.execute("select count(*) from clause").fetchone()[0],
                "with meaning": c.execute("select count(*) from clause where vector is not null").fetchone()[0],
            }

    # ------------------------------------------------------------------ reading

    def folders(self) -> list[tuple[str, int]]:
        with self._connect() as c:
            rows = c.execute(
                "select folder, count(distinct document) n from filing group by folder order by folder"
            ).fetchall()
        return [(r["folder"], r["n"]) for r in rows]

    def documents(self, folder: str) -> list[dict]:
        """The folder's documents, each with its opening lines: what the document says it is."""
        with self._connect() as c:
            rows = c.execute(
                "select d.*, f.path from document d join filing f on f.document = d.id where f.folder = ?"
                " order by d.name",
                (folder,),
            ).fetchall()
        return [{"id": r["id"], "name": r["name"], "path": r["path"], "pages": r["pages"], "opening": r["opening"]} for r in rows]

    def search(self, query: str, folder: str = "", show: int = 6) -> list[Hit]:
        """Clauses nearest the query, by words and, where the library was embedded, by meaning;
        the two rankings fused. A provision turns on exact tokens (a defined term, a clause
        number, a figure) that a vector blurs, and on meaning that words miss."""
        with self._connect() as c:
            words = [w for w in re.findall(r"[\w£$€%.-]+", query) if len(w) > 1]
            ranked_words: list[int] = []
            if words:
                match = " OR ".join('"' + w.replace('"', "") + '"' for w in words)
                ranked_words = [
                    r[0]
                    for r in c.execute(
                        "select rowid from clause_words where clause_words match ? order by bm25(clause_words) limit 200",
                        (match,),
                    )
                ]
            ranked_meaning: list[int] = []
            vectors = llm.embed([query]) if llm.embedding_model() else []
            if vectors:
                q = vectors[0]
                qn = math.sqrt(sum(x * x for x in q)) or 1.0
                scored = []
                for r in c.execute("select id, vector from clause where vector is not null"):
                    v = json.loads(r["vector"])
                    vn = math.sqrt(sum(x * x for x in v)) or 1.0
                    scored.append((sum(a * b for a, b in zip(q, v, strict=False)) / (qn * vn), r["id"]))
                ranked_meaning = [i for _, i in sorted(scored, reverse=True)[:200]]
            fused: dict[int, float] = {}
            for ranking in (ranked_words, ranked_meaning):
                for rank, cid in enumerate(ranking):
                    fused[cid] = fused.get(cid, 0.0) + 1.0 / (60 + rank)
            hits: list[Hit] = []
            for cid, score in sorted(fused.items(), key=lambda x: -x[1]):
                r = c.execute(
                    "select c.*, d.name, f.folder, f.path from clause c join document d on d.id = c.document"
                    " join filing f on f.document = c.document where c.id = ?"
                    + (" and f.folder = ?" if folder else ""),
                    (cid, folder) if folder else (cid,),
                ).fetchone()
                if r is None:
                    continue
                hits.append(
                    Hit(cid, r["document"], r["name"], r["folder"], r["path"], r["number"], r["heading"],
                        r["text"], r["first_page"], r["last_page"], score)
                )
                if len(hits) >= show:
                    break
        return hits

    def clause(self, document: str, naming: str) -> list[dict]:
        """A document's clause by its number ("5.3", "Schedule 2") or its heading, with every
        sub-clause under a number asked for whole."""
        doc = self._document(document)
        if doc is None:
            return []
        naming = naming.strip().rstrip(".")
        with self._connect() as c:
            rows = c.execute(
                "select * from clause where document = ? order by ordinal", (doc["id"],)
            ).fetchall()
        under = [r for r in rows if r["number"] and (r["number"] == naming or r["number"].startswith(naming + "."))]
        if not under:
            under = [r for r in rows if naming.casefold() in (r["heading"] or "").casefold()][:1]
        return [dict(r) | {"name": doc["name"]} for r in under]

    def pages(self, document: str, first: int, last: int) -> list[tuple[int, str]]:
        doc = self._document(document)
        if doc is None:
            return []
        with self._connect() as c:
            return [
                (r["page"], r["text"])
                for r in c.execute(
                    "select page, text from page where document = ? and page between ? and ? order by page",
                    (doc["id"], first, min(last, first + 4)),
                )
            ]

    def _document(self, named: str) -> sqlite3.Row | None:
        """A document by its id, its file name or the start of its file name."""
        with self._connect() as c:
            row = c.execute("select * from document where id = ? or name = ?", (named, named)).fetchone()
            if row is None:
                found = c.execute(
                    "select * from document where name like ? escape '\\'",
                    (named.replace("%", "\\%").replace("_", "\\_") + "%",),
                ).fetchall()
                row = found[0] if len(found) == 1 else None
        return row

    def path_of(self, document: str) -> str:
        with self._connect() as c:
            row = c.execute("select path from filing where document = ? order by folder", (document,)).fetchone()
        return row["path"] if row else ""


def _opening(pages: list[str], most: int = 240) -> str:
    """The document's first lines, as it states them: its title, its date, its parties."""
    lines = [line.strip() for line in (pages[0] if pages else "").split("\n") if line.strip()]
    return " / ".join(lines[:4])[:most]
