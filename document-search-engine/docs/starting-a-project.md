# Starting a project with this engine

Use it when a client wants questions answered from a fixed set of signed documents, with every
answer citing the page: a contract library, a deal's documents, a portfolio's leases.

## Week 1: get it answering

1. **Collect the documents**, one folder per matter, client or asset, as the client files them.
   PDFs give page links. Scans need OCR first.
2. **Ingest:** `document-search ingest LIBRARY --db library.db`. It reads only new or changed files.
3. **Check the graph:** for each folder, see that the documents marked as ended really are ended
   (`documents` in a Python session, or ask "which documents in folder X still apply?"). Wrong
   ones usually mean a document names an earlier one by a date or title that isn't there.
4. **Write 20–30 checked questions** with a client lawyer: each with the right value and the page.
   Include "what is it now" questions where an amendment changed something, one about a terminated
   document, one with no answer, and one across every folder.
5. **Run `eval`** and read every wrong answer. Fix the brief (`src/document_search/brief.md`) for a
   wrong habit, or the reader (`read.py`) for a document cut into clauses badly. Never fix an
   answer by hand-editing the library.

## Week 2 onwards: make it theirs

- **Brief:** add the client's own words and rules (what they call things, which documents to
  trust first) to the brief, in plain sentences, with the reason.
- **Search:** set `EMBEDDING_MODEL` so search works by meaning as well as words.
- **Front end and access:** the engine is a library and a command line. Put an API or a chat page
  in front of `agent.answer`, behind the client's login. Store the library file where only
  permitted users can reach it.
- **Size:** the local file and its search are fine for thousands of documents. Beyond that, move
  the clause search to a search service and keep the rest.

## Keep it honest

- The questions file is the test. Re-run it after every change, and report counts of right and
  wrong answers a lawyer can look at, never a score the engine gives itself.
- A citation that doesn't open the right page is a defect, even when the answer is right.
