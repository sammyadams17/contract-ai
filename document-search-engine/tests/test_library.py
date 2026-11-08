"""The library read once: folders, documents by their opening lines, search, clauses by number."""

from __future__ import annotations


def test_each_document_is_listed_with_what_it_says_it_is(library):
    acme = {d["name"]: d for d in library.documents("Acme Services")}
    assert set(acme) == {"Services Agreement 2024.pdf", "Termination Agreement.pdf", "New Services Agreement.pdf", "Deed of Novation.pdf"}
    assert acme["Termination Agreement.pdf"]["opening"].startswith("TERMINATION AGREEMENT / THIS AGREEMENT is dated 1 March 2025")


def test_search_finds_a_clause_by_its_words_in_every_document(library):
    hits = library.search("fee per month", folder="Acme Services")
    assert {h.name for h in hits} >= {"New Services Agreement.pdf", "Services Agreement 2024.pdf"}


def test_a_sub_clause_is_found_under_its_clauses_heading(library):
    hits = library.search("deposit", folder="Harbour Street Lease")
    assert any(h.number == "4.1" and h.first_page == 2 for h in hits)


def test_a_clause_is_read_whole_by_its_number(library):
    rows = library.clause("Lease of Unit 4 Harbour Street.pdf", "5")
    assert [r["number"] for r in rows] == ["5", "5.1"]
    assert "six (6) months" in rows[1]["text"]


def test_an_unchanged_library_is_not_read_again(library):
    from conftest import LIBRARY

    assert library.ingest(LIBRARY)["read now"] == 0
