"""Comparing a quotation with its source by what it says, not how it was typeset."""

from __future__ import annotations

import re
import unicodedata

# Typography that differs between a PDF's text layer and anything retyped from it,
# in either direction. None of these changes what the document says, and treating
# them as differences would report genuine quotes as fabricated.
_EQUIVALENT_CHARS = {
    "\u2018": "'",
    "\u2019": "'",
    "\u201a": "'",
    "\u2032": "'",
    "\u201c": '"',
    "\u201d": '"',
    "\u201e": '"',
    "\u2033": '"',
    "\u2010": "-",
    "\u2011": "-",
    "\u2012": "-",
    "\u2013": "-",
    "\u2014": "-",
    "\u2212": "-",
    "\u00a0": " ",
    "\u2007": " ",
    "\u202f": " ",
    "\u200b": "",
    "\u2026": "...",
}


def normalise_for_comparison(text: str) -> str:
    """Reduce text to what it says, discarding how it was typeset.

    Whitespace is collapsed because a PDF text layer breaks lines mid-sentence and
    an agent quoting it will not reproduce the breaks. Typographic variants are
    folded because curly quotes, non-breaking spaces and dashes survive extraction
    inconsistently. Case is folded because legal drafting capitalises party names
    unpredictably.

    Deliberately NOT normalised: word order, digits, and punctuation between
    words. Those are what make the quote evidence rather than a resemblance.
    """
    text = unicodedata.normalize("NFKC", text)
    text = "".join(_EQUIVALENT_CHARS.get(char, char) for char in text)
    return re.sub(r"\s+", " ", text).strip().casefold()
