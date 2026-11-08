"""PDF text extraction for term sheets, with an OCR fallback for scanned signature copies.

Term sheets commonly arrive as PDFs, not DOCX - they are signed execution
versions - so without this the pipeline cannot ingest a typical matter. Most
extract as text normally, and some are flattened scans returning next to no
characters from the whole document. Those are not corrupt; they are images of pages, typically because the file went through a
DocuSign/print-scan cycle.

The fallback therefore rasterises and OCRs rather than treating an empty
extraction as an empty document - silently ingesting 0 characters would hand the
reader agents nothing and produce a draft with every fact missing,
which looks like a drafting failure rather than an ingestion one.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

# Judged PER PAGE, not in total. An absolute threshold is not safe here: the
# only text layer of a long flattened scan can be a DocuSign envelope stamp repeated
# once per page. That totals thousands of characters - enough to
# clear any plausible absolute floor - while containing no document content
# whatsoever, so it was accepted as text and OCR never ran.
#
# Per page the distinction is unambiguous: genuine text layers run to thousands of
# chars/page, scans to tens or none. Nothing sits near this threshold.
#
# Kept as a number, deliberately, unlike the other thresholds in this codebase.
# Those were proxies - length standing in for distinctiveness or for bundling -
# and each had an exact property underneath it to measure instead. Here density
# IS the property: the question is literally "does this page carry text", and
# characters per page is the direct measure of it, with an order of magnitude
# between the two populations rather than a fitted boundary.
_MIN_CHARS_PER_PAGE = 250
_ASSUMED_PAGES_IF_UNKNOWN = 1
_OCR_DPI = 200


class PdfExtractionError(RuntimeError):
    pass


def page_count(path: str) -> int:
    try:
        from pypdf import PdfReader

        return max(1, len(PdfReader(path).pages))
    except Exception:  # noqa: BLE001 - page count is only used to scale a threshold
        return _ASSUMED_PAGES_IF_UNKNOWN


def has_usable_text_layer(text: str, pages: int) -> bool:
    return (len(_content_text(text)) / max(1, pages)) >= _MIN_CHARS_PER_PAGE


def _content_text(text: str) -> str:
    """Strip per-page boilerplate that inflates a scan's apparent text layer.

    The DocuSign envelope stamp is emitted once per page on signed copies and is
    the sole content of some scans, so counting it as document text is what let a
    long image-only term sheet pass as readable.
    """
    lines = [
        line
        for line in text.splitlines()
        if line.strip() and not line.strip().lower().startswith("docusign envelope id")
    ]
    return "\n".join(lines).strip()


def extract_pdf_text(path: str, allow_ocr: bool = True) -> str:
    """Extract text from a PDF, falling back to OCR when there is no real text layer."""
    pages = page_count(path)
    text = _extract_text_layer(path)
    if has_usable_text_layer(text, pages):
        return text

    found = len(_content_text(text))
    if not allow_ocr:
        raise PdfExtractionError(
            f"{Path(path).name} has no usable text layer ({found} content chars over {pages} pages) "
            "and OCR is disabled."
        )
    ocr_text = _extract_via_ocr(path)
    if not has_usable_text_layer(ocr_text, pages):
        raise PdfExtractionError(
            f"{Path(path).name} yielded only {len(_content_text(ocr_text))} content characters over "
            f"{pages} pages even after OCR - check the file is a document scan and not, say, a cover sheet."
        )
    return ocr_text


def _extract_text_layer(path: str) -> str:
    """Prefer pdftotext (better layout fidelity on table-heavy term sheets than pypdf),
    falling back to pypdf where poppler isn't installed."""
    if shutil.which("pdftotext"):
        try:
            result = subprocess.run(
                ["pdftotext", "-layout", path, "-"],
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout
        except subprocess.TimeoutExpired:
            pass

    try:
        from pypdf import PdfReader

        reader = PdfReader(path)
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception as e:  # noqa: BLE001 - surfaced as a clear ingestion error below
        raise PdfExtractionError(f"Could not read {Path(path).name}: {e}") from e


def _extract_via_ocr(path: str) -> str:
    """Rasterise each page and OCR it.

    `--psm 6` ("assume a single uniform block of text") measurably outperforms
    Tesseract's default page segmentation on term sheets, whose pages are
    dense two-column tables that the default mode fragments.
    """
    for tool in ("pdftoppm", "tesseract"):
        if not shutil.which(tool):
            raise PdfExtractionError(
                f"{Path(path).name} needs OCR (no text layer) but '{tool}' is not installed. "
                "Install poppler and tesseract, or supply the term sheet as DOCX/text."
            )

    with tempfile.TemporaryDirectory() as tmp:
        prefix = str(Path(tmp) / "page")
        subprocess.run(
            ["pdftoppm", "-r", str(_OCR_DPI), "-png", path, prefix],
            capture_output=True,
            timeout=900,
            check=True,
        )
        pages = sorted(Path(tmp).glob("page*.png"))
        if not pages:
            raise PdfExtractionError(f"Could not rasterise {Path(path).name} for OCR.")

        chunks = []
        for page in pages:
            result = subprocess.run(
                ["tesseract", str(page), "stdout", "--psm", "6"],
                capture_output=True,
                text=True,
                timeout=300,
                check=False,
            )
            if result.stdout:
                chunks.append(result.stdout)
        return "\n".join(chunks)
