from __future__ import annotations

import pathlib
import subprocess
import sys

import pytest

from document_search.library import Library

ROOT = pathlib.Path(__file__).resolve().parents[1]
LIBRARY = ROOT / "examples" / "library"


@pytest.fixture(scope="session")
def library(tmp_path_factory) -> Library:
    if not LIBRARY.exists():
        subprocess.run([sys.executable, str(ROOT / "examples" / "make_library.py")], check=True)
    lib = Library(tmp_path_factory.mktemp("db") / "library.db")
    lib.ingest(LIBRARY)
    return lib
