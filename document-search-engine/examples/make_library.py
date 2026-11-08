"""Write the demo library: three folders of made-up signed documents, as PDFs with real pages.

Every party, address and figure is invented. Each folder is a history the answers must get right:

- an NDA and an amendment that extends its term and leaves the rest;
- a lease and a rent review memorandum that replaces its rent;
- a services agreement terminated and replaced on the same day, the replacement then novated.

Run from the repository root: .venv/bin/python examples/make_library.py
"""

from __future__ import annotations

import pathlib

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

ROOT = pathlib.Path(__file__).resolve().parent / "library"

# folder -> file -> pages -> lines
LIBRARY: dict[str, dict[str, list[list[str]]]] = {
    "Northwind - Bluefield NDA": {
        "Mutual Non-Disclosure Agreement.pdf": [
            [
                "MUTUAL NON-DISCLOSURE AGREEMENT",
                "THIS AGREEMENT is dated 1 March 2026",
                "BETWEEN Northwind Analytics Limited (company number 01234567) and",
                "Bluefield Robotics GmbH (registration number HRB 987654).",
                "1. Definitions",
                '"Purpose" means evaluating a possible joint development of warehouse picking software.',
                '"Confidential Information" means all information disclosed by one party to the other',
                "in connection with the Purpose, including product roadmaps and test data.",
                "2. Confidentiality",
                "2.1 Each party shall keep the other party's Confidential Information confidential and",
                "use it only for the Purpose.",
                "2.2 Each party shall protect the other party's Confidential Information with reasonable care.",
            ],
            [
                "3. Exceptions",
                "The obligations in Clause 2 do not apply to information which is or becomes public",
                "other than through a breach of this Agreement.",
                "4. Term",
                "4.1 This Agreement continues for three (3) years from its date.",
                "4.2 The obligations in Clause 2 survive termination for five (5) years.",
                "5. Governing Law",
                "5.1 This Agreement is governed by the laws of England and Wales and the courts of",
                "England have exclusive jurisdiction.",
                "Signed for Northwind Analytics Limited and for Bluefield Robotics GmbH.",
            ],
        ],
        "Amendment No 1 to the NDA.pdf": [
            [
                "AMENDMENT NO. 1",
                "THIS AMENDMENT is dated 1 September 2026",
                "BETWEEN Northwind Analytics Limited and Bluefield Robotics GmbH.",
                "1. Background",
                "The parties entered into a Mutual Non-Disclosure Agreement dated 1 March 2026",
                "(the Agreement) and wish to extend its term.",
                "2. Amendment",
                "2.1 Clause 4.1 of the Agreement is deleted and replaced with: This Agreement continues",
                "for four (4) years from its date.",
                "3. Continuation",
                "3.1 All other terms of the Agreement remain in full force and effect.",
            ]
        ],
    },
    "Harbour Street Lease": {
        "Lease of Unit 4 Harbour Street.pdf": [
            [
                "LEASE",
                "THIS LEASE is dated 1 April 2023",
                "BETWEEN Harbour Street Properties Limited (the Landlord) and",
                "Northwind Analytics Limited (the Tenant).",
                "1. Premises",
                "1.1 The Landlord lets to the Tenant Unit 4, Harbour Street, Exampletown.",
                "2. Term",
                "2.1 The term is ten (10) years from 1 April 2023.",
                "3. Rent",
                "3.1 The Tenant shall pay rent of £50,000 per year, by equal quarterly payments in advance.",
                "3.2 The rent shall be reviewed on 1 April 2025 and every two years after that.",
            ],
            [
                "4. Deposit",
                "4.1 The Tenant shall pay a deposit of £12,500 on the date of this Lease.",
                "5. Break Clause",
                "5.1 The Tenant may end this Lease on 1 April 2028 by giving the Landlord at least",
                "six (6) months' written notice.",
                "6. Governing Law",
                "6.1 This Lease is governed by the laws of England and Wales.",
                "Signed for the Landlord and for the Tenant.",
            ],
        ],
        "Rent Review Memorandum 2025.pdf": [
            [
                "RENT REVIEW MEMORANDUM",
                "THIS MEMORANDUM is dated 1 April 2025",
                "BETWEEN Harbour Street Properties Limited and Northwind Analytics Limited.",
                "1. Review",
                "1.1 Under clause 3.2 of the Lease dated 1 April 2023 the rent has been reviewed.",
                "1.2 With effect from 1 April 2025 the rent is £56,000 per year.",
                "2. Continuation",
                "2.1 The Lease dated 1 April 2023 continues in full force, varied only as above.",
            ]
        ],
    },
    "Acme Services": {
        "Services Agreement 2024.pdf": [
            [
                "SERVICES AGREEMENT",
                "THIS AGREEMENT is dated 1 January 2024",
                "BETWEEN Acme Services Limited (the Supplier) and Bluefield Robotics GmbH (the Customer).",
                "1. Services",
                "1.1 The Supplier shall provide warehouse software support services.",
                "2. Fees",
                "2.1 The Customer shall pay a fee of £8,000 per month.",
                "3. Governing Law",
                "3.1 This Agreement is governed by the laws of England and Wales.",
            ]
        ],
        "Termination Agreement.pdf": [
            [
                "TERMINATION AGREEMENT",
                "THIS AGREEMENT is dated 1 March 2025",
                "BETWEEN Acme Services Limited and Bluefield Robotics GmbH.",
                "1. Termination",
                "1.1 The Services Agreement dated 1 January 2024 terminates on the date of this",
                "agreement and ceases to have effect.",
                "1.2 Any fees unpaid under it shall be paid within thirty (30) days.",
                "2. New agreement",
                "2.1 On the same date the parties are entering into a new Services Agreement.",
            ]
        ],
        "New Services Agreement.pdf": [
            [
                "SERVICES AGREEMENT",
                "THIS AGREEMENT is dated 1 March 2025",
                "BETWEEN Acme Services Limited (the Supplier) and Bluefield Robotics GmbH (the Customer).",
                "1. Services",
                "1.1 The Supplier shall provide warehouse software support and hosting services.",
                "2. Fees",
                "2.1 The Customer shall pay a fee of £9,500 per month.",
            ],
            [
                "3. Term",
                "3.1 This Agreement continues for three (3) years from its date.",
                "4. Governing Law",
                "4.1 This Agreement is governed by the laws of Germany and the courts of Berlin",
                "have exclusive jurisdiction.",
            ],
        ],
        "Deed of Novation.pdf": [
            [
                "DEED OF NOVATION",
                "THIS DEED is dated 1 June 2025",
                "BETWEEN Acme Services Limited (the Outgoing Supplier), Acme Holdings GmbH",
                "(the Incoming Supplier) and Bluefield Robotics GmbH (the Customer).",
                "1. Novation",
                "1.1 The Services Agreement dated 1 March 2025 is novated from the Outgoing Supplier",
                "to the Incoming Supplier with effect from the date of this deed.",
                "1.2 The Services Agreement dated 1 March 2025 continues in force on its terms.",
            ]
        ],
    },
}


def write(path: pathlib.Path, pages: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(path), pagesize=A4)
    pdf.setTitle(path.stem)
    pdf.setAuthor("")
    for lines in pages:
        y = 800
        for line in lines:
            pdf.setFont("Helvetica-Bold" if line.isupper() else "Helvetica", 10.5)
            pdf.drawString(60, y, line)
            y -= 18
        pdf.showPage()
    pdf.save()


# one Word document, so the library is not PDFs alone: Word files have no pages until printed
SIDE_LETTER = [
    "SIDE LETTER",
    "THIS LETTER is dated 15 April 2023",
    "FROM Harbour Street Properties Limited TO Northwind Analytics Limited.",
    "1. Rent-free period",
    "1.1 Despite clause 3.1 of the Lease dated 1 April 2023, no rent is payable for the first three (3) months of the term.",
    "2. Continuation",
    "2.1 The Lease continues in full force, varied only as above.",
]


def write_docx(path: pathlib.Path, lines: list[str]) -> None:
    from docx import Document

    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.core_properties.author = ""
    doc.save(str(path))


if __name__ == "__main__":
    for folder, files in LIBRARY.items():
        for name, pages in files.items():
            write(ROOT / folder / name, pages)
    write_docx(ROOT / "Harbour Street Lease" / "Side Letter on Rent.docx", SIDE_LETTER)
    print("wrote", sum(len(f) for f in LIBRARY.values()), "documents in", len(LIBRARY), "folders")
