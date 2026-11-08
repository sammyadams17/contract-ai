"""Write the demo's made-up documents: a mutual NDA template, a term sheet for one deal, and
the client's last signed NDA with another company, to draft from as a precedent.

Every name, figure and address here is invented. Run from the repository root:

    .venv/bin/python examples/make_demo.py
"""

from __future__ import annotations

import pathlib

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

HERE = pathlib.Path(__file__).resolve().parent / "nda"

TEMPLATE = [
    ("title", "MUTUAL NON-DISCLOSURE AGREEMENT"),
    ("body", "THIS AGREEMENT is dated [•]"),
    ("body", "BETWEEN:"),
    ("body", "(1) [•], a company incorporated in [•] with company number [•] whose registered "
             "office is at [•] (the \"First Party\"); and"),
    ("body", "(2) [•], a company incorporated in [•] with company number [•] whose registered "
             "office is at [•] (the \"Second Party\")."),
    ("h1", "Definitions"),
    ("clause", "In this Agreement:"),
    ("body", "\"Affiliate\" means, in relation to a party, any entity that controls, is controlled by "
             "or is under common control with that party;"),
    ("body", "\"Confidential Information\" means all information of whatever nature disclosed by "
             "one party to the other in connection with the Purpose, whether before or after the "
             "date of this Agreement, including:"),
    ("limb", "business plans, customer lists and pricing; and"),
    ("limb", "technical information, know-how and source code;"),
    ("body", "\"Purpose\" means [•]; and"),
    ("body", "\"Representatives\" means a party's Affiliates and its and their officers, employees "
             "and professional advisers."),
    ("h1", "Confidentiality Obligations"),
    ("clause", "Each party shall keep the other party's Confidential Information confidential "
               "and shall use it only for the Purpose."),
    ("clause", "Each party may disclose the other party's Confidential Information to its "
               "Representatives who need to know it for the Purpose, provided that it procures that "
               "they keep it confidential on the terms of this Agreement."),
    ("clause", "Each party shall protect the other party's Confidential Information with "
               "[at least the same degree of care it uses for its own confidential information]"
               "[reasonable care]."),
    ("h1", "Exceptions"),
    ("clause", "The obligations in Clause 2 do not apply to information which:"),
    ("limb", "is or becomes public other than through a breach of this Agreement;"),
    ("limb", "was lawfully in the recipient's possession before disclosure; or"),
    ("limb", "is required to be disclosed by law or by any regulatory authority."),
    ("h1", "Term and Termination"),
    ("clause", "This Agreement continues for [two (2)] years from its date."),
    ("clause", "The obligations in Clause 2 survive termination for [•] years."),
    ("h1", "Return or Destruction"),
    ("clause", "On the other party's written request, each party shall promptly return or destroy "
               "the other party's Confidential Information."),
    ("h1", "Remedies"),
    ("clause", "Each party acknowledges that damages alone may not be an adequate remedy for a "
               "breach of this Agreement and that the other party is entitled to seek injunctive "
               "relief."),
    ("h1", "Governing Law and Jurisdiction"),
    ("clause", "This Agreement is governed by the laws of [England and Wales] and the courts of "
               "[England] have exclusive jurisdiction."),
    ("title2", "Execution page"),
    ("body", "Signed for and on behalf of the First Party: ____________________"),
    ("body", "Signed for and on behalf of the Second Party: ____________________"),
]

# The client's last signed NDA, with another company: what a lawyer would start from instead of
# the blank template. Its parties, purpose, terms and law are that deal's, and it carries a clause
# the new deal does not ask for.
PRECEDENT = [
    ("title", "MUTUAL NON-DISCLOSURE AGREEMENT"),
    ("body", "THIS AGREEMENT is dated 12 March 2025"),
    ("body", "BETWEEN:"),
    ("body", "(1) Northwind Analytics Limited, a company incorporated in England and Wales with company "
             "number 01234567 whose registered office is at 10 Example Street, Exampletown XX1 1XX "
             "(the \"First Party\"); and"),
    ("body", "(2) Saltmarsh Data Limited, a company incorporated in Ireland with company number 654321 "
             "whose registered office is at 4 Sample Quay, Dublin 2 (the \"Second Party\")."),
    ("h1", "Definitions"),
    ("clause", "In this Agreement:"),
    ("body", "\"Affiliate\" means, in relation to a party, any entity that controls, is controlled by "
             "or is under common control with that party;"),
    ("body", "\"Confidential Information\" means all information of whatever nature disclosed by "
             "one party to the other in connection with the Purpose, whether before or after the "
             "date of this Agreement, including:"),
    ("limb", "business plans, customer lists and pricing; and"),
    ("limb", "technical information, know-how and source code;"),
    ("body", "\"Purpose\" means evaluating a possible data-sharing arrangement for retail demand "
             "forecasting; and"),
    ("body", "\"Representatives\" means a party's Affiliates and its and their officers, employees "
             "and professional advisers."),
    ("h1", "Confidentiality Obligations"),
    ("clause", "Each party shall keep the other party's Confidential Information confidential "
               "and shall use it only for the Purpose."),
    ("clause", "Each party may disclose the other party's Confidential Information to its "
               "Representatives who need to know it for the Purpose, provided that it procures that "
               "they keep it confidential on the terms of this Agreement."),
    ("clause", "Each party shall protect the other party's Confidential Information with "
               "at least the same degree of care it uses for its own confidential information."),
    ("h1", "Exceptions"),
    ("clause", "The obligations in Clause 2 do not apply to information which:"),
    ("limb", "is or becomes public other than through a breach of this Agreement;"),
    ("limb", "was lawfully in the recipient's possession before disclosure; or"),
    ("limb", "is required to be disclosed by law or by any regulatory authority."),
    ("h1", "Term and Termination"),
    ("clause", "This Agreement continues for two (2) years from its date."),
    ("clause", "The obligations in Clause 2 survive termination for three (3) years."),
    ("h1", "Return or Destruction"),
    ("clause", "On the other party's written request, each party shall promptly return or destroy "
               "the other party's Confidential Information."),
    ("h1", "Non-Solicitation"),
    ("clause", "Neither party shall, during the term of this Agreement and for six (6) months after it "
               "ends, solicit for employment any employee of the other party involved in the Purpose."),
    ("h1", "Remedies"),
    ("clause", "Each party acknowledges that damages alone may not be an adequate remedy for a "
               "breach of this Agreement and that the other party is entitled to seek injunctive "
               "relief."),
    ("h1", "Governing Law and Jurisdiction"),
    ("clause", "This Agreement is governed by the laws of Ireland and the courts of Ireland have "
               "exclusive jurisdiction."),
    ("title2", "Execution page"),
    ("body", "Signed for and on behalf of the First Party: A. Example, Director"),
    ("body", "Signed for and on behalf of the Second Party: B. Sample, Director"),
]

TERMSHEET = [
    ("title", "Heads of Terms: Mutual Confidentiality Agreement"),
    ("body", "These heads of terms are not legally binding, except for paragraph 9."),
    ("h1", "1. Parties"),
    ("body", "Northwind Analytics Limited, a company incorporated in England and Wales with company "
             "number 01234567, registered office 10 Example Street, Exampletown XX1 1XX; and "
             "Bluefield Robotics GmbH, a company incorporated in Germany with registration number "
             "HRB 987654, registered office Musterstrasse 5, 10115 Berlin."),
    ("h1", "2. Purpose"),
    ("body", "Evaluating a possible joint development of warehouse picking software."),
    ("h1", "3. Additional Confidential Information"),
    ("body", "Confidential Information also includes product roadmaps and test data."),
    ("h1", "4. Standard of care"),
    ("body", "Reasonable care."),
    ("h1", "5. Permitted recipients"),
    ("body", "Each party may also disclose to its actual or prospective funders, on the same terms."),
    ("h1", "6. Duration"),
    ("body", "The agreement lasts three (3) years. The confidentiality obligations survive for five "
             "(5) years after it ends."),
    ("h1", "7. Retention"),
    ("body", "A party may keep copies it is required to keep by law, which remain confidential."),
    ("h1", "8. Governing law"),
    ("body", "English law; courts of England."),
    ("h1", "9. Confidentiality of these heads of terms"),
    ("body", "These heads of terms are themselves confidential."),
    ("h1", "10. Timetable"),
    ("body", "The parties aim to sign the agreement within two weeks."),
]


# One outline list, as a template numbers its clauses: 1. heading, 1.1 clause, (a) limb.
_NUMBERING = """
<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="90">
  <w:multiLevelType w:val="multilevel"/>
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
    <w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/>
    <w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/>
    <w:pPr><w:ind w:left="1134" w:hanging="567"/></w:pPr></w:lvl>
</w:abstractNum>"""


def _numbered(paragraph, level: int, num_id: str) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    numpr = OxmlElement("w:numPr")
    ilvl, num = OxmlElement("w:ilvl"), OxmlElement("w:numId")
    ilvl.set(qn("w:val"), str(level))
    num.set(qn("w:val"), num_id)
    numpr.extend([ilvl, num])
    ppr.insert(0, numpr)


def _outline(doc) -> str:
    numbering = doc.part.numbering_part.element
    numbering.insert(0, parse_xml(_NUMBERING))
    num = parse_xml(
        '<w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' w:numId="90"><w:abstractNumId w:val="90"/></w:num>'
    )
    numbering.append(num)
    return "90"


def write(rows: list[tuple[str, str]], path: pathlib.Path) -> None:
    doc = Document()
    num_id = _outline(doc)
    for kind, text in rows:
        if kind == "title":
            doc.add_heading(text, level=0)
        elif kind == "title2":
            doc.add_heading(text, level=1)
        elif kind == "h1":
            if num_id and any(k == "clause" for k, _ in rows):
                _numbered(doc.add_heading(text, level=1), 0, num_id)
            else:
                doc.add_heading(text, level=1)
        elif kind == "clause":
            _numbered(doc.add_paragraph(text), 1, num_id)
        elif kind == "limb":
            _numbered(doc.add_paragraph(text), 2, num_id)
        else:
            doc.add_paragraph(text)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


if __name__ == "__main__":
    write(TEMPLATE, HERE / "template.docx")
    write(TERMSHEET, HERE / "termsheet.docx")
    write(PRECEDENT, HERE / "precedent.docx")
    (HERE / "client-details.json").write_text(
        '{\n "Client registered office": "10 Example Street, Exampletown XX1 1XX",\n'
        ' "Client notice email": "legal@northwind.example"\n}\n'
    )
    print("wrote", *sorted(p.name for p in HERE.iterdir()))
