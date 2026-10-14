"""Build Anwar-Orozco-CV.docx from the same content as index.html.

Dev-time tool only (the website never runs this).
Usage:  python tools/build-docx.py      (run from the cv-site folder)
Needs:  pip install python-docx
Keep CV text here in sync with index.html when you edit either one.
"""
from pathlib import Path
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

OUT = Path(__file__).resolve().parent.parent / "Anwar-Orozco-CV.docx"

NAME = "Anwar Orozco"
ROLE = "Senior QA Engineer"
CONTACT = "Chihuahua, Mexico  ·  amot.anly@gmail.com  ·  flamingflamingo-net.github.io"
SUMMARY = ("8+ years and 18+ projects in software testing. I build test automation with "
           "Playwright and Karate, test Salesforce platforms and lead QA teams in Agile environments.")
SKILLS = [
    ("Automation & AI", "Playwright, Karate, AI workflow, AI automation, Claude Code"),
    ("Languages & APIs", "TypeScript, JavaScript, SQL, Postman (API testing)"),
    ("Platforms", "Salesforce, Magento"),
    ("QA & Methods", "Test planning, Manual testing, Verification & Validation, BDD, Agile, Scrum"),
]
JOBS = [
    ("Senior QA Engineer, Playwright Automation", "Concentrix Catalyst", "Feb 2025 – May 2026",
     "Built a Playwright + TypeScript automation framework from scratch, including custom API verification tools."),
    ("Senior QA Engineer, Salesforce", "Concentrix Catalyst", "Nov 2023 – Feb 2025",
     "Manual test + test plans and Karate automation for enterprise Salesforce releases. API validation with Postman."),
    ("Senior QA Engineer", "Concentrix Catalyst (formerly PK)", "Oct 2019 – Nov 2023",
     "Led a 10-person QA team across 5 projects (12 in total). Test planning, code audits, SQL and API validation."),
    ("Junior QA / Tester", "Gamma Partners", "Feb 2018 – Oct 2019",
     "Manual testing, verification & validation and code audits across 3 projects."),
]
EDUCATION = ("Computational Technologies Engineering", "Tecnológico de Monterrey (ITESM) · 2021")
LANGUAGES = "Spanish (Native)  ·  English (C2)"

FG = RGBColor(0x18, 0x18, 0x1B)
MUTED = RGBColor(0x5F, 0x5F, 0x66)
CONTENT_W = Mm(210 - 18 - 18)


def para(doc, space_after=2, space_before=0):
    p = doc.add_paragraph()
    f = p.paragraph_format
    f.space_after, f.space_before, f.line_spacing = Pt(space_after), Pt(space_before), 1.1
    return p


def run(p, text, size=10.5, bold=False, color=FG):
    r = p.add_run(text)
    r.font.size, r.font.bold, r.font.color.rgb = Pt(size), bold, color
    return r


def heading(doc, text):
    p = para(doc, space_after=4, space_before=12)
    r = run(p, text.upper(), size=9, bold=True, color=MUTED)
    r.font.spacing = Pt(1)  # letter-spacing (w:spacing in twentieths of a point via python-docx)
    # bottom border line
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in {"val": "single", "sz": "4", "space": "2", "color": "DCDCE0"}.items():
        bottom.set(qn(f"w:{k}"), v)
    bdr.append(bottom)
    pPr.append(bdr)


def build():
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.left_margin = s.right_margin = Mm(18)
    s.top_margin = s.bottom_margin = Mm(16)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal.font.size = Pt(10.5)

    doc.core_properties.title = f"{NAME} – CV"
    doc.core_properties.author = NAME

    run(para(doc, 0), NAME, size=22, bold=True)
    run(para(doc, 4), ROLE, size=12, color=MUTED)
    run(para(doc, 8), CONTACT, size=9.5, color=MUTED)
    run(para(doc, 0), SUMMARY)

    heading(doc, "Skills")
    for label, items in SKILLS:
        p = para(doc, 1)
        run(p, f"{label}: ", bold=True)
        run(p, items, color=MUTED)

    heading(doc, "Experience")
    for title, org, dates, desc in JOBS:
        p = para(doc, 0, space_before=4)
        p.paragraph_format.tab_stops.add_tab_stop(CONTENT_W, WD_TAB_ALIGNMENT.RIGHT)
        p.paragraph_format.keep_with_next = True
        run(p, title, bold=True)
        run(p, f"  ·  {org}", color=MUTED)
        run(p, f"\t{dates}", size=9.5, color=MUTED)
        run(para(doc, 2), desc, size=10, color=MUTED)

    heading(doc, "Education")
    run(para(doc, 0), EDUCATION[0], bold=True)
    run(para(doc, 2), EDUCATION[1], size=10, color=MUTED)

    heading(doc, "Languages")
    run(para(doc, 0), LANGUAGES)

    doc.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
