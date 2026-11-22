"""Assemble the review edition. Requires python-docx, PyMuPDF, resvg-py and LibreOffice Writer."""
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import fitz
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH as Align, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
import resvg_py

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "final"
NAME = "Magister_Thesis_Khilchevskyi"
FONT = "Times New Roman"
TITLE = "Педагогічні умови підтримки навчальної мотивації здобувачів медичної освіти в умовах дистанційного навчання в період воєнного часу"
SOURCES = [ROOT / "drafts" / (n + ".md") for n in
           ["introduction", "chapter_1", "chapter_2", "conclusions", "bibliography"]]
SOURCES += [ROOT / "appendices" / ("appendix_" + str(i) + ".md") for i in (1, 2)]


def element(tag, **attrs):
    e = OxmlElement("w:" + tag)
    for key, value in attrs.items():
        e.set(qn("w:" + key), str(value))
    return e


def rich(p, text):
    for i, part in enumerate(re.split(r"\*\*", text)):
        if part:
            run = p.add_run(part)
            if i % 2:
                run.bold = True
    return p


def configure(d):
    sec = d.sections[0]
    sec.page_height, sec.page_width = Mm(297), Mm(210)
    sec.top_margin, sec.bottom_margin = Mm(20), Mm(20)
    sec.left_margin, sec.right_margin = Mm(30), Mm(10)
    sec.header_distance = Mm(9)
    sec.different_first_page_header_footer = True
    normal = d.styles["Normal"]
    pf = normal.paragraph_format
    pf.alignment = Align.JUSTIFY
    pf.first_line_indent = Mm(12.5)
    pf.line_spacing = 1.5
    pf.space_before = pf.space_after = Pt(0)
    pf.widow_control = True
    for name in ["Normal", "Heading 1", "Heading 2", "Header", "Caption",
                 "ThesisTitle", "ThesisPlain", "ThesisTable", "ThesisTOC", "ThesisRef"]:
        sty = d.styles[name] if name in d.styles else d.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        if name != "Normal":
            sty.base_style = normal
            sty.paragraph_format.first_line_indent = Pt(0)
        sty.font.name, sty.font.size = FONT, Pt(14)
        sty.font.bold, sty.font.italic = name == "Heading 1", False
        sty.font.color.rgb = RGBColor(0, 0, 0)
        rp = sty.element.get_or_add_rPr()
        fonts = rp.find(qn("w:rFonts"))
        for k in list(fonts.attrib):
            if k.lower().endswith("theme"):
                del fonts.attrib[k]
        for k in ["ascii", "hAnsi", "eastAsia", "cs"]:
            fonts.set(qn("w:" + k), FONT)
        rp.append(element("lang", val="uk-UA"))
    for name in ["Heading 1", "Heading 2"]:
        p = d.styles[name].paragraph_format
        p.keep_with_next = p.keep_together = True
        p.space_before = Pt(10 if name.endswith("2") else 0)
        p.space_after = Pt(7)
        p.page_break_before = name.endswith("1")
        p.alignment = Align.LEFT if name.endswith("2") else Align.CENTER
        if name.endswith("2"):
            p.first_line_indent = Mm(12.5)
    for name in ["Caption", "ThesisTitle"]:
        d.styles[name].paragraph_format.alignment = Align.CENTER
    d.styles["ThesisTable"].paragraph_format.alignment = Align.LEFT
    d.styles["ThesisTOC"].paragraph_format.line_spacing = 1.0
    d.styles["ThesisTOC"].paragraph_format.space_after = Pt(3)
    d.styles["ThesisRef"].paragraph_format.alignment = Align.LEFT
    d.styles["ThesisRef"].paragraph_format.left_indent = Mm(8)
    d.styles["ThesisRef"].paragraph_format.first_line_indent = Mm(-8)
    d.styles["ThesisRef"].paragraph_format.keep_together = True
    p = sec.header.paragraphs[0]
    p.alignment = Align.CENTER
    p.paragraph_format.line_spacing = 1
    p._p.append(element("fldSimple", instr=" PAGE "))
    d.settings.element.append(element("updateFields", val="true"))
    d.core_properties.title = TITLE
    d.core_properties.author = "Богдан Хільчевський"
    d.core_properties.subject = "Магістерська робота з освітніх наук"


def title(d):
    # These fields concern this author's actual enrolment, rather than the example programme.
    def line(value, before=0, bold=False, right=False):
        p = d.add_paragraph(style="ThesisTitle")
        if right:
            p.alignment = Align.LEFT
            p.paragraph_format.left_indent = Mm(73)
        p.paragraph_format.space_before = Pt(before)
        p.add_run(value).bold = bold
    line("Міністерство охорони здоров’я України")
    line("Харківський національний медичний університет")
    line("Кафедра української мови та педагогіки", 8)
    line("МАГІСТЕРСЬКА РОБОТА", 42, True)
    line("за спеціальністю А1 «Освітні науки»", 8)
    line("Освітньо-професійна програма «Педагогіка вищої школи»")
    line("на тему", 10)
    line("«" + TITLE + "»", 5, True)
    line("Виконав студент ___ курсу, групи ______", 26, right=True)
    line("Факультет ______________________", right=True)
    line("Хільчевський Богдан Сергійович", right=True)
    line("Наукова керівниця", 10, right=True)
    line("Разумна Алла Григорівна,", right=True)
    line("кандидат педагогічних наук, доцент", right=True)
    line("Рецензент ______________________", 10, right=True)
    line("Харків — 20___", 20)


def bookmark(p, idx):
    p._p.insert(1, element("bookmarkStart", id=idx, name="Thesis_" + str(idx)))
    p._p.append(element("bookmarkEnd", id=idx))


def parse(path):
    lines = [s.strip() for s in path.read_text(encoding="utf-8-sig").splitlines() if s.strip()]
    events = []
    i = 0
    while i < len(lines):
        s = lines[i]
        if s.startswith("# "):
            h = s[2:]
            if i + 1 < len(lines) and lines[i + 1].startswith("# "):
                i += 1
                h += "\n" + lines[i][2:]
            events.append(("h1", h))
        elif s.startswith("## "):
            events.append(("h2", s[3:]))
        elif s.startswith("**Таблиця "):
            caption = s.strip("*")
            data = []
            while i + 1 < len(lines) and lines[i + 1].startswith("|"):
                i += 1
                cells = [x.strip() for x in lines[i].strip("|").split("|")]
                if not all(re.fullmatch(r"[-: ]+", x) for x in cells):
                    data.append(cells)
            events.append(("table", (caption, data)))
        elif s.startswith("!["):
            i += 1
            caption = lines[i].strip("*").replace("Рисунок ", "Рис. ", 1)
            events.append(("figure", caption))
        else:
            events.append(("ref" if path.stem == "bibliography" else "p", s))
        i += 1
    return events


def toc(d, headings, pages):
    p = d.add_paragraph("ЗМІСТ", style="ThesisTitle")
    p.runs[0].bold = True
    p.paragraph_format.page_break_before = True
    p.paragraph_format.space_after = Pt(12)
    for k, (level, name) in enumerate(headings, 1):
        p = d.add_paragraph(style="ThesisTOC")
        p.paragraph_format.keep_together = True
        p.paragraph_format.right_indent = Mm(11)
        p.paragraph_format.left_indent = Mm(7 if level == "h2" else 0)
        p.paragraph_format.tab_stops.add_tab_stop(Mm(170), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        if k == 1:
            p.add_run()._r.append(element("fldChar", fldCharType="begin"))
            instr = element("instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = ' TOC \\o "1-2" \\h \\z '
            p.add_run()._r.append(instr)
            p.add_run()._r.append(element("fldChar", fldCharType="separate"))
        a = element("hyperlink", anchor="Thesis_" + str(k))
        run = element("r")
        tx = element("t")
        tx.text = name.replace("\n", " ")
        run.append(tx)
        a.append(run)
        p._p.append(a)
        p.add_run("\t" + str(pages[k-1] if pages else 1))
        if k == len(headings):
            p.add_run()._r.append(element("fldChar", fldCharType="end"))


def add_table(d, content, table_number):
    caption, data = content
    num, label = caption.split(" — ", 1)
    widths = [36, 49, 57, 28] if table_number == 1 else [32, 43, 47, 48]
    for s, align in [(num, Align.RIGHT), (label, Align.CENTER)]:
        p = d.add_paragraph(s, style="ThesisPlain")
        p.alignment = align
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.space_before = Pt(3)
    table = d.add_table(rows=0, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False
    for col, width in zip(table.columns, widths):
        col.width = Mm(width)
    for ri, vals in enumerate(data):
        row = table.add_row()
        row.height = Mm(8)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        pr = row._tr.get_or_add_trPr()
        pr.append(element("cantSplit"))
        if not ri:
            pr.append(element("tblHeader", val="true"))
        for cell, value, width in zip(row.cells, vals, widths):
            cell.width = Mm(width)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.style = "ThesisTable"
            p.paragraph_format.keep_with_next = ri < len(data)-1
            # Compact spacing within table cells, full thesis type size.
            p.paragraph_format.line_spacing = 1
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            p.add_run(value).bold = not ri


def assemble(events, headings, pages, image, path):
    d = Document()
    configure(d)
    title(d)
    toc(d, headings, pages)
    n = nt = 0
    for tag, value in events:
        if tag.startswith("h"):
            n += 1
            p = d.add_paragraph(value, style="Heading " + tag[-1])
            bookmark(p, n)
        elif tag == "table":
            nt += 1
            add_table(d, value, nt)
        elif tag == "figure":
            p = d.add_paragraph(style="ThesisPlain")
            p.alignment = Align.CENTER
            p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(str(image), width=Mm(156))
            p = d.add_paragraph(value, style="Caption")
            p.paragraph_format.keep_together = True
        else:
            p = d.add_paragraph(style="ThesisRef" if tag == "ref" else "Normal")
            rich(p, value)
    d.save(path)


def normalize(text):
    return re.sub(r"\s+", " ", text).strip()


def locate(pdf, headings):
    outline = pdf.get_toc()
    if len(outline) == len(headings) and all(normalize(row[1]).casefold() == normalize(h).casefold() for row, (_, h) in zip(outline, headings)):
        return [row[2] for row in outline]
    contents = [normalize(p.get_text()) for p in pdf]
    page_nos = []
    start = next(i for i in range(2,len(pdf)) if contents[i].lstrip("0123456789 ").startswith("ВСТУП"))
    cursor = start
    for _, title_ in headings:
        needle = normalize(title_)
        # Headings from the outline also permit wrapping on hyphens or a title break.
        for i in range(cursor, len(pdf)):
            if needle in contents[i]:
                cursor = i
                page_nos.append(i+1)
                break
        else:
            raise ValueError("Heading not found: " + needle)
    return page_nos


def render(office, input_path, folder, profile):
    subprocess.run([str(office), "-env:UserInstallation=" + profile.as_uri(), "--headless",
                    "--convert-to", "pdf:writer_pdf_Export", "--outdir", str(folder), str(input_path)],
                   check=True, timeout=200)
    dest = folder / (input_path.stem + ".pdf")
    if not dest.is_file() or not dest.stat().st_size:
        raise RuntimeError("PDF was not rendered")
    return dest


def main():
    OUT.mkdir(exist_ok=True)
    exe = (os.environ.get("SOFFICE") or shutil.which("soffice") or
           r"C:\Program Files\LibreOffice\program\soffice.com")
    events = [item for path in SOURCES for item in parse(path)]
    headings = [(t, v) for t, v in events if t in ("h1", "h2")]
    with tempfile.TemporaryDirectory() as temp:
        tmp = Path(temp).resolve()
        png = tmp / "figure.png"
        svg = (ROOT/"figures/model_2_1.svg").read_text(encoding="utf-8")
        png.write_bytes(resvg_py.svg_to_bytes(svg_string=svg, width=1650))
        docx = OUT / (NAME + ".docx")
        pages = []
        for attempt in range(4):
            assemble(events, headings, pages, png, docx)
            folder = tmp / str(attempt)
            folder.mkdir()
            dest = render(exe, docx, folder, tmp/"lo")
            with fitz.open(dest) as pdf:
                observed = locate(pdf, headings)
            if observed == pages:
                shutil.copyfile(dest, OUT / (NAME + ".pdf"))
                break
            pages = observed
        else:
            raise RuntimeError("TOC pagination did not stabilize")
    print(json.dumps({normalize(h):pg for (_,h),pg in zip(headings,pages)},ensure_ascii=False))


if __name__ == "__main__":
    main()
