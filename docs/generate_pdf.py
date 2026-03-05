#!/usr/bin/env python3
"""Convert azure-admin-setup.md to a styled PDF."""

import re
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib import colors

md_file = Path(__file__).parent / "azure-admin-setup.md"
pdf_file = Path(__file__).parent / "azure-admin-setup.pdf"

with open(md_file) as f:
    lines = f.readlines()

styles = getSampleStyleSheet()
styles.add(ParagraphStyle("DocTitle", parent=styles["Title"], fontSize=18, textColor=HexColor("#1a1a2e"), spaceAfter=6))
styles.add(ParagraphStyle("H2Custom", parent=styles["Heading2"], fontSize=13, textColor=HexColor("#1a1a2e"), spaceBefore=16, spaceAfter=6))
styles.add(ParagraphStyle("H3Custom", parent=styles["Heading3"], fontSize=11, textColor=HexColor("#333"), spaceBefore=12, spaceAfter=4))
styles.add(ParagraphStyle("BodyCustom", parent=styles["BodyText"], fontSize=10, leading=14, spaceAfter=4))
styles.add(ParagraphStyle("MetaText", parent=styles["BodyText"], fontSize=9, textColor=HexColor("#666"), spaceAfter=2))
styles.add(ParagraphStyle("BulletCustom", parent=styles["BodyText"], fontSize=10, leading=14, leftIndent=20, bulletIndent=10, spaceAfter=2))
styles.add(ParagraphStyle("CodeCustom", parent=styles["Code"], fontSize=8, backColor=HexColor("#f4f4f4"), leftIndent=10, spaceAfter=4))
styles.add(ParagraphStyle("CheckItem", parent=styles["BodyText"], fontSize=10, leading=14, leftIndent=20, bulletIndent=10, spaceAfter=3))

doc = SimpleDocTemplate(str(pdf_file), pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
story = []


def clean(text):
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"`(.+?)`", r'<font face="Courier" size="8" color="#c0392b">\1</font>', text)
    return text


i = 0
in_code = False
code_block = []
table_data = []

while i < len(lines):
    line = lines[i].rstrip()

    # Code blocks
    if line.startswith("```"):
        if in_code:
            in_code = False
            code_text = "<br/>".join(code_block)
            story.append(Paragraph(code_text, styles["CodeCustom"]))
            story.append(Spacer(1, 4))
            code_block = []
        else:
            in_code = True
        i += 1
        continue

    if in_code:
        code_block.append(line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        i += 1
        continue

    # Tables
    if "|" in line and line.strip().startswith("|"):
        cells = [c.strip() for c in line.strip().split("|")[1:-1]]
        if all(set(c) <= set("- :") for c in cells):
            i += 1
            continue
        table_data.append(cells)
        next_is_table = (i + 1 < len(lines) and "|" in lines[i + 1] and lines[i + 1].strip().startswith("|"))
        if not next_is_table and table_data:
            cleaned = []
            for row in table_data:
                cleaned.append([Paragraph(clean(c), styles["BodyCustom"]) for c in row])
            col_count = max(len(r) for r in cleaned)
            col_width = (A4[0] - 4 * cm) / col_count
            t = Table(cleaned, colWidths=[col_width] * col_count)
            t_style = [
                ("BACKGROUND", (0, 0), (-1, 0), HexColor("#1a1a2e")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#ccc")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ]
            for ri in range(1, len(cleaned)):
                if ri % 2 == 0:
                    t_style.append(("BACKGROUND", (0, ri), (-1, ri), HexColor("#f8f8f8")))
            t.setStyle(TableStyle(t_style))
            story.append(t)
            story.append(Spacer(1, 8))
            table_data = []
        i += 1
        continue

    if not line:
        i += 1
        continue

    if line == "---":
        story.append(HRFlowable(width="100%", thickness=0.5, color=HexColor("#ddd"), spaceBefore=8, spaceAfter=8))
        i += 1
        continue

    if line.startswith("# "):
        story.append(Paragraph(clean(line[2:]), styles["DocTitle"]))
        story.append(HRFlowable(width="100%", thickness=2, color=HexColor("#e63946"), spaceBefore=2, spaceAfter=10))
        i += 1
        continue

    if line.startswith("## "):
        story.append(Paragraph(clean(line[3:]), styles["H2Custom"]))
        i += 1
        continue

    if line.startswith("### "):
        story.append(Paragraph(clean(line[4:]), styles["H3Custom"]))
        i += 1
        continue

    if line.strip().startswith("- [ ]"):
        text = line.strip()[6:]
        story.append(Paragraph(clean(text), styles["CheckItem"], bulletText="\u2610"))
        i += 1
        continue

    if line.strip().startswith("- "):
        text = line.strip()[2:]
        story.append(Paragraph(clean(text), styles["BulletCustom"], bulletText="\u2022"))
        i += 1
        continue

    story.append(Paragraph(clean(line), styles["BodyCustom"]))
    i += 1

doc.build(story)
print(f"PDF created: {pdf_file}")
