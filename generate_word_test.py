#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional matematik test generatori
Chiziqli tengsizlik va tengsizliklar sistemasi (Masalalar 46-83)
Natija: Microsoft Word (.docx) OMML equation formatida
Faqat standart kutubxonalar: xml.etree.ElementTree + zipfile
"""

import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

# ─── Namespaces ───────────────────────────────────────────────────────────────
W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M  = "http://schemas.openxmlformats.org/officeDocument/2006/math"

ET.register_namespace('w', W)
ET.register_namespace('m', M)
ET.register_namespace('r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace('xml', "http://www.w3.org/XML/1998/namespace")

def w(t): return f"{{{W}}}{t}"
def m(t): return f"{{{M}}}{t}"

# ─── OMML string helpers (pure string concatenation) ─────────────────────────

def mr(text):
    """Math run"""
    safe = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return (f'<m:r xmlns:m="{M}">'
            f'<m:rPr><m:rFont m:val="Cambria Math"/><m:sty m:val="p"/></m:rPr>'
            f'<m:t xml:space="preserve">{safe}</m:t></m:r>')

def mf(num_str, den_str):
    """Stacked fraction"""
    safe_n = num_str.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    safe_d = den_str.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return (f'<m:f xmlns:m="{M}">'
            f'<m:num>{mr(safe_n)}</m:num>'
            f'<m:den>{mr(safe_d)}</m:den>'
            f'</m:f>')

def msup(base, sup):
    """Superscript"""
    return (f'<m:sSup xmlns:m="{M}">'
            f'<m:e>{mr(base)}</m:e>'
            f'<m:sup>{mr(sup)}</m:sup>'
            f'</m:sSup>')

def mg(*parts):
    """Group multiple OMML strings"""
    return "".join(parts)

def omml_para(inner):
    """Block math paragraph"""
    return (f'<m:oMathPara xmlns:m="{M}">'
            f'<m:oMath xmlns:m="{M}">{inner}</m:oMath>'
            f'</m:oMathPara>')

def omml_inline(inner):
    """Inline math"""
    return f'<m:oMath xmlns:m="{M}">{inner}</m:oMath>'


# ─── Questions data ───────────────────────────────────────────────────────────
# Each entry: num, source (original problem #), q, A, B, C, D, ans
# All values are OMML inner XML strings

QUESTIONS = [
    # 1 — problem 46
    {"num":1,"src":"46",
     "q": mr("3 < 2x + 7 ≤ 11"),
     "A": mr("−2 < x ≤ 2"),
     "B": mr("−2 ≤ x ≤ 2"),
     "C": mr("−2 < x < 2"),
     "D": mr("−1 < x ≤ 2"),
     "ans":"A"},
    # 2 — problem 47
    {"num":2,"src":"47",
     "q": mr("−3 < 5 + 4x < 21"),
     "A": mr("−2 ≤ x < 4"),
     "B": mr("−2 < x < 4"),
     "C": mr("−2 < x ≤ 4"),
     "D": mr("−1 < x < 4"),
     "ans":"B"},
    # 3 — problem 48
    {"num":3,"src":"48",
     "q": mr("−7 ≤ 1 − 4x ≤ 13"),
     "A": mr("−3 < x ≤ 2"),
     "B": mr("−3 ≤ x < 2"),
     "C": mr("−3 ≤ x ≤ 2"),
     "D": mr("−2 ≤ x ≤ 3"),
     "ans":"C"},
    # 4 — problem 49
    {"num":4,"src":"49",
     "q": mr("−8 ≤ 7 − 3x < 16"),
     "A": mr("−3 ≤ x ≤ 5"),
     "B": mr("−3 < x < 5"),
     "C": mr("−3 ≤ x < 5"),
     "D": mr("−3 < x ≤ 5"),
     "ans":"D"},
    # 5 — problem 50
    {"num":5,"src":"50",
     "q": mr("3 < 2(x+7) − (4−x) ≤ 12"),
     "A": mg(mr("−"), mf("7","3"), mr(" < x ≤ "), mf("2","3")),
     "B": mg(mr("−"), mf("7","3"), mr(" ≤ x ≤ "), mf("2","3")),
     "C": mg(mr("−"), mf("7","3"), mr(" < x < "), mf("2","3")),
     "D": mg(mr("−"), mf("5","3"), mr(" < x ≤ "), mf("2","3")),
     "ans":"A"},
    # 6 — problem 51
    {"num":6,"src":"51",
     "q": mr("3x + 7 < 5x − 1 < 7x − 9"),
     "A": mr("x > 3"),
     "B": mr("x > 5"),
     "C": mr("x > 4"),
     "D": mr("x ≥ 4"),
     "ans":"C"},
    # 7 — problem 52
    {"num":7,"src":"52",
     "q": mr("2x + 3 ≤ 4x + 7 ≤ x + 19"),
     "A": mr("−2 ≤ x ≤ 4"),
     "B": mr("−2 < x ≤ 4"),
     "C": mr("−2 ≤ x < 4"),
     "D": mr("−1 ≤ x ≤ 4"),
     "ans":"A"},
    # 8 — problem 53
    {"num":8,"src":"53",
     "q": mr("7x − 6(x+2) < 3x + 8 ≤ 6x − 2(6−x)"),
     "A": mr("x > 4"),
     "B": mr("x ≥ 4"),
     "C": mr("x ≥ 5"),
     "D": mr("x > 3"),
     "ans":"B"},
    # 9 — problem 54
    {"num":9,"src":"54",
     "q": mg(mf("2x+1","3"), mr(" + "), mf("3x+1","6"), mr(" > "), mf("3x+7","4"), mr(" > "), mf("3x−7","2")),
     "A": mr("3 < x < 8"),
     "B": mr("3 ≤ x < 7"),
     "C": mr("3 < x ≤ 7"),
     "D": mr("3 < x < 7"),
     "ans":"D"},
    # 10 — problem 55
    {"num":10,"src":"55",
     "q": mg(mr("(x+2)(x+3) ≤ x"), msup("","2"), mr("+4x+6 ≤ (x+4)(x+5)")),
     "A": mg(mr("−"), mf("14","5"), mr(" ≤ x ≤ 0")),
     "B": mr("−3 ≤ x ≤ 0"),
     "C": mg(mr("−"), mf("14","5"), mr(" ≤ x ≤ 1")),
     "D": mg(mr("−"), mf("14","5"), mr(" < x ≤ 0")),
     "ans":"A"},
]


QUESTIONS += [
    # 11 — problem 56
    {"num":11,"src":"56",
     "q": mr("2x + 3(x−2) < −1   [eng katta butun yechim]"),
     "A": mr("−1"),
     "B": mr("1"),
     "C": mr("0"),
     "D": mr("2"),
     "ans":"C"},
    # 12 — problem 57
    {"num":12,"src":"57",
     "q": mr("2x − 3(4−x) > 5   [eng kichik butun yechim]"),
     "A": mr("3"),
     "B": mr("5"),
     "C": mr("4"),
     "D": mr("2"),
     "ans":"C"},
    # 13 — problem 58
    {"num":13,"src":"58",
     "q": mr("3(x−1) + 2(x−9) > 0   [eng kichik natural yechim]"),
     "A": mr("4"),
     "B": mr("6"),
     "C": mr("3"),
     "D": mr("5"),
     "ans":"D"},
    # 14 — problem 59
    {"num":14,"src":"59",
     "q": mr("2(x−5) + 3(x−7) > 0   [eng kichik natural yechim]"),
     "A": mr("6"),
     "B": mr("8"),
     "C": mr("5"),
     "D": mr("7"),
     "ans":"D"},
    # 15 — problem 60
    {"num":15,"src":"60",
     "q": mr("2x + 3(x−5) < 8   [natural yechimlari yig'indisi]"),
     "A": mr("6"),
     "B": mr("15"),
     "C": mr("10"),
     "D": mr("8"),
     "ans":"C"},
    # 16 — problem 61
    {"num":16,"src":"61",
     "q": mg(mf("3x+1","2"), mr(" − 4 < 2x − "), mf("x−3","4"), mr("   [eng kichik butun yechim]")),
     "A": mr("−15"),
     "B": mr("−17"),
     "C": mr("−16"),
     "D": mr("−18"),
     "ans":"C"},
    # 17 — problem 62
    {"num":17,"src":"62",
     "q": mg(mf("5","4"), mr("(x+1) + "), mf("3","4"), mr("(x−2) ≥ 3(x−1)   [natural yechimlari yig'indisi]")),
     "A": mr("6"),
     "B": mr("1"),
     "C": mr("2"),
     "D": mr("3"),
     "ans":"D"},
    # 18 — problem 63
    {"num":18,"src":"63",
     "q": mg(mf("3x−2","4"), mr(" − "), mf("2x+1","3"), mr(" ≤ 1   [eng katta natural yechim]")),
     "A": mr("20"),
     "B": mr("21"),
     "C": mr("23"),
     "D": mr("22"),
     "ans":"D"},
    # 19 — problem 64
    {"num":19,"src":"64",
     "q": mg(mr("(2x−3)"), msup("","2"), mr(" − (2x−7)(2x+7) > 0   [natural yechimlar soni]")),
     "A": mr("5"),
     "B": mr("3"),
     "C": mr("2"),
     "D": mr("4"),
     "ans":"D"},
    # 20 — problem 65
    {"num":20,"src":"65",
     "q": mg(mr("(3x−5)"), msup("","2"), mr(" − (3x−8)(3x+8) > 0   [natural yechimlar soni]")),
     "A": mr("3"),
     "B": mr("2"),
     "C": mr("4"),
     "D": mr("1"),
     "ans":"B"},
]


QUESTIONS += [
    # 21 — problem 66
    {"num":21,"src":"66",
     "q": mg(mr("(x−3)"), msup("","2"), mr(" − (x−4)(x+4) − 3(x+12) < 16   [eng kichik butun yechim]")),
     "A": mr("−3"),
     "B": mr("−1"),
     "C": mr("−2"),
     "D": mr("0"),
     "ans":"C"},
    # 22 — problem 67
    {"num":22,"src":"67",
     "q": mg(mr("(x−4)"), msup("","2"), mr(" − (x−2)(x+2) − 4(x−3) < 8   [eng kichik butun yechim]")),
     "A": mr("2"),
     "B": mr("4"),
     "C": mr("1"),
     "D": mr("3"),
     "ans":"D"},
    # 23 — problem 68
    {"num":23,"src":"68",
     "q": mg(mr("(x+3)"), msup("","3"), mr(" − (x−3)"), msup("","3"), mr(" − 2(3x+2)"), msup("","2"), mr(" + 2 > 0   [natural sonlar soni]")),
     "A": mr("2"),
     "B": mr("3"),
     "C": mr("0"),
     "D": mr("1"),
     "ans":"D"},
    # 24 — problem 69
    {"num":24,"src":"69",
     "q": mg(mr("(x+1)"), msup("","3"), mr(" − x(x+1)"), msup("","2"), mr(" − x"), msup("","2"), mr(" + 3 > 0   [eng kichik butun yechim]")),
     "A": mr("−2"),
     "B": mr("0"),
     "C": mr("−1"),
     "D": mr("1"),
     "ans":"C"},
    # 25 — problem 70
    {"num":25,"src":"70",
     "q": mr("{ 2(x−1) ≤ 6  va  3(x−2) ≥ −3   [butun yechimlari yig'indisi]"),
     "A": mr("6"),
     "B": mr("8"),
     "C": mr("12"),
     "D": mr("10"),
     "ans":"D"},
    # 26 — problem 71
    {"num":26,"src":"71",
     "q": mr("{ 5(x−1) ≤ 10  va  4(x−2) ≥ −8   [butun yechimlari yig'indisi]"),
     "A": mr("4"),
     "B": mr("8"),
     "C": mr("6"),
     "D": mr("10"),
     "ans":"C"},
    # 27 — problem 72
    {"num":27,"src":"72",
     "q": mr("−7 ≤ 1 − 4x ≤ 13   [butun yechimlari yig'indisi]"),
     "A": mr("−3"),
     "B": mr("3"),
     "C": mr("0"),
     "D": mr("−6"),
     "ans":"A"},
    # 28 — problem 73
    {"num":28,"src":"73",
     "q": mr("−3 ≤ 7 − 2x ≤ 11   [butun yechimlari yig'indisi]"),
     "A": mr("8"),
     "B": mr("16"),
     "C": mr("12"),
     "D": mr("10"),
     "ans":"C"},
    # 29 — problem 74
    {"num":29,"src":"74",
     "q": mg(mf("2x²−5x+3","6"), mr(" − "), mf("4−x","12"), mr(" ≥ "), mf("15+x²","3"), mr(" − "), mf("1−2x","9"), mr("   [eng katta butun son]")),
     "A": mr("−4"),
     "B": mr("−6"),
     "C": mr("−3"),
     "D": mr("−5"),
     "ans":"D"},
    # 30 — problem 75
    {"num":30,"src":"75",
     "q": mg(mr("(x+4)"), msup("","2"), mr(" − (x−10)"), msup("","2"), mr(" ≤ 140   [eng katta butun son]")),
     "A": mr("7"),
     "B": mr("9"),
     "C": mr("6"),
     "D": mr("8"),
     "ans":"D"},
]


QUESTIONS += [
    # 31 — problem 76
    {"num":31,"src":"76",
     "q": mr("6x + 12 − 3(x+2) ifoda qachon manfiy?"),
     "A": mr("x < −3"),
     "B": mr("x > −2"),
     "C": mr("x < −2"),
     "D": mr("x ≤ −2"),
     "ans":"C"},
    # 32 — problem 77
    {"num":32,"src":"77",
     "q": mr("3x − 8 ning qiymati 4x + 7 dan katta, x ning qiymatlari:"),
     "A": mr("x < −16"),
     "B": mr("x > −15"),
     "C": mr("x < −14"),
     "D": mr("x < −15"),
     "ans":"D"},
    # 33 — problem 78
    {"num":33,"src":"78",
     "q": mr("3x+7 va 2(x+3) larning yig'indisi 3 dan katta, x ning qiymatlari:"),
     "A": mr("x > −1"),
     "B": mr("x > −3"),
     "C": mr("x > 0"),
     "D": mr("x > −2"),
     "ans":"D"},
    # 34 — problem 79
    {"num":34,"src":"79",
     "q": mr("n natural soni nechta qiymatida 5n+2 soni 40 dan katta, lekin 90 dan kichik?"),
     "A": mr("9"),
     "B": mr("11"),
     "C": mr("8"),
     "D": mr("10"),
     "ans":"D"},
    # 35 — problem 80
    {"num":35,"src":"80",
     "q": mr("n natural sonining nechta qiymatida 8n+5 soni ikki xonali son bo'ladi?"),
     "A": mr("12"),
     "B": mr("10"),
     "C": mr("9"),
     "D": mr("11"),
     "ans":"D"},
    # 36 — problem 81
    {"num":36,"src":"81",
     "q": mr("n natural sonining nechta qiymatida 7n+3 soni ikki xonali son bo'ladi?"),
     "A": mr("12"),
     "B": mr("14"),
     "C": mr("13"),
     "D": mr("11"),
     "ans":"C"},
    # 37 — problem 82
    {"num":37,"src":"82",
     "q": mr("Agar 3 < a < 7 va 4 < b < 7 bo'lsa, 2a+3b ning qiymatlar sohasi:"),
     "A": mr("(18 ; 35)"),
     "B": mr("(20 ; 35)"),
     "C": mr("[18 ; 35]"),
     "D": mr("(18 ; 37)"),
     "ans":"A"},
    # 38 — problem 83
    {"num":38,"src":"83",
     "q": mr("Agar 1 < a ≤ 9 va 3 < b ≤ 7 bo'lsa, 3a+b ning qiymatlar sohasi:"),
     "A": mr("(6 ; 34]"),
     "B": mr("[6 ; 34]"),
     "C": mr("(6 ; 34)"),
     "D": mr("[7 ; 34]"),
     "ans":"A"},
]

print(f"Jami savollar: {len(QUESTIONS)}")


# ─── XML template strings ─────────────────────────────────────────────────────

CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

WORD_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>'''

SETTINGS_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W}">
  <w:defaultTabStop w:val="720"/>
</w:settings>'''

STYLES_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="22"/><w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="80"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>'''


# ─── Document XML builder ─────────────────────────────────────────────────────

def text_run_xml(text, bold=False, size=22, font="Times New Roman"):
    """Return XML string for a plain text run."""
    safe = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    b_tag = "<w:b/><w:bCs/>" if bold else ""
    return (f'<w:r xmlns:w="{W}">'
            f'<w:rPr>{b_tag}'
            f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
            f'</w:rPr>'
            f'<w:t xml:space="preserve">{safe}</w:t>'
            f'</w:r>')

def para_xml(inner_runs_xml, spacing_after=80, jc="left", indent_left=0):
    """Return XML string for a paragraph."""
    jc_tag = f'<w:jc w:val="{jc}"/>' if jc != "left" else ""
    ind_tag = f'<w:ind w:left="{indent_left}"/>' if indent_left else ""
    return (f'<w:p xmlns:w="{W}" xmlns:m="{M}">'
            f'<w:pPr><w:spacing w:after="{spacing_after}"/>{jc_tag}{ind_tag}</w:pPr>'
            f'{inner_runs_xml}'
            f'</w:p>')

def separator_para_xml():
    return (f'<w:p xmlns:w="{W}">'
            f'<w:pPr><w:spacing w:after="40"/>'
            f'<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="BBBBBB"/></w:pBdr>'
            f'</w:pPr></w:p>')

def page_break_xml():
    return (f'<w:p xmlns:w="{W}"><w:r><w:br w:type="page"/></w:r></w:p>')

def table_row_xml(cells, col_w=2340):
    tcs = ""
    for cell in cells:
        tcs += (f'<w:tc xmlns:w="{W}">'
                f'<w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>'
                f'<w:tcBorders>'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                f'</w:tcBorders></w:tcPr>'
                f'<w:p xmlns:w="{W}"><w:pPr><w:jc w:val="center"/><w:spacing w:after="40"/></w:pPr>'
                f'{text_run_xml(cell, bold=True, size=22)}'
                f'</w:p></w:tc>')
    return f'<w:tr xmlns:w="{W}">{tcs}</w:tr>'


def build_document(questions):
    """Build the complete word/document.xml string."""
    
    body_parts = []
    
    # ── Title
    body_parts.append(para_xml(
        text_run_xml("24. CHIZIQLI TENGSIZLIK VA TENGSIZLIKLAR SISTEMASI", bold=True, size=28),
        spacing_after=160, jc="center"))
    body_parts.append(para_xml(
        text_run_xml("Qo'shtengsizliklarni yeching (46–83-masalalar)", bold=True, size=24),
        spacing_after=200, jc="center"))
    body_parts.append(para_xml(text_run_xml(""), spacing_after=40))
    
    opt_letters = ["A","B","C","D"]
    
    for q in questions:
        num    = q["num"]
        q_xml  = q["q"]        # OMML inner string
        
        # Question line: "N.  [math]"
        q_line = (
            text_run_xml(f"{num}.   ", bold=True, size=22) +
            omml_inline(q_xml)
        )
        body_parts.append(para_xml(q_line, spacing_after=60))
        
        # Options line: "  A)  [math]   B)  [math]  ..."
        opt_line = ""
        for letter in opt_letters:
            opt_line += text_run_xml(f"    {letter})  ", bold=True, size=22)
            opt_line += omml_inline(q[letter])
        body_parts.append(para_xml(opt_line, spacing_after=130))
        
        # Separator
        body_parts.append(separator_para_xml())
    
    # ── Answer Key page
    body_parts.append(page_break_xml())
    body_parts.append(para_xml(
        text_run_xml("JAVOBLAR KALITI  —  ANSWER KEY", bold=True, size=28),
        spacing_after=80, jc="center"))
    body_parts.append(para_xml(
        text_run_xml("24. Chiziqli tengsizlik va tengsizliklar sistemasi", bold=True, size=22),
        spacing_after=160, jc="center"))
    body_parts.append(para_xml(text_run_xml(""), spacing_after=60))
    
    # Answer key table
    COLS = 4
    rows = []
    row  = []
    for q in questions:
        row.append(f"{q['num']}) — {q['ans']}")
        if len(row) == COLS:
            rows.append(row); row = []
    if row:
        while len(row) < COLS: row.append("")
        rows.append(row)
    
    tbl_rows_xml = "".join(table_row_xml(r) for r in rows)
    tbl_xml = (f'<w:tbl xmlns:w="{W}">'
               f'<w:tblPr>'
               f'<w:tblW w:w="9360" w:type="dxa"/>'
               f'<w:tblBorders>'
               f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               f'</w:tblBorders></w:tblPr>'
               f'{tbl_rows_xml}</w:tbl>')
    body_parts.append(tbl_xml)
    
    # Distribution note
    dist = Counter(q['ans'] for q in questions)
    body_parts.append(para_xml(text_run_xml(""), spacing_after=60))
    body_parts.append(para_xml(
        text_run_xml(f"Javoblar taqsimoti:  A = {dist['A']}  |  B = {dist['B']}  |  C = {dist['C']}  |  D = {dist['D']}", size=20),
        spacing_after=80, jc="center"))
    
    # Section properties
    sect_xml = (f'<w:sectPr xmlns:w="{W}">'
                f'<w:pgSz w:w="12240" w:h="15840"/>'
                f'<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1700"'
                f'         w:header="709" w:footer="709" w:gutter="0"/>'
                f'</w:sectPr>')
    
    body_content = "\n".join(body_parts) + "\n" + sect_xml
    
    doc_xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               f'<w:document xmlns:w="{W}" xmlns:m="{M}"'
               f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
               f'<w:body>\n{body_content}\n</w:body>\n</w:document>')
    return doc_xml


def write_docx(output_path, questions):
    doc_xml = build_document(questions)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels",          RELS_XML)
        zf.writestr("word/_rels/document.xml.rels", WORD_RELS_XML)
        zf.writestr("word/document.xml",    doc_xml.encode("utf-8"))
        zf.writestr("word/styles.xml",      STYLES_XML)
        zf.writestr("word/settings.xml",    SETTINGS_XML)
    
    size_kb = round(zipfile.ZipFile(output_path).getinfo("word/document.xml").file_size / 1024, 1)
    print(f"✅  Fayl tayyor: {output_path}")
    print(f"   Savollar soni  : {len(questions)}")
    dist = Counter(q['ans'] for q in questions)
    print(f"   Javoblar: A={dist['A']}  B={dist['B']}  C={dist['C']}  D={dist['D']}")
    print(f"   document.xml hajmi: {size_kb} KB")


def rebalance_answers(questions):
    """
    Swap option positions so correct answers distribute ~evenly across A/B/C/D.
    Target: each letter appears roughly len(q)//4 times.
    The CONTENT of options stays the same, only which slot is 'correct' changes
    by rotating the four option values.
    """
    import random
    # Target sequence cycling A B C D A B C D ...
    target_seq = []
    for i in range(len(questions)):
        target_seq.append(["A","B","C","D"][i % 4])
    
    balanced = []
    for q, target_ans in zip(questions, target_seq):
        current_ans = q["ans"]
        if current_ans == target_ans:
            balanced.append(q)
            continue
        # Rotate options: move current correct answer to target slot
        opts = ["A","B","C","D"]
        # Build value map
        vals = {L: q[L] for L in opts}
        correct_val = vals[current_ans]
        target_val  = vals[target_ans]
        # Swap current_ans <-> target_ans
        new_q = dict(q)
        new_q[current_ans] = target_val
        new_q[target_ans]  = correct_val
        new_q["ans"] = target_ans
        balanced.append(new_q)
    return balanced


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "Chiziqli_Tengsizlik_Test.docx"
    balanced = rebalance_answers(QUESTIONS)
    dist = Counter(q['ans'] for q in balanced)
    print(f"Balanslangan taqsimot: A={dist['A']}  B={dist['B']}  C={dist['C']}  D={dist['D']}")
    write_docx(out, balanced)
