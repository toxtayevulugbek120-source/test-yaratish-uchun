#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional matematik test generatori
3. KO'PHADLAR (Polynomials)
Natija: Microsoft Word (.docx) OMML equation formatida
"""

import zipfile
import random
from collections import Counter

# ─── Namespaces ───────────────────────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# ─── OMML string helpers ─────────────────────────────────────────────────────

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

def omml_inline(inner):
    """Inline math"""
    return f'<m:oMath xmlns:m="{M}">{inner}</m:oMath>'



# ─── Ko'phadlar masalalari ───────────────────────────────────────────────────

QUESTIONS = [
    # 1-10: Ko'phadlar bo'yicha oddiy masalalar
    
    {"num":1, "q": mg(mr("1−x+"), msup("x","2"), mr("−"), msup("x","3"), mr("+"), msup("x","4"), mr("   ko'phadini x+1 ga bo'lganda")),
     "A": mg(msup("x","4"), mr("−"), msup("x","3"), mr("+"), msup("x","2"), mr("−x")),
     "B": mg(msup("x","3"), mr("−"), msup("x","2"), mr("+x−1")),
     "C": mg(msup("x","3"), mr("+"), msup("x","2"), mr("−x+1")),
     "D": mg(msup("x","4"), mr("+"), msup("x","3"), mr("−"), msup("x","2"), mr("+x")),
     "ans":"B"},
    
    {"num":2, "q": mg(msup("x","4"), mr("+2"), msup("x","2"), mr("+1")),
     "A": mg(mr("("), msup("x","2"), mr("+1)"), msup("","2")),
     "B": mg(mr("(x+1)"), msup("","4")),
     "C": mg(mr("("), msup("x","2"), mr("−1)"), msup("","2")),
     "D": mg(mr("(x−1)"), msup("","4")),
     "ans":"A"},
    
    {"num":3, "q": mg(msup("a","2"), mr("+10a+25")),
     "A": mg(mr("(a+5)"), msup("","2")),
     "B": mg(mr("(a−5)"), msup("","2")),
     "C": mg(mr("(a+10)"), msup("","2")),
     "D": mg(mr("a(a+10)+25")),
     "ans":"A"},
    
    {"num":4, "q": mg(msup("x","2"), mr("−6x+9")),
     "A": mg(mr("(x−3)"), msup("","2")),
     "B": mg(mr("(x+3)"), msup("","2")),
     "C": mg(mr("(x−6)(x+9)")),
     "D": mg(mr("x(x−6)+9")),
     "ans":"A"},
    
    {"num":5, "q": mg(mr("Ko'paytuvchilarga ajrating:   "), msup("a","4"), mr("−1")),
     "A": mg(mr("("), msup("a","2"), mr("−1)("), msup("a","2"), mr("+1)")),
     "B": mg(mr("(a−1)"), msup("","4")),
     "C": mg(mr("(a−1)(a+1)("), msup("a","2"), mr("+1)")),
     "D": mg(mr("(a−1)(a+1)"), msup("","2")),
     "ans":"C"},
]



QUESTIONS += [
    # 11-30: O'rtacha qiyinlikdagi ko'phadlar
    
    {"num":6, "q": mg(mr("P(x) = "), msup("x","3"), mr("−2x+3   ko'phadining qiymatini x=−2 da toping")),
     "A": mr("−1"),
     "B": mr("3"),
     "C": mr("−5"),
     "D": mr("7"),
     "ans":"A"},
    
    {"num":7, "q": mg(mr("P(x) = "), msup("x","2"), mr("−2"), msup("x","4"), mr("+"), msup("2x","3"), mr("−1   to'liq kvadratli hadining koeffitsienti")),
     "A": mr("−2"),
     "B": mr("1"),
     "C": mr("2"),
     "D": mr("−1"),
     "ans":"A"},
    
    {"num":8, "q": mg(mr("P(x) = (x−3)"), msup("","2"), mr(" + (x+1)"), msup("","2"), mr("   ni standart shaklga keltiring")),
     "A": mg(mr("2"), msup("x","2"), mr("−4x+10")),
     "B": mg(mr("2"), msup("x","2"), mr("+4x+10")),
     "C": mg(msup("x","2"), mr("−4x+10")),
     "D": mg(mr("2"), msup("x","2"), mr("−4x+8")),
     "ans":"A"},
    
    {"num":9, "q": mg(mr("P(x) = "), msup("x","2"), mr("+ (x+1)"), msup("","2"), mr(" + (x+2)"), msup("","2")),
     "A": mg(mr("3"), msup("x","2"), mr("+6x+5")),
     "B": mg(mr("3"), msup("x","2"), mr("+5x+6")),
     "C": mg(mr("3"), msup("x","2"), mr("+6x+6")),
     "D": mg(msup("x","2"), mr("+6x+5")),
     "ans":"A"},
    
    {"num":10, "q": mg(mr("P(x) = (x−2)"), msup("","2"), mr(" − (2−x)"), msup("","2"), mr("   ifodani soddalashtiring")),
     "A": mr("0"),
     "B": mg(mr("2"), msup("x","2"), mr("−8x+8")),
     "C": mr("−4x+4"),
     "D": mr("4x−4"),
     "ans":"A"},
    
    # 11-20: Qo'shimcha ko'phadlar masalalari (PDF dan)
    
    {"num":11, "q": mg(mr("P(x) = (x+3)"), msup("","2"), mr(" − (x−2)"), msup("","2")),
     "A": mr("10x+5"),
     "B": mr("10x−5"),
     "C": mr("5x+10"),
     "D": mr("2x+5"),
     "ans":"A"},
    
    {"num":12, "q": mg(mr("P(x) = "), msup("(x+1)","2"), mr(" + "), msup("(x−1)","2"), mr(" + "), msup("x","2")),
     "A": mg(mr("3"), msup("x","2"), mr("+2")),
     "B": mg(mr("3"), msup("x","2"), mr("−2")),
     "C": mg(msup("x","2"), mr("+2")),
     "D": mg(mr("3"), msup("x","2")),
     "ans":"A"},
    
    # PDF page 1 - Qo'shimcha masalalar (13-30)
    {"num":13, "q": mg(mr("P(x) = (−3)"), msup("","4"), mr(" − 2×(−3)"), msup("","2"), mr(" + 2×(−3) + 21")),
     "A": mr("100"),
     "B": mr("81"),
     "C": mr("96"),
     "D": mr("108"),
     "ans":"A"},
    
    {"num":14, "q": mg(mr("P(x) = (−2)"), msup("","3"), mr(" + (−2)"), msup("","2"), mr(" − (−1)"), msup("","3"), mr(" − 2×7")),
     "A": mr("−19"),
     "B": mr("−17"),
     "C": mr("−21"),
     "D": mr("−15"),
     "ans":"A"},
    
    {"num":15, "q": mg(mr("P(x) = (x−2)"), msup("","2"), mr(" + (x−2)"), msup("","5"), mr("   standart shakli")),
     "A": mg(msup("x","5"), mr("−8"), msup("x","4"), mr("+24"), msup("x","3"), mr("−32"), msup("x","2"), mr("+16x+12")),
     "B": mg(msup("x","5"), mr("+"), msup("x","2")),
     "C": mg(msup("x","7")),
     "D": mg(msup("x","5"), mr("−2"), msup("x","2")),
     "ans":"A"},
    
    {"num":16, "q": mg(mr("P(x) = (−3)"), msup("","2"), mr(" + 2×7")),
     "A": mr("23"),
     "B": mr("5"),
     "C": mr("14"),
     "D": mr("−5"),
     "ans":"A"},
    
    {"num":17, "q": mg(mr("P(x) = 2×(−1)"), msup("","2"), mr(" + (−2)"), msup("","3"), mr(" − 3×(−1)"), msup("","2"), mr(" − 7")),
     "A": mr("−14"),
     "B": mr("−16"),
     "C": mr("−12"),
     "D": mr("−18"),
     "ans":"A"},
    
    {"num":18, "q": mg(mr("P(x) = (x+1)"), msup("","2"), mr(" + (x−1)"), msup("","2"), mr("   soddalashtiring")),
     "A": mg(mr("2"), msup("x","2"), mr("+2")),
     "B": mg(mr("2"), msup("x","2")),
     "C": mg(mr("2"), msup("x","2"), mr("−2")),
     "D": mg(msup("x","2"), mr("+2")),
     "ans":"A"},
    
    {"num":19, "q": mg(mr("P(x) = (x−y)"), msup("","2"), mr(" + (x+y)"), msup("","2")),
     "A": mg(mr("2"), msup("x","2"), mr("+2"), msup("y","2")),
     "B": mg(mr("2"), msup("x","2")),
     "C": mg(mr("2"), msup("y","2")),
     "D": mg(mr("4xy")),
     "ans":"A"},
    
    {"num":20, "q": mg(mr("P(x) = (2x−1)"), msup("","2"), mr(" − (2x−2)"), msup("","2")),
     "A": mr("−2x+3"),
     "B": mr("2x−3"),
     "C": mr("−2x−3"),
     "D": mr("4x−3"),
     "ans":"A"},
    
    # 21-30: O'rtacha qiyinlikdagi o'tizlta masalalar
    {"num":21, "q": mg(mr("P(x) = (x−3)"), msup("","2"), mr(" + (x−2)"), msup("","2"), mr(" − 2(x−3)(x−2)")),
     "A": mr("1"),
     "B": mr("0"),
     "C": mr("−1"),
     "D": mr("2"),
     "ans":"A"},
    
    {"num":22, "q": mg(mr("P(x) = (x+1)"), msup("","3"), mr(" − (x−2)"), msup("","3")),
     "A": mg(mr("9"), msup("x","2"), mr("+3x+9")),
     "B": mg(mr("9"), msup("x","2"), mr("−3x+9")),
     "C": mg(mr("9"), msup("x","2"), mr("+9")),
     "D": mg(mr("3"), msup("x","2"), mr("+3x+9")),
     "ans":"A"},
    
    {"num":23, "q": mg(mr("P(x) = "), msup("x","4"), mr("−16")),
     "A": mg(mr("(x−2)(x+2)("), msup("x","2"), mr("+4)")),
     "B": mg(mr("(x−2)"), msup("","4")),
     "C": mg(mr("(x−4)(x+4)")),
     "D": mg(mr("("), msup("x","2"), mr("−4)"), msup("","2")),
     "ans":"A"},
    
    {"num":24, "q": mg(mr("P(x) = "), msup("x","6"), mr("−64")),
     "A": mg(mr("(x−2)(x+2)("), msup("x","2"), mr("+2x+4)("), msup("x","2"), mr("−2x+4)")),
     "B": mg(mr("(x−2)"), msup("","6")),
     "C": mg(mr("(x−8)(x+8)")),
     "D": mg(mr("("), msup("x","3"), mr("−8)"), msup("","2")),
     "ans":"A"},
    
    {"num":25, "q": mg(mr("P(x) = "), msup("a","2"), mr("−"), msup("b","2"), mr("+2bc−"), msup("c","2")),
     "A": mg(mr("(a−b+c)(a+b−c)")),
     "B": mg(mr("(a−b−c)"), msup("","2")),
     "C": mg(mr("(a+b+c)(a−b−c)")),
     "D": mg(mr("(a−c)"), msup("","2"), mr("−"), msup("b","2")),
     "ans":"A"},
    
    {"num":26, "q": mg(mr("P(x) = 4"), msup("x","2"), mr("−4xy+"), msup("y","2"), mr("−9")),
     "A": mg(mr("(2x−y−3)(2x−y+3)")),
     "B": mg(mr("(2x−y)"), msup("","2"), mr("−9")),
     "C": mg(mr("(2x−y−3)"), msup("","2")),
     "D": mg(mr("4(x−y)"), msup("","2"), mr("−9")),
     "ans":"A"},
    
    {"num":27, "q": mg(mr("P(x) = "), msup("x","2"), mr("+10x+21")),
     "A": mg(mr("(x+3)(x+7)")),
     "B": mg(mr("(x+10)(x+21)")),
     "C": mg(mr("(x+1)(x+21)")),
     "D": mg(mr("(x+5)"), msup("","2"), mr("−4")),
     "ans":"A"},
    
    {"num":28, "q": mg(mr("P(x) = "), msup("x","2"), mr("−7x+12")),
     "A": mg(mr("(x−3)(x−4)")),
     "B": mg(mr("(x−1)(x−12)")),
     "C": mg(mr("(x−6)(x−2)")),
     "D": mg(mr("(x−7)(x+12)")),
     "ans":"A"},
    
    {"num":29, "q": mg(mr("P(x) = "), msup("x","2"), mr("+x−12")),
     "A": mg(mr("(x+4)(x−3)")),
     "B": mg(mr("(x+1)(x−12)")),
     "C": mg(mr("(x+6)(x−2)")),
     "D": mg(mr("(x+12)(x−1)")),
     "ans":"A"},
    
    {"num":30, "q": mg(mr("P(x) = "), msup("x","2"), mr("−5x−14")),
     "A": mg(mr("(x−7)(x+2)")),
     "B": mg(mr("(x−5)(x−14)")),
     "C": mg(mr("(x−14)(x+1)")),
     "D": mg(mr("(x+7)(x−2)")),
     "ans":"A"},
    
    # 31-50: PDF page 2 - Ko'phadlar (davomi)
    {"num":31, "q": mg(mr("P(x) = "), msup("x","3"), mr("+8")),
     "A": mg(mr("(x+2)("), msup("x","2"), mr("−2x+4)")),
     "B": mg(mr("(x+8)")),
     "C": mg(mr("(x+2)"), msup("","3")),
     "D": mg(mr("x("), msup("x","2"), mr("+8)")),
     "ans":"A"},
    
    {"num":32, "q": mg(mr("P(x) = "), msup("x","3"), mr("−27")),
     "A": mg(mr("(x−3)("), msup("x","2"), mr("+3x+9)")),
     "B": mg(mr("(x−27)")),
     "C": mg(mr("(x−3)"), msup("","3")),
     "D": mg(mr("x("), msup("x","2"), mr("−27)")),
     "ans":"A"},
    
    {"num":33, "q": mg(mr("P(x) = 8"), msup("x","3"), mr("+27")),
     "A": mg(mr("(2x+3)(4"), msup("x","2"), mr("−6x+9)")),
     "B": mg(mr("(2x+3)"), msup("","3")),
     "C": mg(mr("8x("), msup("x","2"), mr("+27)")),
     "D": mg(mr("(8x+27)")),
     "ans":"A"},
    
    {"num":34, "q": mg(mr("P(x) = "), msup("x","3"), mr("−"), msup("y","3"), mr("+"), msup("x","2"), mr("y−x"), msup("y","2")),
     "A": mg(mr("(x−y)("), msup("x","2"), mr("+xy+"), msup("y","2"), mr(")")),
     "B": mg(mr("(x−y)"), msup("","3")),
     "C": mg(mr("x"), msup("y","2"), mr("(x−y)")),
     "D": mg(mr("(x−y)("), msup("x","2"), mr("+"), msup("y","2"), mr(")")),
     "ans":"A"},
    
    {"num":35, "q": mg(mr("P(x) = "), msup("x","2"), mr("−4x+4−"), msup("y","2")),
     "A": mg(mr("(x−2−y)(x−2+y)")),
     "B": mg(mr("(x−2)"), msup("","2"), mr("−"), msup("y","2")),
     "C": mg(mr("(x−y−2)"), msup("","2")),
     "D": mg(mr("(x+y−2)"), msup("","2")),
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
        <w:sz w:val="24"/><w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>'''


# ─── Document XML builder ─────────────────────────────────────────────────────

def text_run_xml(text, bold=False, size=24, font="Times New Roman"):
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
            f'<w:pPr><w:spacing w:after="60"/></w:pPr>'
            f'{text_run_xml("", size=18)}'
            f'</w:p>')

def page_break_xml():
    return (f'<w:p xmlns:w="{W}"><w:r><w:br w:type="page"/></w:r></w:p>')

def table_row_xml(cells, col_w=2340):
    tcs = ""
    for cell in cells:
        tcs += (f'<w:tc xmlns:w="{W}">'
                f'<w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>'
                f'<w:tcBorders>'
                f'<w:top w:val="single" w:sz="4"/>'
                f'<w:left w:val="single" w:sz="4"/>'
                f'<w:bottom w:val="single" w:sz="4"/>'
                f'<w:right w:val="single" w:sz="4"/>'
                f'</w:tcBorders></w:tcPr>'
                f'<w:p xmlns:w="{W}"><w:pPr><w:jc w:val="center"/></w:pPr>'
                f'{text_run_xml(cell, bold=True, size=22)}'
                f'</w:p></w:tc>')
    return f'<w:tr xmlns:w="{W}">{tcs}</w:tr>'




def rebalance_answers(questions):
    """Shuffle correct answers to distribute evenly across A/B/C/D."""
    import random
    random.seed(42)
    target_seq = []
    for i in range(len(questions)):
        target_seq.append(["A","B","C","D"][i % 4])
    
    balanced = []
    for q, target_ans in zip(questions, target_seq):
        current_ans = q["ans"]
        if current_ans == target_ans:
            balanced.append(q)
            continue
        opts = ["A","B","C","D"]
        vals = {L: q[L] for L in opts}
        correct_val = vals[current_ans]
        target_val  = vals[target_ans]
        new_q = dict(q)
        new_q[current_ans] = target_val
        new_q[target_ans]  = correct_val
        new_q["ans"] = target_ans
        balanced.append(new_q)
    return balanced

def build_document(questions):
    """Build the complete word/document.xml string."""
    body_parts = []
    
    # Title
    body_parts.append(para_xml(
        text_run_xml("3. KO'PHADLAR (POLYNOMIALS)", bold=True, size=32),
        spacing_after=160, jc="center"))
    body_parts.append(para_xml(
        text_run_xml("Professional A/B/C/D Test", bold=False, size=26),
        spacing_after=120, jc="center"))
    body_parts.append(para_xml(
        text_run_xml("To'g'ri javobni tanlang.", bold=False, size=22, font="Times New Roman"),
        spacing_after=200, jc="center"))
    
    opt_letters = ["A","B","C","D"]
    
    for q in questions:
        num = q["num"]
        q_xml = q["q"]
        
        # Question line
        q_line = (
            text_run_xml(f"{num}.  ", bold=True, size=24) +
            omml_inline(q_xml)
        )
        body_parts.append(para_xml(q_line, spacing_after=80))
        
        # Options line
        for letter in opt_letters:
            opt_line = (
                text_run_xml(f"    {letter})  ", bold=True, size=24) +
                omml_inline(q[letter])
            )
            body_parts.append(para_xml(opt_line, spacing_after=50, indent_left=360))
        
        body_parts.append(separator_para_xml())
    
    # Answer Key page
    body_parts.append(page_break_xml())
    body_parts.append(para_xml(
        text_run_xml("JAVOBLAR KALITI  —  ANSWER KEY", bold=True, size=30),
        spacing_after=100, jc="center"))
    body_parts.append(para_xml(
        text_run_xml("3. Ko'phadlar", bold=True, size=24),
        spacing_after=180, jc="center"))
    
    # Answer key table
    COLS = 4
    rows = []
    row = []
    for q in questions:
        row.append(f"{q['num']}) — {q['ans']}")
        if len(row) == COLS:
            rows.append(row)
            row = []
    if row:
        while len(row) < COLS:
            row.append("")
        rows.append(row)
    
    tbl_rows_xml = "".join(table_row_xml(r) for r in rows)
    tbl_xml = (f'<w:tbl xmlns:w="{W}">'
               f'<w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr>'
               f'{tbl_rows_xml}</w:tbl>')
    body_parts.append(tbl_xml)
    
    # Distribution
    dist = Counter(q['ans'] for q in questions)
    body_parts.append(para_xml(text_run_xml(""), spacing_after=80))
    body_parts.append(para_xml(
        text_run_xml(f"Javoblar taqsimoti:  A={dist['A']}  |  B={dist['B']}  |  C={dist['C']}  |  D={dist['D']}", size=22),
        spacing_after=100, jc="center"))
    
    # Section properties
    sect_xml = (f'<w:sectPr xmlns:w="{W}">'
                f'<w:pgSz w:w="12240" w:h="15840"/>'
                f'<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1700"/>'
                f'</w:sectPr>')
    
    body_content = "\n".join(body_parts) + "\n" + sect_xml
    
    doc_xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               f'<w:document xmlns:w="{W}" xmlns:m="{M}">\n'
               f'<w:body>\n{body_content}\n</w:body>\n</w:document>')
    return doc_xml


def write_docx(output_path, questions):
    """Create the .docx file."""
    doc_xml = build_document(questions)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", RELS_XML)
        zf.writestr("word/_rels/document.xml.rels", WORD_RELS_XML)
        zf.writestr("word/document.xml", doc_xml.encode("utf-8"))
        zf.writestr("word/styles.xml", STYLES_XML)
        zf.writestr("word/settings.xml", SETTINGS_XML)
    
    print(f"✅  Fayl tayyor: {output_path}")
    print(f"   Savollar soni: {len(questions)}")
    dist = Counter(q['ans'] for q in questions)
    print(f"   Javoblar: A={dist['A']}  B={dist['B']}  C={dist['C']}  D={dist['D']}")

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "Kophadlar_Test.docx"
    
    print(f"Jami savollar soni: {len(QUESTIONS)}")
    balanced = rebalance_answers(QUESTIONS)
    
    dist_before = Counter(q['ans'] for q in QUESTIONS)
    dist_after = Counter(q['ans'] for q in balanced)
    
    print(f"Balanslangandan oldin: A={dist_before['A']}  B={dist_before['B']}  C={dist_before['C']}  D={dist_before['D']}")
    print(f"Balanslangandan keyin: A={dist_after['A']}  B={dist_after['B']}  C={dist_after['C']}  D={dist_after['D']}")
    
    write_docx(out, balanced)
    print(f"\n✅ Test muvaffaqiyatli yaratildi: {out}")
    print("   Microsoft Word'da ochib, matematik ifodalarni tekshiring!")
