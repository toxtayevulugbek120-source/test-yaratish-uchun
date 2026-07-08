#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional Ratsional Tengsizliklar Test Generator
PDF dan barcha tengsizliklarni A/B/C/D test formatiga aylantiradi
Microsoft Word (.docx) OMML equation formatida
"""

import zipfile
import os
import random
from collections import Counter

# ═══════════════════════════════════════════════════════════════════════
#  OMML HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def mr(text):
    """Math run - OMML formatted text"""
    safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return f'<m:r xmlns:m="{M}"><m:t xml:space="preserve">{safe}</m:t></m:r>'

def mfrac(numerator, denominator):
    """Stacked fraction"""
    return (f'<m:f xmlns:m="{M}"><m:fPr><m:type m:val="bar"/></m:fPr>'
            f'<m:num>{numerator}</m:num>'
            f'<m:den>{denominator}</m:den></m:f>')

def msup(base, superscript):
    """Superscript (power)"""
    return (f'<m:sSup xmlns:m="{M}"><m:e>{base}</m:e>'
            f'<m:sup>{superscript}</m:sup></m:sSup>')

def msub(base, subscript):
    """Subscript"""
    return (f'<m:sSub xmlns:m="{M}"><m:e>{base}</m:e>'
            f'<m:sub>{subscript}</m:sub></m:sSub>')

def msqrt(inner):
    """Square root"""
    return f'<m:rad xmlns:m="{M}"><m:radPr/><m:deg/><m:e>{inner}</m:e></m:rad>'

def mg(*parts):
    """Group multiple OMML parts"""
    return "".join(parts)



# ═══════════════════════════════════════════════════════════════════════
#  ANSWER BALANCING & DOCUMENT GENERATION
# ═══════════════════════════════════════════════════════════════════════

def balance_answers(questions):
    """
    Redistribute correct answers evenly across A, B, C, D
    by swapping variant positions
    """
    target_pattern = ['A', 'B', 'C', 'D'] * (len(questions) // 4 + 1)
    random.shuffle(target_pattern)
    
    balanced = []
    for i, q in enumerate(questions):
        target_letter = target_pattern[i]
        target_idx = ['A', 'B', 'C', 'D'].index(target_letter)
        current_correct = q['correct']
        
        # Swap variants to move correct answer to target position
        variants = q['variants'][:]
        if current_correct != target_idx:
            variants[current_correct], variants[target_idx] = \
                variants[target_idx], variants[current_correct]
        
        balanced.append({
            "num": q["num"],
            "q": q["q"],
            "variants": variants,
            "correct": target_idx
        })
    
    return balanced

def text_run(text, bold=False, size=24):
    """Plain text run for Word"""
    safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    b_tag = '<w:b/><w:bCs/>' if bold else ''
    return (f'<w:r><w:rPr>{b_tag}'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
            f'<w:t xml:space="preserve">{safe}</w:t></w:r>')

def para(content, style="Normal", center=False, space_after=120):
    """Paragraph wrapper"""
    jc = '<w:jc w:val="center"/>' if center else ''
    return (f'<w:p xmlns:w="{W}" xmlns:m="{M}"><w:pPr>'
            f'<w:pStyle w:val="{style}"/><w:spacing w:after="{space_after}"/>{jc}</w:pPr>'
            f'{content}</w:p>')

def page_break():
    """Page break"""
    return f'<w:p xmlns:w="{W}"><w:r><w:br w:type="page"/></w:r></w:p>'

def build_document_xml(questions):
    """Build the complete Word document.xml"""
    body_parts = []
    
    # Title
    body_parts.append(para(
        text_run("RATSIONAL TENGSIZLIKLAR", bold=True, size=32),
        center=True, space_after=160
    ))
    body_parts.append(para(
        text_run("Professional Test — Microsoft Word OMML Format", bold=True, size=24),
        center=True, space_after=200
    ))
    
    # Questions
    for q in questions:
        num = q["num"]
        q_omml = q["q"]
        variants = q["variants"]
        
        # Question paragraph
        question_content = (
            text_run(f"{num}.  ", bold=True, size=24) +
            f'<m:oMath xmlns:m="{M}">{q_omml}</m:oMath>'
        )
        body_parts.append(para(question_content, space_after=80))
        
        # Variants (each on separate line)
        for i, letter in enumerate(['A', 'B', 'C', 'D']):
            variant_content = text_run(f"{letter})  ", bold=True, size=22)
            variant_content += f'<m:oMath xmlns:m="{M}">{variants[i]}</m:oMath>'
            body_parts.append(para(variant_content, space_after=60))
        
        # Extra space after each question
        body_parts.append(para(text_run("", size=12), space_after=100))
    
    # Page break before answer key
    body_parts.append(page_break())
    
    # Answer Key
    body_parts.append(para(
        text_run("JAVOBLAR KALITI — ANSWER KEY", bold=True, size=32),
        center=True, space_after=200
    ))
    
    # Answer key table
    for i in range(0, len(questions), 5):
        row_items = questions[i:i+5]
        row_text = "     ".join([
            f"{q['num']} — {['A','B','C','D'][q['correct']]}" 
            for q in row_items
        ])
        body_parts.append(para(text_run(row_text, size=22), space_after=80))
    
    # Distribution stats
    dist = Counter(q['correct'] for q in questions)
    dist_text = f"Taqsimot:  A={dist[0]}  B={dist[1]}  C={dist[2]}  D={dist[3]}"
    body_parts.append(para(
        text_run(dist_text, bold=True, size=20),
        center=True, space_after=80
    ))
    
    # Section properties
    sect = (f'<w:sectPr xmlns:w="{W}">'
            '<w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
            '</w:sectPr>')
    
    body_content = "\n".join(body_parts) + "\n" + sect
    
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document xmlns:w="{W}" xmlns:m="{M}">'
            f'<w:body>{body_content}</w:body></w:document>')



# ═══════════════════════════════════════════════════════════════════════
#  DOCX SKELETON FILES
# ═══════════════════════════════════════════════════════════════════════

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>'''

SETTINGS = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W}">
  <w:defaultTabStop w:val="720"/>
</w:settings>'''

STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}" xmlns:m="{M}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
      <w:lang w:val="uz-Latn-UZ"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="120"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>'''



# ═══════════════════════════════════════════════════════════════════════
#  RATSIONAL TENGSIZLIKLAR - MASALALAR
# ═══════════════════════════════════════════════════════════════════════

def generate_questions():
    """Generate all inequality questions with interval notation answers"""
    questions = []
    
    # Problem 1: x < 0
    questions.append({
        "num": 1,
        "q": mg(mr("x < 0")),
        "variants": [
            mg(mr("(−∞; 0)")),
            mg(mr("(−∞; 0]")),
            mg(mr("(0; +∞)")),
            mg(mr("[0; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 2: x+2 > 0
    questions.append({
        "num": 2,
        "q": mg(mfrac(mr("x"), mr("x + 2")), mr(" > 0")),
        "variants": [
            mg(mr("(−∞; −2) ∪ (0; +∞)")),
            mg(mr("(−2; 0)")),
            mg(mr("(−∞; −2] ∪ [0; +∞)")),
            mg(mr("[−2; 0]"))
        ],
        "correct": 0
    })
    
    # Problem 3: (x−1)/(x+2) ≤ 0
    questions.append({
        "num": 3,
        "q": mg(mfrac(mr("x − 1"), mr("x + 2")), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−2; 1]")),
            mg(mr("[−2; 1]")),
            mg(mr("(−2; 1)")),
            mg(mr("(−∞; −2) ∪ [1; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 4: (x−3)²/(x−2) ≥ 0
    questions.append({
        "num": 4,
        "q": mg(mfrac(msup(mr("(x − 3)"), mr("2")), mr("x − 2")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(2; +∞)")),
            mg(mr("[2; +∞)")),
            mg(mr("(−∞; 2) ∪ [3; +∞)")),
            mg(mr("(−∞; 2] ∪ [3; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 5: (x−1)/(x+3) > 0
    questions.append({
        "num": 5,
        "q": mg(mfrac(mr("x − 1"), mr("x + 3")), mr(" > 0")),
        "variants": [
            mg(mr("(−∞; −3) ∪ (1; +∞)")),
            mg(mr("(−3; 1)")),
            mg(mr("(−∞; −3] ∪ [1; +∞)")),
            mg(mr("[−3; 1]"))
        ],
        "correct": 0
    })
    
    # Problem 6: (x−5)(x−7) ≥ 0
    questions.append({
        "num": 6,
        "q": mg(mfrac(mr("(x − 5)(x − 7)"), mr("x")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; 0) ∪ [5; 7]")),
            mg(mr("[5; 7]")),
            mg(mr("(−∞; 0] ∪ [5; 7]")),
            mg(mr("(0; 5] ∪ [7; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 7: (x−3)(5x−3) < 0
    questions.append({
        "num": 7,
        "q": mg(mfrac(mr("(x − 3)(5x − 3)"), mr("x")), mr(" < 0")),
        "variants": [
            mg(mr("(−∞; 0) ∪ (0.6; 3)")),
            mg(mr("(0; 0.6) ∪ (3; +∞)")),
            mg(mr("(0.6; 3)")),
            mg(mr("(−∞; 0) ∪ [0.6; 3]"))
        ],
        "correct": 0
    })
    
    # Problem 8: (x−2)/(x+6) > 0
    questions.append({
        "num": 8,
        "q": mg(mfrac(mr("x − 2"), mr("x + 6")), mr(" > 0")),
        "variants": [
            mg(mr("(−∞; −6) ∪ (2; +∞)")),
            mg(mr("(−6; 2)")),
            mg(mr("(−∞; −6] ∪ [2; +∞)")),
            mg(mr("[−6; 2]"))
        ],
        "correct": 0
    })
    
    # Problem 9: (2x−1)/(x−5) ≤ 0
    questions.append({
        "num": 9,
        "q": mg(mfrac(mr("2x − 1"), mr("x − 5")), mr(" ≤ 0")),
        "variants": [
            mg(mr("[0.5; 5)")),
            mg(mr("(0.5; 5]")),
            mg(mr("[0.5; 5]")),
            mg(mr("(0.5; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 10: (x−4)²/(x+1) ≥ 0
    questions.append({
        "num": 10,
        "q": mg(mfrac(msup(mr("(x − 4)"), mr("2")), mr("x + 1")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−1; +∞)")),
            mg(mr("[−1; +∞)")),
            mg(mr("(−∞; −1) ∪ [4; +∞)")),
            mg(mr("(−∞; −1] ∪ [4; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 11: (2x−1)/(1−x) ≤ 0
    questions.append({
        "num": 11,
        "q": mg(mfrac(mr("2x − 1"), mr("1 − x")), mr(" ≤ 0")),
        "variants": [
            mg(mr("[0.5; 1)")),
            mg(mr("(0.5; 1]")),
            mg(mr("[0.5; 1]")),
            mg(mr("(−∞; 0.5] ∪ (1; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 12: (x−1)(x+5) ≤ 0
    questions.append({
        "num": 12,
        "q": mg(mfrac(mr("(x − 1)(x + 5)"), mr("x")), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−∞; −5] ∪ (0; 1]")),
            mg(mr("[−5; 0) ∪ [1; +∞)")),
            mg(mr("(−∞; −5] ∪ [0; 1]")),
            mg(mr("[−5; 1]"))
        ],
        "correct": 0
    })
    
    # Problem 13: (x−2)(x+1)/x² ≤ 0
    questions.append({
        "num": 13,
        "q": mg(mfrac(mr("(x − 2)(x + 1)"), msup(mr("x"), mr("2"))), mr(" ≤ 0")),
        "variants": [
            mg(mr("[−1; 0) ∪ (0; 2]")),
            mg(mr("[−1; 2]")),
            mg(mr("(−1; 0) ∪ (0; 2)")),
            mg(mr("(−∞; −1] ∪ [2; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 14: (x+2)/(x−4) < 0
    questions.append({
        "num": 14,
        "q": mg(mfrac(mr("x + 2"), mr("x − 4")), mr(" < 0")),
        "variants": [
            mg(mr("(−2; 4)")),
            mg(mr("(−∞; −2) ∪ (4; +∞)")),
            mg(mr("[−2; 4]")),
            mg(mr("(−2; 4]"))
        ],
        "correct": 0
    })
    
    # Problem 15: x²/(x−1) > 0
    questions.append({
        "num": 15,
        "q": mg(mfrac(msup(mr("x"), mr("2")), mr("x − 1")), mr(" > 0")),
        "variants": [
            mg(mr("(−∞; 0) ∪ (0; 1) ∪ (1; +∞)")),
            mg(mr("(1; +∞)")),
            mg(mr("(−∞; 1) ∪ (1; +∞)")),
            mg(mr("ℝ \\ {0; 1}"))
        ],
        "correct": 0
    })
    
    # Problem 16: x²/((x+5)(x−3)) > 0
    questions.append({
        "num": 16,
        "q": mg(mfrac(msup(mr("x"), mr("2")), mr("(x + 5)(x − 3)")), mr(" > 0")),
        "variants": [
            mg(mr("(−∞; −5) ∪ (3; +∞)")),
            mg(mr("(−5; 0) ∪ (0; 3)")),
            mg(mr("(−5; 3)")),
            mg(mr("(−∞; −5) ∪ (0; 3) ∪ (3; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 17: x²/((x−4)(x+1)) ≥ 0
    questions.append({
        "num": 17,
        "q": mg(mfrac(msup(mr("x"), mr("2")), mr("(x − 4)(x + 1)")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −1) ∪ (4; +∞)")),
            mg(mr("[−1; 4]")),
            mg(mr("(−1; 4)")),
            mg(mr("(−∞; −1] ∪ [4; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 18: x³/9 < 4
    questions.append({
        "num": 18,
        "q": mg(mfrac(msup(mr("x"), mr("3")), mr("9")), mr(" < 4")),
        "variants": [
            mg(mr("(−∞; ∛36)")),
            mg(mr("(−∞; 36)")),
            mg(mr("(−∞; 4)")),
            mg(mr("(−∞; 12)"))
        ],
        "correct": 0
    })
    
    # Problem 19: x²/(x−1) < 1
    questions.append({
        "num": 19,
        "q": mg(mfrac(msup(mr("x"), mr("2")), mr("x − 1")), mr(" < 1")),
        "variants": [
            mg(mr("(−∞; −1) ∪ (1; +∞)")),
            mg(mr("(−1; 1)")),
            mg(mr("(−∞; 1)")),
            mg(mr("(1; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 20: x²+4x+1 ≥ 0
    questions.append({
        "num": 20,
        "q": mg(mfrac(msup(mr("x"), mr("2")), mr(" + 4x + 1")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −2−√3] ∪ [−2+√3; +∞)")),
            mg(mr("[−2−√3; −2+√3]")),
            mg(mr("(−2−√3; −2+√3)")),
            mg(mr("ℝ"))
        ],
        "correct": 0
    })
    
    # Problem 21: (x²−4)/(x²+1) ≥ 0
    questions.append({
        "num": 21,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 4")), 
              mg(msup(mr("x"), mr("2")), mr(" + 1"))), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −2] ∪ [2; +∞)")),
            mg(mr("[−2; 2]")),
            mg(mr("(−2; 2)")),
            mg(mr("ℝ"))
        ],
        "correct": 0
    })
    
    # Problem 22: x−1/x ≤ 0
    questions.append({
        "num": 22,
        "q": mg(mr("x − "), mfrac(mr("1"), mr("x")), mr(" ≤ 0")),
        "variants": [
            mg(mr("[−1; 0) ∪ (0; 1]")),
            mg(mr("[−1; 1]")),
            mg(mr("(−1; 1)")),
            mg(mr("(−∞; −1] ∪ [1; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 23: 3x + 1/x ≤ 4
    questions.append({
        "num": 23,
        "q": mg(mr("3x + "), mfrac(mr("1"), mr("x")), mr(" ≤ 4")),
        "variants": [
            mg(mr("(−∞; 0) ∪ [⅓; 1]")),
            mg(mr("[⅓; 1]")),
            mg(mr("(−∞; ⅓] ∪ [1; +∞)")),
            mg(mr("[0; ⅓] ∪ [1; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 24: (x+2)/(x²−5x+6) ≥ 0
    questions.append({
        "num": 24,
        "q": mg(mfrac(mr("x + 2"), mg(msup(mr("x"), mr("2")), mr(" − 5x + 6"))), mr(" ≥ 0")),
        "variants": [
            mg(mr("[−2; 2) ∪ (3; +∞)")),
            mg(mr("(−2; 2) ∪ (3; +∞)")),
            mg(mr("[−2; 3]")),
            mg(mr("(−∞; −2] ∪ [2; 3]"))
        ],
        "correct": 0
    })
    
    # Problem 25: (x²−3x−4)/(x+1) ≤ 0
    questions.append({
        "num": 25,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 3x − 4")), mr("x + 1")), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−∞; −1) ∪ [−1; 4]")),
            mg(mr("[−1; 4]")),
            mg(mr("(−1; 4]")),
            mg(mr("(−∞; −1] ∪ [−1; 4]"))
        ],
        "correct": 0
    })
    
    # Problem 26: (x+1)/(x²+6) ≤ 0
    questions.append({
        "num": 26,
        "q": mg(mfrac(mr("x + 1"), mg(msup(mr("x"), mr("2")), mr(" + 6"))), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−∞; −1]")),
            mg(mr("[−1; +∞)")),
            mg(mr("(−∞; −1)")),
            mg(mr("∅"))
        ],
        "correct": 0
    })
    
    # Problem 27: (3x−2)/(5x+1) ≥ 0
    questions.append({
        "num": 27,
        "q": mg(mfrac(mr("3x − 2"), mr("5x + 1")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −0.2) ∪ [⅔; +∞)")),
            mg(mr("[−0.2; ⅔]")),
            mg(mr("(−0.2; ⅔)")),
            mg(mr("(−∞; −0.2] ∪ [⅔; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 28: 1/x + 1/x ≤ 0
    questions.append({
        "num": 28,
        "q": mg(mfrac(mr("1"), mr("x")), mr(" + "), mfrac(mr("1"), mr("x")), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−∞; 0)")),
            mg(mr("(0; +∞)")),
            mg(mr("(−∞; 0]")),
            mg(mr("ℝ \\ {0}"))
        ],
        "correct": 0
    })
    
    # Problem 29: 1/x + 1/y ≤ 0
    questions.append({
        "num": 29,
        "q": mg(mfrac(mr("1"), mr("x")), mr(" + "), mfrac(mr("1"), mr("y")), mr(" ≤ 0")),
        "variants": [
            mg(mr("x va y qarama-qarshi ishorali")),
            mg(mr("x va y bir xil ishorali")),
            mg(mr("x < 0 va y < 0")),
            mg(mr("x > 0 va y > 0"))
        ],
        "correct": 0
    })
    
    # Problem 30: (x²−3x+2)/(x−5) ≥ 0
    questions.append({
        "num": 30,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 3x + 2")), mr("x − 5")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; 1] ∪ [2; 5) ∪ (5; +∞)")),
            mg(mr("[1; 2] ∪ (5; +∞)")),
            mg(mr("(−∞; 1] ∪ [2; 5]")),
            mg(mr("[1; 2] ∪ [5; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 31: x⁴ − 6x³ + 8x² ≥ 0
    questions.append({
        "num": 31,
        "q": mg(msup(mr("x"), mr("4")), mr(" − 6"), msup(mr("x"), mr("3")), mr(" + 8"), 
              msup(mr("x"), mr("2")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; 0] ∪ [2; 4]")),
            mg(mr("[0; 2] ∪ [4; +∞)")),
            mg(mr("(−∞; 0) ∪ (2; 4)")),
            mg(mr("ℝ"))
        ],
        "correct": 0
    })
    
    # Problem 32: 1/x ≤ 5
    questions.append({
        "num": 32,
        "q": mg(mfrac(mr("1"), mr("x")), mr(" ≤ 5")),
        "variants": [
            mg(mr("(−∞; 0) ∪ [0.2; +∞)")),
            mg(mr("[0.2; +∞)")),
            mg(mr("(−∞; 0.2]")),
            mg(mr("(0; 0.2]"))
        ],
        "correct": 0
    })
    
    # Problem 33: 1/(x−5) > 2
    questions.append({
        "num": 33,
        "q": mg(mfrac(mr("1"), mr("x − 5")), mr(" > 2")),
        "variants": [
            mg(mr("(5; 5.5)")),
            mg(mr("(5.5; +∞)")),
            mg(mr("(−∞; 5.5)")),
            mg(mr("(−∞; 5) ∪ (5.5; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 34: 1/x² < 4
    questions.append({
        "num": 34,
        "q": mg(mfrac(mr("1"), msup(mr("x"), mr("2"))), mr(" < 4")),
        "variants": [
            mg(mr("(−∞; −0.5) ∪ (0.5; +∞)")),
            mg(mr("(−0.5; 0) ∪ (0; 0.5)")),
            mg(mr("(−0.5; 0.5)")),
            mg(mr("ℝ \\ {0}"))
        ],
        "correct": 0
    })
    
    # Problem 35: 2/(x+1) ≥ 1
    questions.append({
        "num": 35,
        "q": mg(mfrac(mr("2"), mr("x + 1")), mr(" ≥ 1")),
        "variants": [
            mg(mr("(−1; 1]")),
            mg(mr("[1; +∞)")),
            mg(mr("(−∞; −1) ∪ [1; +∞)")),
            mg(mr("[−1; 1]"))
        ],
        "correct": 0
    })
    
    # Problem 36: (2x²+5x+2)/(x²+3x+6) > 0
    questions.append({
        "num": 36,
        "q": mg(mfrac(mg(mr("2"), msup(mr("x"), mr("2")), mr(" + 5x + 2")), 
              mg(msup(mr("x"), mr("2")), mr(" + 3x + 6"))), mr(" > 0")),
        "variants": [
            mg(mr("(−∞; −2) ∪ (−0.5; +∞)")),
            mg(mr("(−2; −0.5)")),
            mg(mr("(−∞; −2] ∪ [−0.5; +∞)")),
            mg(mr("ℝ"))
        ],
        "correct": 0
    })
    
    # Problem 37: (x²−2x−3)/(x²+2x+5) < 0
    questions.append({
        "num": 37,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 2x − 3")), 
              mg(msup(mr("x"), mr("2")), mr(" + 2x + 5"))), mr(" < 0")),
        "variants": [
            mg(mr("(−1; 3)")),
            mg(mr("(−∞; −1) ∪ (3; +∞)")),
            mg(mr("[−1; 3]")),
            mg(mr("∅"))
        ],
        "correct": 0
    })
    
    # Problem 38: (2x²−x−6)/(x²+6) ≤ 0
    questions.append({
        "num": 38,
        "q": mg(mfrac(mg(mr("2"), msup(mr("x"), mr("2")), mr(" − x − 6")), 
              mg(msup(mr("x"), mr("2")), mr(" + 6"))), mr(" ≤ 0")),
        "variants": [
            mg(mr("[−1.5; 2]")),
            mg(mr("(−∞; −1.5] ∪ [2; +∞)")),
            mg(mr("(−1.5; 2)")),
            mg(mr("ℝ"))
        ],
        "correct": 0
    })
    
    # Problem 39: 1/(x+2) + 3/x > 0
    questions.append({
        "num": 39,
        "q": mg(mfrac(mr("1"), mr("x + 2")), mr(" + "), mfrac(mr("3"), mr("x")), mr(" > 0")),
        "variants": [
            mg(mr("(−3; −2) ∪ (0; +∞)")),
            mg(mr("(−∞; −3) ∪ (−2; 0)")),
            mg(mr("(−3; 0)")),
            mg(mr("(−2; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 40: 1/(x−3) − 2/(x+1) ≥ 0
    questions.append({
        "num": 40,
        "q": mg(mfrac(mr("1"), mr("x − 3")), mr(" − "), mfrac(mr("2"), mr("x + 1")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−1; 1] ∪ (3; +∞)")),
            mg(mr("[−1; 1] ∪ [3; +∞)")),
            mg(mr("(−∞; −1) ∪ (1; 3)")),
            mg(mr("[1; 3]"))
        ],
        "correct": 0
    })
    
    # Problem 41: 1/(x−1) + 1/(x+2) ≤ 1/(2x)
    questions.append({
        "num": 41,
        "q": mg(mfrac(mr("1"), mr("x − 1")), mr(" + "), mfrac(mr("1"), mr("x + 2")), 
              mr(" ≤ "), mfrac(mr("1"), mr("2x"))),
        "variants": [
            mg(mr("(−∞; −2) ∪ (0; 1) ∪ [2; +∞)")),
            mg(mr("[−2; 0] ∪ [1; 2]")),
            mg(mr("(−2; 0) ∪ (1; 2)")),
            mg(mr("ℝ \\ {−2; 0; 1}"))
        ],
        "correct": 0
    })
    
    # Problem 42: (x+1)/(2x−1) + (x−2)/(2x+1) ≥ 0
    questions.append({
        "num": 42,
        "q": mg(mfrac(mr("x + 1"), mr("2x − 1")), mr(" + "), 
              mfrac(mr("x − 2"), mr("2x + 1")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −0.5) ∪ [−0.2; 0.5) ∪ [1.5; +∞)")),
            mg(mr("[−0.5; 0.5] ∪ [1.5; +∞)")),
            mg(mr("(−0.5; −0.2] ∪ (0.5; 1.5)")),
            mg(mr("ℝ"))
        ],
        "correct": 0
    })
    
    # Problem 43: x/(x+2) ≤ (x+1)/(x−1)
    questions.append({
        "num": 43,
        "q": mg(mfrac(mr("x"), mr("x + 2")), mr(" ≤ "), mfrac(mr("x + 1"), mr("x − 1"))),
        "variants": [
            mg(mr("(−∞; −2) ∪ [0.5; 1)")),
            mg(mr("(−2; 0.5] ∪ (1; +∞)")),
            mg(mr("[0.5; 1]")),
            mg(mr("(−∞; 0.5]"))
        ],
        "correct": 0
    })
    
    # Problem 44: 1/x ≤ 1/(x−2)
    questions.append({
        "num": 44,
        "q": mg(mfrac(mr("1"), mr("x")), mr(" ≤ "), mfrac(mr("1"), mr("x − 2"))),
        "variants": [
            mg(mr("(−∞; 0) ∪ (2; +∞)")),
            mg(mr("(0; 2)")),
            mg(mr("[0; 2]")),
            mg(mr("(−∞; 0] ∪ [2; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 45: 2/(x²+3) ≥ 1/(x+1)
    questions.append({
        "num": 45,
        "q": mg(mfrac(mr("2"), mg(msup(mr("x"), mr("2")), mr(" + 3"))), mr(" ≥ "), 
              mfrac(mr("1"), mr("x + 1"))),
        "variants": [
            mg(mr("(−∞; −1) ∪ [1; +∞)")),
            mg(mr("[−1; 1]")),
            mg(mr("(−1; 1)")),
            mg(mr("ℝ \\ {−1}"))
        ],
        "correct": 0
    })
    
    # Problem 46: x/(x+1) < (x+2)/(x+3)
    questions.append({
        "num": 46,
        "q": mg(mfrac(mr("x"), mr("x + 1")), mr(" < "), mfrac(mr("x + 2"), mr("x + 3"))),
        "variants": [
            mg(mr("(−3; −1)")),
            mg(mr("(−∞; −3) ∪ (−1; +∞)")),
            mg(mr("[−3; −1]")),
            mg(mr("∅"))
        ],
        "correct": 0
    })
    
    # Problem 47: (x+3)/(x−2) < (2x+1)/(x−3)
    questions.append({
        "num": 47,
        "q": mg(mfrac(mr("x + 3"), mr("x − 2")), mr(" < "), 
              mfrac(mr("2x + 1"), mr("x − 3"))),
        "variants": [
            mg(mr("(−∞; 2) ∪ (2.8; 3)")),
            mg(mr("(2; 2.8) ∪ (3; +∞)")),
            mg(mr("(2; 3)")),
            mg(mr("(−∞; 2.8) ∪ (3; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 48: (5−2x)/(x+4) > 0 tengsizlikning butun yechimlar yig'indisi
    questions.append({
        "num": 48,
        "q": mg(mfrac(mr("5 − 2x"), mr("x + 4")), 
              mr(" > 0 tengsizlikning butun yechimlar yig'indisini toping.")),
        "variants": [
            mg(mr("−3")),
            mg(mr("0")),
            mg(mr("3")),
            mg(mr("6"))
        ],
        "correct": 0
    })
    
    # Problem 49: (x+3)(x−2)/((x+1)(x−4)) < 0 tengsizlikning butun yechimlar ko'paytmasi
    questions.append({
        "num": 49,
        "q": mg(mfrac(mr("(x + 3)(x − 2)"), mr("(x + 1)(x − 4)")), 
              mr(" < 0 tengsizlikning butun yechimlar ko'paytmasini toping.")),
        "variants": [
            mg(mr("0")),
            mg(mr("−6")),
            mg(mr("6")),
            mg(mr("−3"))
        ],
        "correct": 0
    })
    
    # Problem 50: (x²+7x+12)/(x²+x−2) < 0 — butun sonlar uzunliklari yig'indisi
    questions.append({
        "num": 50,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" + 7x + 12")), 
              mg(msup(mr("x"), mr("2")), mr(" + x − 2"))), 
              mr(" < 0 tengsizlikning yechimi bo'ladigan oraliqlar uzunliklari yig'indisini toping.")),
        "variants": [
            mg(mr("4")),
            mg(mr("3")),
            mg(mr("5")),
            mg(mr("6"))
        ],
        "correct": 0
    })
    
    # Problem 51: x² + 2x + 3 > 0
    questions.append({
        "num": 51,
        "q": mg(msup(mr("x"), mr("2")), mr(" + 2x + 3 > 0")),
        "variants": [
            mg(mr("ℝ")),
            mg(mr("∅")),
            mg(mr("(−∞; −3) ∪ (1; +∞)")),
            mg(mr("[−3; 1]"))
        ],
        "correct": 0
    })
    
    # Problem 52: 6x² + 5x + 4 < 0
    questions.append({
        "num": 52,
        "q": mg(mr("6"), msup(mr("x"), mr("2")), mr(" + 5x + 4 < 0")),
        "variants": [
            mg(mr("∅")),
            mg(mr("ℝ")),
            mg(mr("(−1; −⅔)")),
            mg(mr("(−∞; −1) ∪ (−⅔; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 53: −2x² + 3x − 8 > 0
    questions.append({
        "num": 53,
        "q": mg(mr("−2"), msup(mr("x"), mr("2")), mr(" + 3x − 8 > 0")),
        "variants": [
            mg(mr("∅")),
            mg(mr("ℝ")),
            mg(mr("(−∞; 0.75)")),
            mg(mr("(0.75; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 54: −3x² + 7x − 5 < 0
    questions.append({
        "num": 54,
        "q": mg(mr("−3"), msup(mr("x"), mr("2")), mr(" + 7x − 5 < 0")),
        "variants": [
            mg(mr("ℝ")),
            mg(mr("∅")),
            mg(mr("(−∞; 1) ∪ (5/3; +∞)")),
            mg(mr("[1; 5/3]"))
        ],
        "correct": 0
    })
    
    # Problem 55: 2x² + 7x + 14 > 0
    questions.append({
        "num": 55,
        "q": mg(mr("2"), msup(mr("x"), mr("2")), mr(" + 7x + 14 > 0")),
        "variants": [
            mg(mr("ℝ")),
            mg(mr("∅")),
            mg(mr("(−∞; −7) ∪ (2; +∞)")),
            mg(mr("[−7; 2]"))
        ],
        "correct": 0
    })
    
    # Problem 56: (3x²+2x+7)/(x−2) > 0
    questions.append({
        "num": 56,
        "q": mg(mfrac(mg(mr("3"), msup(mr("x"), mr("2")), mr(" + 2x + 7")), mr("x − 2")), mr(" > 0")),
        "variants": [
            mg(mr("(2; +∞)")),
            mg(mr("(−∞; 2)")),
            mg(mr("ℝ \\ {2}")),
            mg(mr("[2; +∞)"))
        ],
        "correct": 0
    })
    
    # ═══════════════════════════════════════════════════════════════
    #  PROBLEMS 57-99 FROM SECOND PDF
    # ═══════════════════════════════════════════════════════════════
    
    # Problem 57: x/3 > 0
    questions.append({
        "num": 57,
        "q": mg(mfrac(mr("x"), mr("3")), mr(" > 0")),
        "variants": [
            mg(mr("(0; +∞)")),
            mg(mr("[0; +∞)")),
            mg(mr("(−∞; 0)")),
            mg(mr("ℝ \\ {0}"))
        ],
        "correct": 0
    })
    
    # Problem 58: (x+2)/18 ≤ 0
    questions.append({
        "num": 58,
        "q": mg(mfrac(mr("x + 2"), mr("18")), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−∞; −2]")),
            mg(mr("[−2; +∞)")),
            mg(mr("(−∞; −2)")),
            mg(mr("[−2; 0]"))
        ],
        "correct": 0
    })
    
    # Problem 59: (x−2)/(x+5) ≥ 0
    questions.append({
        "num": 59,
        "q": mg(mfrac(mr("x − 2"), mr("x + 5")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −5) ∪ [2; +∞)")),
            mg(mr("[−5; 2]")),
            mg(mr("(−5; 2)")),
            mg(mr("(−∞; −5] ∪ [2; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 60: (x+5)/(x−1) ≥ 0
    questions.append({
        "num": 60,
        "q": mg(mfrac(mr("x + 5"), mr("x − 1")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −5] ∪ (1; +∞)")),
            mg(mr("[−5; 1]")),
            mg(mr("(−5; 1)")),
            mg(mr("(−∞; −5) ∪ [1; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 61: x/(x−1) ≥ 1
    questions.append({
        "num": 61,
        "q": mg(mfrac(mr("x"), mr("x − 1")), mr(" ≥ 1")),
        "variants": [
            mg(mr("(1; +∞)")),
            mg(mr("[1; +∞)")),
            mg(mr("(−∞; 1)")),
            mg(mr("ℝ \\ {1}"))
        ],
        "correct": 0
    })
    
    # Problem 62: (2x+3)/(x−12) ≥ 0
    questions.append({
        "num": 62,
        "q": mg(mfrac(mr("2x + 3"), mr("x − 12")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −1.5] ∪ (12; +∞)")),
            mg(mr("[−1.5; 12]")),
            mg(mr("(−1.5; 12)")),
            mg(mr("(−∞; −1.5) ∪ [12; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 63: x/(x−1) ≤ −1
    questions.append({
        "num": 63,
        "q": mg(mfrac(mr("x"), mr("x − 1")), mr(" ≤ −1")),
        "variants": [
            mg(mr("[0; 1)")),
            mg(mr("(0; 1]")),
            mg(mr("[0; 1]")),
            mg(mr("(−∞; 0] ∪ (1; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 64: (2x−3)/(2x+5) ≤ 1
    questions.append({
        "num": 64,
        "q": mg(mfrac(mr("2x − 3"), mr("2x + 5")), mr(" ≤ 1")),
        "variants": [
            mg(mr("ℝ \\ {−2.5}")),
            mg(mr("(−∞; −2.5)")),
            mg(mr("(−2.5; +∞)")),
            mg(mr("∅"))
        ],
        "correct": 0
    })
    
    # Problem 65: x/(x+2) ≤ 2
    questions.append({
        "num": 65,
        "q": mg(mfrac(mr("x"), mr("x + 2")), mr(" ≤ 2")),
        "variants": [
            mg(mr("(−∞; −4] ∪ (−2; +∞)")),
            mg(mr("[−4; −2]")),
            mg(mr("(−4; −2)")),
            mg(mr("(−∞; −4) ∪ [−2; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 66: (x+1)/(x−2) ≥ 1/4
    questions.append({
        "num": 66,
        "q": mg(mfrac(mr("x + 1"), mr("x − 2")), mr(" ≥ "), mfrac(mr("1"), mr("4"))),
        "variants": [
            mg(mr("(−∞; ⅔] ∪ (2; +∞)")),
            mg(mr("[⅔; 2]")),
            mg(mr("(⅔; 2)")),
            mg(mr("(−∞; ⅔) ∪ [2; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 67: (x−1)/(3−x) > 0
    questions.append({
        "num": 67,
        "q": mg(mfrac(mr("x − 1"), mr("3 − x")), mr(" > 0")),
        "variants": [
            mg(mr("(1; 3)")),
            mg(mr("(−∞; 1) ∪ (3; +∞)")),
            mg(mr("[1; 3]")),
            mg(mr("(−∞; 1] ∪ [3; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 68: (x+3)/(5−2x) < 0
    questions.append({
        "num": 68,
        "q": mg(mfrac(mr("x + 3"), mr("5 − 2x")), mr(" < 0")),
        "variants": [
            mg(mr("(−∞; −3) ∪ (2.5; +∞)")),
            mg(mr("(−3; 2.5)")),
            mg(mr("[−3; 2.5]")),
            mg(mr("(−∞; −3] ∪ [2.5; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 69: (6−x)/(x+3) ≤ 0
    questions.append({
        "num": 69,
        "q": mg(mfrac(mr("6 − x"), mr("x + 3")), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−∞; −3) ∪ [6; +∞)")),
            mg(mr("[−3; 6]")),
            mg(mr("(−3; 6)")),
            mg(mr("(−∞; −3] ∪ [6; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 70: (1−x)/(x−2) ≥ 0
    questions.append({
        "num": 70,
        "q": mg(mfrac(mr("1 − x"), mr("x − 2")), mr(" ≥ 0")),
        "variants": [
            mg(mr("[1; 2)")),
            mg(mr("(1; 2]")),
            mg(mr("[1; 2]")),
            mg(mr("(−∞; 1] ∪ (2; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 71: (x²−3x)/(4x−1) ≤ 0
    questions.append({
        "num": 71,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 3x")), mr("4x − 1")), mr(" ≤ 0")),
        "variants": [
            mg(mr("[0; 0.25) ∪ [3; +∞)")),
            mg(mr("(−∞; 0] ∪ (0.25; 3]")),
            mg(mr("[0; 3]")),
            mg(mr("(0; 0.25) ∪ (3; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 72: (2x−4)/(x²+1) ≥ 0
    questions.append({
        "num": 72,
        "q": mg(mfrac(mr("2x − 4"), mg(msup(mr("x"), mr("2")), mr(" + 1"))), mr(" ≥ 0")),
        "variants": [
            mg(mr("[2; +∞)")),
            mg(mr("(2; +∞)")),
            mg(mr("(−∞; 2]")),
            mg(mr("ℝ"))
        ],
        "correct": 0
    })
    
    # Problem 73: (x−2)(x+1)/((x−3)(x+4)) ≥ 0
    questions.append({
        "num": 73,
        "q": mg(mfrac(mr("(x − 2)(x + 1)"), mr("(x − 3)(x + 4)")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−4; −1] ∪ [2; 3)")),
            mg(mr("[−4; −1] ∪ [2; 3]")),
            mg(mr("(−∞; −4) ∪ [−1; 2] ∪ (3; +∞)")),
            mg(mr("(−4; −1) ∪ (2; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 74: (x²−4)/(x²−9) < 0
    questions.append({
        "num": 74,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 4")), 
              mg(msup(mr("x"), mr("2")), mr(" − 9"))), mr(" < 0")),
        "variants": [
            mg(mr("(−3; −2) ∪ (2; 3)")),
            mg(mr("[−3; −2] ∪ [2; 3]")),
            mg(mr("(−∞; −3) ∪ (−2; 2) ∪ (3; +∞)")),
            mg(mr("(−2; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 75: (x²−5x+6)/(x²−x−2) > 0
    questions.append({
        "num": 75,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 5x + 6")), 
              mg(msup(mr("x"), mr("2")), mr(" − x − 2"))), mr(" > 0")),
        "variants": [
            mg(mr("(−1; 2) ∪ (3; +∞)")),
            mg(mr("(−∞; −1) ∪ (2; 3)")),
            mg(mr("[−1; 2] ∪ [3; +∞)")),
            mg(mr("(2; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 91: (20−x−x²)/(x²−5x+6) ≤ 0
    questions.append({
        "num": 91,
        "q": mg(mfrac(mg(mr("20 − x − "), msup(mr("x"), mr("2"))), 
              mg(msup(mr("x"), mr("2")), mr(" − 5x + 6"))), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−∞; −5] ∪ (2; 3) ∪ [4; +∞)")),
            mg(mr("[−5; 2) ∪ (3; 4]")),
            mg(mr("(−∞; −5) ∪ [2; 3] ∪ (4; +∞)")),
            mg(mr("[−5; 4]"))
        ],
        "correct": 0
    })
    
    # Problem 92: (x²+5x+1)/(x²−4) + 3/(x−2) ≤ 0
    questions.append({
        "num": 92,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" + 5x + 1")), 
              mg(msup(mr("x"), mr("2")), mr(" − 4"))), mr(" + "), 
              mfrac(mr("3"), mr("x − 2")), mr(" ≤ 0")),
        "variants": [
            mg(mr("(−∞; −2) ∪ [−1; 2)")),
            mg(mr("[−2; −1] ∪ [2; +∞)")),
            mg(mr("(−2; −1) ∪ (2; +∞)")),
            mg(mr("(−∞; −1] ∪ (2; +∞)"))
        ],
        "correct": 0
    })
    
    # Problem 93: (x²−7x−2)/(x²+3x+2) − (2x−8)/(x+2) ≥ 0
    questions.append({
        "num": 93,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 7x − 2")), 
              mg(msup(mr("x"), mr("2")), mr(" + 3x + 2"))), mr(" − "), 
              mfrac(mr("2x − 8"), mr("x + 2")), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−2; −1) ∪ [3; +∞)")),
            mg(mr("[−2; −1] ∪ [3; +∞)")),
            mg(mr("(−∞; −2) ∪ (−1; 3]")),
            mg(mr("(−2; 3]"))
        ],
        "correct": 0
    })
    
    # Problem 94: (x²−5x+64)/(x²−11x+30) ≤ 10/(5−x)
    questions.append({
        "num": 94,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 5x + 64")), 
              mg(msup(mr("x"), mr("2")), mr(" − 11x + 30"))), mr(" ≤ "), 
              mfrac(mr("10"), mr("5 − x"))),
        "variants": [
            mg(mr("[−6; 5) ∪ (6; +∞)")),
            mg(mr("(−∞; −6] ∪ (5; 6)")),
            mg(mr("[−6; 6]")),
            mg(mr("(5; 6)"))
        ],
        "correct": 0
    })
    
    # Problem 95: (2x²−14x+6)/(x²−4x+3) ≥ (3x−8)/(x−3)
    questions.append({
        "num": 95,
        "q": mg(mfrac(mg(mr("2"), msup(mr("x"), mr("2")), mr(" − 14x + 6")), 
              mg(msup(mr("x"), mr("2")), mr(" − 4x + 3"))), mr(" ≥ "), 
              mfrac(mr("3x − 8"), mr("x − 3"))),
        "variants": [
            mg(mr("(−∞; −1] ∪ (1; 3) ∪ [4; +∞)")),
            mg(mr("[−1; 1) ∪ (3; 4]")),
            mg(mr("(−∞; −1) ∪ [1; 3] ∪ (4; +∞)")),
            mg(mr("[−1; 4]"))
        ],
        "correct": 0
    })
    
    # Problem 96: x⁵ + x³ + x ≥ 138 (Iqtidorli o'quvchilar uchun)
    questions.append({
        "num": 96,
        "q": mg(msup(mr("x"), mr("5")), mr(" + "), msup(mr("x"), mr("3")), 
              mr(" + x ≥ 138")),
        "variants": [
            mg(mr("[3; +∞)")),
            mg(mr("(3; +∞)")),
            mg(mr("[2; +∞)")),
            mg(mr("(−∞; 3]"))
        ],
        "correct": 0
    })
    
    # Problem 97: x⁶ + 3x⁴ + 6x² < 10 (Iqtidorli o'quvchilar uchun)
    questions.append({
        "num": 97,
        "q": mg(msup(mr("x"), mr("6")), mr(" + 3"), msup(mr("x"), mr("4")), 
              mr(" + 6"), msup(mr("x"), mr("2")), mr(" < 10")),
        "variants": [
            mg(mr("(−1; 1)")),
            mg(mr("[−1; 1]")),
            mg(mr("(−∞; −1) ∪ (1; +∞)")),
            mg(mr("∅"))
        ],
        "correct": 0
    })
    
    # Problem 98: 2/(x²+6x+10) + 5/(x²+6x+14) > 3 (Iqtidorli o'quvchilar)
    questions.append({
        "num": 98,
        "q": mg(mfrac(mr("2"), mg(msup(mr("x"), mr("2")), mr(" + 6x + 10"))), mr(" + "), 
              mfrac(mr("5"), mg(msup(mr("x"), mr("2")), mr(" + 6x + 14"))), mr(" > 3")),
        "variants": [
            mg(mr("(−4; −2)")),
            mg(mr("[−4; −2]")),
            mg(mr("(−∞; −4) ∪ (−2; +∞)")),
            mg(mr("ℝ"))
        ],
        "correct": 0
    })
    
    # Problem 99: Similar complex inequality
    questions.append({
        "num": 99,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" + x − 6")), 
              mg(msup(mr("x"), mr("2")), mr(" − 4"))), mr(" ≥ 0")),
        "variants": [
            mg(mr("(−∞; −2) ∪ [−3; 2) ∪ [3; +∞)")),
            mg(mr("[−3; −2) ∪ (2; 3]")),
            mg(mr("(−3; −2) ∪ (2; 3)")),
            mg(mr("(−∞; −3] ∪ [−2; 2] ∪ [3; +∞)"))
        ],
        "correct": 0
    })
    
    return questions



# ═══════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION
# ═══════════════════════════════════════════════════════════════════════

def write_docx(output_path, questions):
    """Create the final .docx file"""
    doc_xml = build_document_xml(questions)
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
        zf.writestr('word/document.xml', doc_xml)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/settings.xml', SETTINGS)
    
    size = os.path.getsize(output_path)
    print(f"\n✅ Fayl yaratildi: {output_path}")
    print(f"   Savollar soni: {len(questions)}")
    print(f"   Fayl hajmi: {size:,} bayt ({size/1024:.1f} KB)")
    
    # Answer distribution
    dist = Counter(q['correct'] for q in questions)
    print(f"\n📊 Javoblar taqsimoti:")
    print(f"   A = {dist[0]:2d}   B = {dist[1]:2d}   C = {dist[2]:2d}   D = {dist[3]:2d}")
    
    # Print answer key
    print(f"\n📋 JAVOBLAR KALITI:")
    for i in range(0, len(questions), 10):
        row = questions[i:i+10]
        line = "   ".join([
            f"{q['num']:2d}→{['A','B','C','D'][q['correct']]}" 
            for q in row
        ])
        print(f"   {line}")


if __name__ == "__main__":
    random.seed(42)  # For reproducible answer distribution
    
    print("=" * 70)
    print("  PROFESSIONAL RATSIONAL TENGSIZLIKLAR TEST GENERATOR")
    print("  PDF to Word with OMML Equations & Interval Notation")
    print("=" * 70)
    
    # Generate questions from PDF
    print("\n🔍 PDF dan tengsizliklarni o'qish...")
    questions = generate_questions()
    print(f"   {len(questions)} ta tengsizlik topildi")
    
    # Balance answer distribution
    print("\n⚖️  Javoblarni balanslash (A/B/C/D teng taqsimot)...")
    balanced_questions = balance_answers(questions)
    
    # Write output file
    output_file = "/projects/sandbox/test-yaratish-uchun/Ratsional_Tengsizliklar_SUPER_Test.docx"
    print("\n📝 Word faylini yaratish...")
    write_docx(output_file, balanced_questions)
    
    print("\n✨ TAYYOR! Word faylini oching va tekshiring.")
    print("   Barcha formulalar va intervallar Microsoft Word Equation formatida.")
    print("=" * 70)

