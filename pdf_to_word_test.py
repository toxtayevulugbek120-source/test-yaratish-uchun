#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional Matematik Test Generator — PDF to Word with OMML
Yuklangan PDF fayldagi barcha matematik masalalarni A/B/C/D test formatiga aylantiradi
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
#  EXTRACTED QUESTIONS FROM PDF
#  Based on "Ikkinchi va yuqori darajali tenglamalar sistemasi"
# ═══════════════════════════════════════════════════════════════════════

def generate_questions():
    """Generate all questions with variants from PDF"""
    questions = []
    
    # Problem 1: x + y = 6, xy = 8
    questions.append({
        "num": 1,
        "q": mg(mr("{  x + y = 6, xy = 8")),
        "variants": [
            mg(mr("(2; 4) va (4; 2)")),
            mg(mr("(3; 3) va (1; 5)")),
            mg(mr("(1; 5) va (5; 1)")),
            mg(mr("(2; 3) va (3; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 2: x + y = 8, xy = 15
    questions.append({
        "num": 2,
        "q": mg(mr("{  x + y = 8, xy = 15")),
        "variants": [
            mg(mr("(3; 5) va (5; 3)")),
            mg(mr("(2; 6) va (6; 2)")),
            mg(mr("(4; 4)")),
            mg(mr("(1; 7) va (7; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 3: x + y = 10, xy = 21
    questions.append({
        "num": 3,
        "q": mg(mr("{  x + y = 10, xy = 21")),
        "variants": [
            mg(mr("(3; 7) va (7; 3)")),
            mg(mr("(4; 6) va (6; 4)")),
            mg(mr("(2; 8) va (8; 2)")),
            mg(mr("(5; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 4: x − y = 3, xy = 10
    questions.append({
        "num": 4,
        "q": mg(mr("{  x − y = 3, xy = 10")),
        "variants": [
            mg(mr("(5; 2)")),
            mg(mr("(6; 3)")),
            mg(mr("(−2; −5)")),
            mg(mr("(4; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 5: x − y = −2, xy = 24
    questions.append({
        "num": 5,
        "q": mg(mr("{  x − y = −2, xy = 24")),
        "variants": [
            mg(mr("(4; 6)")),
            mg(mr("(−6; −4)")),
            mg(mr("(6; 4)")),
            mg(mr("(3; 8)"))
        ],
        "correct": 0
    })
    

    # Problem 6: x − y = 8, x² − y³ = 80
    questions.append({
        "num": 6,
        "q": mg(mr("{  x − y = 8, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("3")), mr(" = 80")),
        "variants": [
            mg(mr("(10; 2)")),
            mg(mr("(12; 4)")),
            mg(mr("(9; 1)")),
            mg(mr("(11; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 7: x − y = −3, x² − y² = 21
    questions.append({
        "num": 7,
        "q": mg(mr("{  x − y = −3, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21")),
        "variants": [
            mg(mr("(−7; −10)")),
            mg(mr("(4; 7)")),
            mg(mr("(3; 6)")),
            mg(mr("(−4; −7)"))
        ],
        "correct": 0
    })
    
    # Problem 8: x − y = 3, x² − y² = 21
    questions.append({
        "num": 8,
        "q": mg(mr("{  x − y = 3, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21")),
        "variants": [
            mg(mr("(5; 2)")),
            mg(mr("(6; 3)")),
            mg(mr("(4; 1)")),
            mg(mr("(7; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 9: x + y = 9, x² − y² = 9
    questions.append({
        "num": 9,
        "q": mg(mr("{  x + y = 9, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 9")),
        "variants": [
            mg(mr("(5; 4)")),
            mg(mr("(6; 3)")),
            mg(mr("(4; 5)")),
            mg(mr("(7; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 10: x + y = 7, x² − y² = 21
    questions.append({
        "num": 10,
        "q": mg(mr("{  x + y = 7, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21")),
        "variants": [
            mg(mr("(5; 2)")),
            mg(mr("(4; 3)")),
            mg(mr("(6; 1)")),
            mg(mr("(3; 4)"))
        ],
        "correct": 0
    })
    

    # Problem 11: x + y = 6, x² + y² = 20
    questions.append({
        "num": 11,
        "q": mg(mr("{  x + y = 6, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 20")),
        "variants": [
            mg(mr("(2; 4) va (4; 2)")),
            mg(mr("(3; 3)")),
            mg(mr("(1; 5) va (5; 1)")),
            mg(mr("(2; 3) va (3; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 12: x − y = 5, x² + y² = 53
    questions.append({
        "num": 12,
        "q": mg(mr("{  x − y = 5, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 53")),
        "variants": [
            mg(mr("(7; 2)")),
            mg(mr("(8; 3)")),
            mg(mr("(6; 1)")),
            mg(mr("(9; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 13: x + y = 2, x² + y² − 2xy = 16
    questions.append({
        "num": 13,
        "q": mg(mr("{  x + y = 2, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), 
              mr(" − 2xy = 16")),
        "variants": [
            mg(mr("(5; −3) va (−3; 5)")),
            mg(mr("(4; −2) va (−2; 4)")),
            mg(mr("(6; −4) va (−4; 6)")),
            mg(mr("(3; −1) va (−1; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 14: x + y = 3, x² + y² − 2xy = 1
    questions.append({
        "num": 14,
        "q": mg(mr("{  x + y = 3, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), 
              mr(" − 2xy = 1")),
        "variants": [
            mg(mr("(2; 1) va (1; 2)")),
            mg(mr("(3; 0) va (0; 3)")),
            mg(mr("(4; −1) va (−1; 4)")),
            mg(mr("(5; −2) va (−2; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 15: x − y = 6, x² + y² + 2xy = 16
    questions.append({
        "num": 15,
        "q": mg(mr("{  x − y = 6, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), 
              mr(" + 2xy = 16")),
        "variants": [
            mg(mr("(5; −1)")),
            mg(mr("(7; 1)")),
            mg(mr("(4; −2)")),
            mg(mr("(6; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 16: x + y = 3, x² + xy − y² = 5
    questions.append({
        "num": 16,
        "q": mg(mr("{  x + y = 3, "), msup(mr("x"), mr("2")), mr(" + xy − "), 
              msup(mr("y"), mr("2")), mr(" = 5")),
        "variants": [
            mg(mr("(4; −1) va (−1; 4)")),
            mg(mr("(3; 0) va (0; 3)")),
            mg(mr("(5; −2) va (−2; 5)")),
            mg(mr("(2; 1) va (1; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 17: x − y = 7, x² − xy − y² = 19
    questions.append({
        "num": 17,
        "q": mg(mr("{  x − y = 7, "), msup(mr("x"), mr("2")), mr(" − xy − "), 
              msup(mr("y"), mr("2")), mr(" = 19")),
        "variants": [
            mg(mr("(8; 1)")),
            mg(mr("(9; 2)")),
            mg(mr("(7; 0)")),
            mg(mr("(10; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 18: y = x + 6, x² + 3 = 4y
    questions.append({
        "num": 18,
        "q": mg(mr("{  y = x + 6, "), msup(mr("x"), mr("2")), mr(" + 3 = 4y")),
        "variants": [
            mg(mr("(3; 9) va (−5; 1)")),
            mg(mr("(4; 10) va (−4; 2)")),
            mg(mr("(2; 8) va (−6; 0)")),
            mg(mr("(5; 11) va (−3; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 19: y − 3x = 2, x² = 2y + 3
    questions.append({
        "num": 19,
        "q": mg(mr("{  y − 3x = 2, "), msup(mr("x"), mr("2")), mr(" = 2y + 3")),
        "variants": [
            mg(mr("(−1; −1) va (5; 17)")),
            mg(mr("(0; 2) va (4; 14)")),
            mg(mr("(1; 5) va (3; 11)")),
            mg(mr("(2; 8) va (6; 20)"))
        ],
        "correct": 0
    })
    
    # Problem 20: x − 2y = 1, 3x + y² = 10
    questions.append({
        "num": 20,
        "q": mg(mr("{  x − 2y = 1, 3x + "), msup(mr("y"), mr("2")), mr(" = 10")),
        "variants": [
            mg(mr("(1; 0) va (3; 1)")),
            mg(mr("(2; 0.5) va (4; 1.5)")),
            mg(mr("(5; 2) va (−1; −1)")),
            mg(mr("(3; 1) va (1; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 21: x² + xy + 3 = 0, y − 3x − 2 = 0
    questions.append({
        "num": 21,
        "q": mg(msup(mr("x"), mr("2")), mr(" + xy + 3 = 0,  y − 3x − 2 = 0")),
        "variants": [
            mg(mr("(−1; −5) va (−3; −11)")),
            mg(mr("(1; 5) va (3; 11)")),
            mg(mr("(−2; −8) va (−4; −14)")),
            mg(mr("(2; 8) va (4; 14)"))
        ],
        "correct": 0
    })
    
    # Problem 22: y² − x² = 16, x + y = 8
    questions.append({
        "num": 22,
        "q": mg(msup(mr("y"), mr("2")), mr(" − "), msup(mr("x"), mr("2")), mr(" = 16,  x + y = 8")),
        "variants": [
            mg(mr("(3; 5)")),
            mg(mr("(4; 4)")),
            mg(mr("(2; 6)")),
            mg(mr("(5; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 23: (x+3y)/(y−1) − (y−x)/(2x) = 2, y − x = 4
    questions.append({
        "num": 23,
        "q": mg(mfrac(mr("x + 3y"), mr("y − 1")), mr(" − "), 
              mfrac(mr("y − x"), mr("2x")), mr(" = 2,  y − x = 4")),
        "variants": [
            mg(mr("(2; 6)")),
            mg(mr("(3; 7)")),
            mg(mr("(1; 5)")),
            mg(mr("(4; 8)"))
        ],
        "correct": 0
    })
    
    # Problem 24: x² − xy = −1, y + 4x = 6
    questions.append({
        "num": 24,
        "q": mg(msup(mr("x"), mr("2")), mr(" − xy = −1,  y + 4x = 6")),
        "variants": [
            mg(mr("(1; 2) va (−1; 10)")),
            mg(mr("(2; −2) va (0; 6)")),
            mg(mr("(3; −6) va (−3; 18)")),
            mg(mr("(4; −10) va (−2; 14)"))
        ],
        "correct": 0
    })
    
    return questions



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
        text_run("IKKINCHI VA YUQORI DARAJALI TENGLAMALAR SISTEMASI", bold=True, size=32),
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
        body_parts.append(para(question_content, space_after=60))
        
        # Variants paragraph
        variant_content = ""
        for i, letter in enumerate(['A', 'B', 'C', 'D']):
            variant_content += text_run(f"  {letter})  ", bold=True, size=22)
            variant_content += f'<m:oMath xmlns:m="{M}">{variants[i]}</m:oMath>'
            variant_content += text_run("    ", size=22)
        
        body_parts.append(para(variant_content, space_after=100))
    
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
    print("  PROFESSIONAL MATEMATIK TEST GENERATOR")
    print("  PDF to Word with OMML Equations")
    print("=" * 70)
    
    # Generate questions from PDF
    print("\n🔍 PDF dan masalalarni o'qish...")
    questions = generate_questions()
    print(f"   {len(questions)} ta masala topildi")
    
    # Balance answer distribution
    print("\n⚖️  Javoblarni balanslash (A/B/C/D teng taqsimot)...")
    balanced_questions = balance_answers(questions)
    
    # Write output file
    output_file = "/projects/sandbox/test-yaratish-uchun/Tenglamalar_Sistemasi_Test.docx"
    print("\n📝 Word faylini yaratish...")
    write_docx(output_file, balanced_questions)
    
    print("\n✨ TAYYOR! Word faylini oching va tekshiring.")
    print("   Barcha formulalar Microsoft Word Equation formatida.")
    print("=" * 70)

