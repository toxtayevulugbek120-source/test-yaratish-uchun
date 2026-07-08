#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PROFESSIONAL TEST GENERATOR - IKKINCHI VA YUQORI DARAJALI TENGLAMALAR SISTEMASI
================================================================================
Bu generator PDF'dagi tenglamalar sistemasini A, B, C, D variantli testga aylantiradi.
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

def mg(*parts):
    """Group multiple OMML parts"""
    return "".join(parts)


def system_with_brace(eq1, eq2, eq3=None):
    """Create system of equations with opening curly brace"""
    # Delimiter with opening brace
    delim_start = (f'<m:d xmlns:m="{M}"><m:dPr>'
                   f'<m:begChr m:val="{{"/><m:endChr m:val=""/>'
                   f'<m:sepChr m:val=""/><m:grow m:val="1"/></m:dPr>')
    
    equations = f'<m:e>{eq1}</m:e><m:e>{eq2}</m:e>'
    if eq3:
        equations += f'<m:e>{eq3}</m:e>'
    
    return delim_start + equations + '</m:d>'


def generate_questions():
    """Generate all questions from PDF"""
    questions = []
    
    print("🔍 PDF dan tenglamalar sistemasini o'qish...")
    
    # Problem 1: x+y=6, xy=8
    questions.append({
        "num": 1,
        "q": system_with_brace(
            mg(mr("x + y = 6")),
            mg(mr("xy = 8"))
        ),
        "variants": [
            mg(mr("(2; 4); (4; 2)")),
            mg(mr("(1; 5); (5; 1)")),
            mg(mr("(3; 3)")),
            mg(mr("(0; 6); (6; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 2: x+y=8, xy=15
    questions.append({
        "num": 2,
        "q": system_with_brace(
            mg(mr("x + y = 8")),
            mg(mr("xy = 15"))
        ),
        "variants": [
            mg(mr("(3; 5); (5; 3)")),
            mg(mr("(2; 6); (6; 2)")),
            mg(mr("(1; 7); (7; 1)")),
            mg(mr("(4; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 3: x+y=10, xy=21
    questions.append({
        "num": 3,
        "q": system_with_brace(
            mg(mr("x + y = 10")),
            mg(mr("xy = 21"))
        ),
        "variants": [
            mg(mr("(3; 7); (7; 3)")),
            mg(mr("(2; 8); (8; 2)")),
            mg(mr("(4; 6); (6; 4)")),
            mg(mr("(5; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 4: x−y=3, xy=10
    questions.append({
        "num": 4,
        "q": system_with_brace(
            mg(mr("x − y = 3")),
            mg(mr("xy = 10"))
        ),
        "variants": [
            mg(mr("(5; 2); (−2; −5)")),
            mg(mr("(4; 1); (−1; −4)")),
            mg(mr("(6; 3); (−3; −6)")),
            mg(mr("(3; 0); (0; −3)"))
        ],
        "correct": 0
    })
    
    # Problem 5: x−y=−2, xy=24
    questions.append({
        "num": 5,
        "q": system_with_brace(
            mg(mr("x − y = −2")),
            mg(mr("xy = 24"))
        ),
        "variants": [
            mg(mr("(4; 6); (−6; −4)")),
            mg(mr("(3; 5); (−5; −3)")),
            mg(mr("(2; 4); (−4; −2)")),
            mg(mr("(6; 8); (−8; −6)"))
        ],
        "correct": 0
    })
    
    # Problem 6: x−y=8, x²−y³=80
    questions.append({
        "num": 6,
        "q": system_with_brace(
            mg(mr("x − y = 8")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("3")), mr(" = 80"))
        ),
        "variants": [
            mg(mr("(10; 2); (−8; −16)")),
            mg(mr("(9; 1); (−7; −15)")),
            mg(mr("(12; 4); (−6; −14)")),
            mg(mr("(11; 3); (−9; −17)"))
        ],
        "correct": 0
    })
    
    # Problem 7: x−y=−3, x²−y²=21
    questions.append({
        "num": 7,
        "q": system_with_brace(
            mg(mr("x − y = −3")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21"))
        ),
        "variants": [
            mg(mr("(−1.5; 1.5)")),
            mg(mr("(−2; 1)")),
            mg(mr("(−3; 0)")),
            mg(mr("(0; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 8: x−y=3, x²−y²=21
    questions.append({
        "num": 8,
        "q": system_with_brace(
            mg(mr("x − y = 3")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21"))
        ),
        "variants": [
            mg(mr("(5; 2)")),
            mg(mr("(4; 1)")),
            mg(mr("(6; 3)")),
            mg(mr("(3; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 9: x+y=9, x²−y²=9
    questions.append({
        "num": 9,
        "q": system_with_brace(
            mg(mr("x + y = 9")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 9"))
        ),
        "variants": [
            mg(mr("(5; 4)")),
            mg(mr("(6; 3)")),
            mg(mr("(7; 2)")),
            mg(mr("(4.5; 4.5)"))
        ],
        "correct": 0
    })
    
    # Problem 10: x+y=7, x²−y²=21
    questions.append({
        "num": 10,
        "q": system_with_brace(
            mg(mr("x + y = 7")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21"))
        ),
        "variants": [
            mg(mr("(5; 2)")),
            mg(mr("(4; 3)")),
            mg(mr("(6; 1)")),
            mg(mr("(3.5; 3.5)"))
        ],
        "correct": 0
    })
    
    # Problem 11: x+y=6, x²+y²=20
    questions.append({
        "num": 11,
        "q": system_with_brace(
            mg(mr("x + y = 6")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 20"))
        ),
        "variants": [
            mg(mr("(2; 4); (4; 2)")),
            mg(mr("(1; 5); (5; 1)")),
            mg(mr("(3; 3)")),
            mg(mr("(0; 6); (6; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 12: x−y=5, x²+y²=53
    questions.append({
        "num": 12,
        "q": system_with_brace(
            mg(mr("x − y = 5")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 53"))
        ),
        "variants": [
            mg(mr("(7; 2); (−2; −7)")),
            mg(mr("(6; 1); (−1; −6)")),
            mg(mr("(8; 3); (−3; −8)")),
            mg(mr("(5; 0); (0; −5)"))
        ],
        "correct": 0
    })
    
    # Problem 13: x+y=2, x²+y²−2xy=16
    questions.append({
        "num": 13,
        "q": system_with_brace(
            mg(mr("x + y = 2")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" − 2xy = 16"))
        ),
        "variants": [
            mg(mr("(5; −3); (−3; 5)")),
            mg(mr("(4; −2); (−2; 4)")),
            mg(mr("(6; −4); (−4; 6)")),
            mg(mr("(1; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 14: x+y=3, x²+y²−2xy=1
    questions.append({
        "num": 14,
        "q": system_with_brace(
            mg(mr("x + y = 3")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" − 2xy = 1"))
        ),
        "variants": [
            mg(mr("(2; 1); (1; 2)")),
            mg(mr("(3; 0); (0; 3)")),
            mg(mr("(1.5; 1.5)")),
            mg(mr("(4; −1); (−1; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 15: x−y=6, x²+y²+2xy=16
    questions.append({
        "num": 15,
        "q": system_with_brace(
            mg(mr("x − y = 6")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" + 2xy = 16"))
        ),
        "variants": [
            mg(mr("(5; −1); (−1; −7)")),
            mg(mr("(4; −2); (−2; −8)")),
            mg(mr("(6; 0); (0; −6)")),
            mg(mr("(3; −3); (−3; −9)"))
        ],
        "correct": 0
    })
    
    # Problem 16: x+y=3, x²+xy−y²=5
    questions.append({
        "num": 16,
        "q": system_with_brace(
            mg(mr("x + y = 3")),
            mg(msup(mr("x"), mr("2")), mr(" + xy − "), msup(mr("y"), mr("2")), mr(" = 5"))
        ),
        "variants": [
            mg(mr("(2.5; 0.5)")),
            mg(mr("(2; 1)")),
            mg(mr("(3; 0)")),
            mg(mr("(1.5; 1.5)"))
        ],
        "correct": 0
    })
    
    # Problem 17: x−y=7, x²−xy−y²=19
    questions.append({
        "num": 17,
        "q": system_with_brace(
            mg(mr("x − y = 7")),
            mg(msup(mr("x"), mr("2")), mr(" − xy − "), msup(mr("y"), mr("2")), mr(" = 19"))
        ),
        "variants": [
            mg(mr("(6; −1); (−2; −9)")),
            mg(mr("(5; −2); (−3; −10)")),
            mg(mr("(7; 0); (0; −7)")),
            mg(mr("(4; −3); (−4; −11)"))
        ],
        "correct": 0
    })
    
    # Problem 18: y=x+6, x²+3=4y
    questions.append({
        "num": 18,
        "q": system_with_brace(
            mg(mr("y = x + 6")),
            mg(msup(mr("x"), mr("2")), mr(" + 3 = 4y"))
        ),
        "variants": [
            mg(mr("(3; 9); (−7; −1)")),
            mg(mr("(2; 8); (−8; −2)")),
            mg(mr("(4; 10); (−6; 0)")),
            mg(mr("(1; 7); (−9; −3)"))
        ],
        "correct": 0
    })
    
    # Problem 19: y−3x=2, x²=2y+3
    questions.append({
        "num": 19,
        "q": system_with_brace(
            mg(mr("y − 3x = 2")),
            mg(msup(mr("x"), mr("2")), mr(" = 2y + 3"))
        ),
        "variants": [
            mg(mr("(−1; −1); (7; 23)")),
            mg(mr("(0; 2); (6; 20)")),
            mg(mr("(1; 5); (5; 17)")),
            mg(mr("(−2; −4); (8; 26)"))
        ],
        "correct": 0
    })
    
    # Problem 20: x−2y=1, 3x+y²=10
    questions.append({
        "num": 20,
        "q": system_with_brace(
            mg(mr("x − 2y = 1")),
            mg(mr("3x + "), msup(mr("y"), mr("2")), mr(" = 10"))
        ),
        "variants": [
            mg(mr("(1; 0); (3; 1)")),
            mg(mr("(2; 0.5); (4; 1.5)")),
            mg(mr("(0; −0.5); (5; 2)")),
            mg(mr("(−1; −1); (6; 2.5)"))
        ],
        "correct": 0
    })
    
    # Problem 21: x²+xy+3=0, y−3x−2=0
    questions.append({
        "num": 21,
        "q": system_with_brace(
            mg(msup(mr("x"), mr("2")), mr(" + xy + 3 = 0")),
            mg(mr("y − 3x − 2 = 0"))
        ),
        "variants": [
            mg(mr("(−1; −1); (−3; −7)")),
            mg(mr("(0; 2); (−2; −4)")),
            mg(mr("(1; 5); (−4; −10)")),
            mg(mr("(−2; −4); (2; 8)"))
        ],
        "correct": 0
    })
    
    # Problem 22: y²−x²=16, x+y=8
    questions.append({
        "num": 22,
        "q": system_with_brace(
            mg(msup(mr("y"), mr("2")), mr(" − "), msup(mr("x"), mr("2")), mr(" = 16")),
            mg(mr("x + y = 8"))
        ),
        "variants": [
            mg(mr("(3; 5)")),
            mg(mr("(2; 6)")),
            mg(mr("(4; 4)")),
            mg(mr("(1; 7)"))
        ],
        "correct": 0
    })
    
    # Problem 23: 3 tenglamali sistema
    questions.append({
        "num": 23,
        "q": system_with_brace(
            mg(mfrac(mr("x + 3y"), mr("y − 1")), mr(" − "), mfrac(mr("y − x"), mr("2x")), mr(" = 2")),
            mg(mr("y − 1 = 2x")),
            mg(mr("y − x = 4"))
        ),
        "variants": [
            mg(mr("(3; 7)")),
            mg(mr("(2; 6)")),
            mg(mr("(4; 8)")),
            mg(mr("(1; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 24: x²−xy=−1, y+4x=6
    questions.append({
        "num": 24,
        "q": system_with_brace(
            mg(msup(mr("x"), mr("2")), mr(" − xy = −1")),
            mg(mr("y + 4x = 6"))
        ),
        "variants": [
            mg(mr("(1; 2); (−1; 10)")),
            mg(mr("(0; 6); (2; −2)")),
            mg(mr("(2; −2); (−2; 14)")),
            mg(mr("(0.5; 4); (−0.5; 8)"))
        ],
        "correct": 0
    })
    
    # Problem 25: 2x²+xy−14=0, 3x−y−3=0
    questions.append({
        "num": 25,
        "q": system_with_brace(
            mg(mr("2"), msup(mr("x"), mr("2")), mr(" + xy − 14 = 0")),
            mg(mr("3x − y − 3 = 0"))
        ),
        "variants": [
            mg(mr("(2; 3); (−7/3; −10)")),
            mg(mr("(1; 0); (−3; −12)")),
            mg(mr("(3; 6); (−2; −9)")),
            mg(mr("(0; −3); (4; 9)"))
        ],
        "correct": 0
    })
    
    print(f"   {len(questions)} ta masala topildi\n")
    return questions


def balance_answers(questions):
    """Balance answer distribution across A/B/C/D"""
    print("⚖️  Javoblarni balanslash (A/B/C/D teng taqsimot)...\n")
    
    random.seed(42)
    for q in questions:
        random.shuffle(q['variants'])
    
    target_per_option = len(questions) // 4
    counts = [0, 0, 0, 0]
    
    for q in questions:
        counts[q['correct']] += 1
    
    for q in questions:
        current = q['correct']
        if counts[current] > target_per_option:
            for i in range(4):
                if i != current and counts[i] < target_per_option:
                    q['variants'][current], q['variants'][i] = q['variants'][i], q['variants'][current]
                    counts[current] -= 1
                    counts[i] += 1
                    q['correct'] = i
                    break


def create_word_docx(questions, output_file):
    """Create professional Word (.docx) test"""
    print("📝 Word faylini yaratish...\n")
    
    # Build document.xml content
    content = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    content.append(f'<w:document xmlns:w="{W}" xmlns:m="{M}">')
    content.append('<w:body>')
    
    # Title
    content.append('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>')
    content.append('<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>')
    content.append('<w:t>IKKINCHI VA YUQORI DARAJALI TENGLAMALAR SISTEMASI TEST</w:t>')
    content.append('</w:r></w:p>')
    content.append('<w:p/>')  # Blank line
    
    # Questions
    for q in questions:
        # Question number and text
        content.append('<w:p>')
        content.append('<w:r><w:rPr><w:b/></w:rPr>')
        content.append(f'<w:t xml:space="preserve">{q["num"]}. Sistemani yeching: </w:t>')
        content.append('</w:r></w:p>')
        
        # System equation
        content.append('<w:p>')
        content.append(f'<m:oMath>{q["q"]}</m:oMath>')
        content.append('</w:p>')
        
        # Answer variants
        for i, var in enumerate(q['variants']):
            content.append('<w:p>')
            content.append(f'<w:r><w:t xml:space="preserve">{chr(65+i)}) </w:t></w:r>')
            content.append(f'<m:oMath>{var}</m:oMath>')
            content.append('</w:p>')
        
        content.append('<w:p/>')  # Blank line
    
    # Answer key
    content.append('<w:p><w:pPr><w:pageBreakBefore/><w:jc w:val="center"/></w:pPr>')
    content.append('<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>')
    content.append('<w:t>Answer Key</w:t></w:r></w:p>')
    content.append('<w:p/>')
    
    for q in questions:
        content.append('<w:p>')
        content.append(f'<w:r><w:t>{q["num"]} — {chr(65 + q["correct"])}</w:t></w:r>')
        content.append('</w:p>')
    
    content.append('</w:body></w:document>')
    
    # Create .docx structure
    docx_content = "\n".join(content)
    
    # Create minimal .docx file
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as docx:
        # [Content_Types].xml
        docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
        
        # _rels/.rels
        docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
        
        # word/_rels/document.xml.rels
        docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')
        
        # word/document.xml
        docx.writestr('word/document.xml', docx_content)
    
    # Stats
    size = os.path.getsize(output_file)
    print(f"✅ Fayl yaratildi: {output_file}")
    print(f"   Savollar soni: {len(questions)}")
    print(f"   Fayl hajmi: {size:,} bayt ({size/1024:.1f} KB)\n")
    
    # Answer distribution
    counts = [0, 0, 0, 0]
    for q in questions:
        counts[q['correct']] += 1
    
    print("📊 Javoblar taqsimoti:")
    print(f"   A = {counts[0]}   B = {counts[1]}   C = {counts[2]}   D = {counts[3]}\n")
    
    # Answer key preview
    print("📋 JAVOBLAR KALITI:")
    for i in range(0, len(questions), 10):
        line = "  "
        for j in range(i, min(i+10, len(questions))):
            q = questions[j]
            line += f" {q['num']:2d}→{chr(65+q['correct'])}"
        print(line)


def main():
    """Main function"""
    print("="*70)
    print("  PROFESSIONAL TENGLAMALAR SISTEMASI TEST GENERATOR")
    print("  PDF to Word with OMML Equations & Solution Pairs")
    print("="*70)
    print()
    
    # Generate questions
    questions = generate_questions()
    
    # Balance answers
    balance_answers(questions)
    
    # Create Word file
    output_file = "/projects/sandbox/test-yaratish-uchun/Ikkinchi_Yuqori_Darajali_Sistemalar_Test.docx"
    create_word_docx(questions, output_file)
    
    print()
    print("="*70)


if __name__ == "__main__":
    main()
