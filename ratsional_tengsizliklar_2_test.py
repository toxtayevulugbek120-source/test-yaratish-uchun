#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PROFESSIONAL TEST GENERATOR - RATSIONAL TENGSIZLIKLAR 2
========================================================
PDF'dagi barcha ratsional tengsizliklarni A/B/C/D test formatiga aylantiradi
Javoblar: interval notatsiyasi
"""

import zipfile
import os
import random

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def mr(text):
    """Math run"""
    safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return f'<m:r xmlns:m="{M}"><m:t xml:space="preserve">{safe}</m:t></m:r>'

def mfrac(num, den):
    """Stacked fraction"""
    return (f'<m:f xmlns:m="{M}"><m:fPr><m:type m:val="bar"/></m:fPr>'
            f'<m:num>{num}</m:num><m:den>{den}</m:den></m:f>')

def msup(base, sup):
    """Superscript"""
    return f'<m:sSup xmlns:m="{M}"><m:e>{base}</m:e><m:sup>{sup}</m:sup></m:sSup>'

def mg(*parts):
    """Group parts"""
    return "".join(parts)


def generate_questions():
    """Generate all questions from PDF"""
    qs = []
    
    print("🔍 PDF dan ratsional tengsizliklarni o'qish...")
    
    # 80: (x-3)/(x+1) ≥ 0
    qs.append({
        "num": 80,
        "q": mg(mfrac(mr("x − 3"), mr("x + 1")), mr(" ≥ 0")),
        "v": [
            mg(mr("(−∞; −1) ∪ [3; +∞)")),
            mg(mr("(−∞; −1] ∪ [3; +∞)")),
            mg(mr("[−1; 3]")),
            mg(mr("(−1; 3]"))
        ],
        "c": 0
    })
    
    # 81: (2+x)/(3-x) ≤ 0
    qs.append({
        "num": 81,
        "q": mg(mfrac(mr("2 + x"), mr("3 − x")), mr(" ≤ 0")),
        "v": [
            mg(mr("(−∞; −2] ∪ (3; +∞)")),
            mg(mr("(−∞; −2) ∪ (3; +∞)")),
            mg(mr("[−2; 3]")),
            mg(mr("(−2; 3)"))
        ],
        "c": 0
    })
    
    # 82: (x+7)/(2x-5) > 0
    qs.append({
        "num": 82,
        "q": mg(mfrac(mr("x + 7"), mr("2x − 5")), mr(" > 0")),
        "v": [
            mg(mr("(−∞; −7) ∪ (2.5; +∞)")),
            mg(mr("(−∞; −7] ∪ [2.5; +∞)")),
            mg(mr("(−7; 2.5)")),
            mg(mr("[−7; 2.5]"))
        ],
        "c": 0
    })
    
    # 83: (2x+7)/(x-1) ≥ 0
    qs.append({
        "num": 83,
        "q": mg(mfrac(mr("2x + 7"), mr("x − 1")), mr(" ≥ 0")),
        "v": [
            mg(mr("(−∞; −3.5] ∪ (1; +∞)")),
            mg(mr("(−∞; −3.5) ∪ (1; +∞)")),
            mg(mr("[−3.5; 1]")),
            mg(mr("(−3.5; 1)"))
        ],
        "c": 0
    })
    
    # 84: (x+1)/(1-x) ≤ 0
    qs.append({
        "num": 84,
        "q": mg(mfrac(mr("x + 1"), mr("1 − x")), mr(" ≤ 0")),
        "v": [
            mg(mr("(−∞; −1] ∪ (1; +∞)")),
            mg(mr("(−∞; −1) ∪ (1; +∞)")),
            mg(mr("[−1; 1]")),
            mg(mr("(−1; 1)"))
        ],
        "c": 0
    })
    
    # 85: 2x/(x+1) < 1
    qs.append({
        "num": 85,
        "q": mg(mfrac(mr("2x"), mr("x + 1")), mr(" < 1")),
        "v": [
            mg(mr("(−1; 1)")),
            mg(mr("(−∞; −1) ∪ (1; +∞)")),
            mg(mr("[−1; 1]")),
            mg(mr("(−∞; 1)"))
        ],
        "c": 0
    })
    
    # 86: 3/(x-5) ≤ 1
    qs.append({
        "num": 86,
        "q": mg(mfrac(mr("3"), mr("x − 5")), mr(" ≤ 1")),
        "v": [
            mg(mr("(−∞; 5) ∪ [8; +∞)")),
            mg(mr("(−∞; 5] ∪ [8; +∞)")),
            mg(mr("[5; 8]")),
            mg(mr("(5; 8)"))
        ],
        "c": 0
    })
    
    # 87: (x-1)/(x+2) ≥ 2
    qs.append({
        "num": 87,
        "q": mg(mfrac(mr("x − 1"), mr("x + 2")), mr(" ≥ 2")),
        "v": [
            mg(mr("(−2; −5]")),
            mg(mr("[−5; −2)")),
            mg(mr("(−∞; −5] ∪ (−2; +∞)")),
            mg(mr("(−5; −2)"))
        ],
        "c": 0
    })
    
    # 88: (x-7)/(x+1) < 3
    qs.append({
        "num": 88,
        "q": mg(mfrac(mr("x − 7"), mr("x + 1")), mr(" < 3")),
        "v": [
            mg(mr("(−5; −1)")),
            mg(mr("(−∞; −5) ∪ (−1; +∞)")),
            mg(mr("[−5; −1]")),
            mg(mr("(−1; +∞)"))
        ],
        "c": 0
    })
    
    # 89: (x+2)/(x-2) > 5
    qs.append({
        "num": 89,
        "q": mg(mfrac(mr("x + 2"), mr("x − 2")), mr(" > 5")),
        "v": [
            mg(mr("(2; 3)")),
            mg(mr("(−∞; 2) ∪ (3; +∞)")),
            mg(mr("[2; 3]")),
            mg(mr("(3; +∞)"))
        ],
        "c": 0
    })
    
    # 90: 5/(x+6) > 2x
    qs.append({
        "num": 90,
        "q": mg(mfrac(mr("5"), mr("x + 6")), mr(" > 2x")),
        "v": [
            mg(mr("(−6; −5) ∪ (−0.5; 0.5)")),
            mg(mr("(−∞; −6) ∪ (−5; 0.5)")),
            mg(mr("(−5; 0.5)")),
            mg(mr("(−6; 0.5)"))
        ],
        "c": 0
    })
    
    # 91: (20-x-x²)/(x²-5x+6) ≤ 0
    qs.append({
        "num": 91,
        "q": mg(mfrac(mg(mr("20 − x − "), msup(mr("x"), mr("2"))), mg(msup(mr("x"), mr("2")), mr(" − 5x + 6"))), mr(" ≤ 0")),
        "v": [
            mg(mr("(−∞; −5] ∪ (2; 3) ∪ [4; +∞)")),
            mg(mr("(−∞; −5) ∪ (2; 3) ∪ (4; +∞)")),
            mg(mr("[−5; 2) ∪ (3; 4]")),
            mg(mr("(−5; 2) ∪ (3; 4)"))
        ],
        "c": 0
    })
    
    # 92: (x²+5x+1)/(x²-4) + 3/(x-2) ≤ 0
    qs.append({
        "num": 92,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" + 5x + 1")), mg(msup(mr("x"), mr("2")), mr(" − 4"))), mr(" + "), mfrac(mr("3"), mr("x − 2")), mr(" ≤ 0")),
        "v": [
            mg(mr("(−∞; −2) ∪ [−1; 2)")),
            mg(mr("(−∞; −2] ∪ [−1; 2]")),
            mg(mr("(−2; −1] ∪ (2; +∞)")),
            mg(mr("[−2; −1] ∪ [2; +∞)"))
        ],
        "c": 0
    })
    
    # 93: (x²-7x-2)/(x²+3x+2) - (2x-8)/(x+2) ≥ 0
    qs.append({
        "num": 93,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 7x − 2")), mg(msup(mr("x"), mr("2")), mr(" + 3x + 2"))), mr(" − "), mfrac(mr("2x − 8"), mr("x + 2")), mr(" ≥ 0")),
        "v": [
            mg(mr("(−2; −1) ∪ [2; +∞)")),
            mg(mr("(−∞; −2) ∪ (−1; 2]")),
            mg(mr("[−2; −1] ∪ [2; +∞)")),
            mg(mr("(−2; −1] ∪ (2; +∞)"))
        ],
        "c": 0
    })
    
    # 94: (x²-5x+64)/(x²-11x+30) ≤ 10/(5-x)
    qs.append({
        "num": 94,
        "q": mg(mfrac(mg(msup(mr("x"), mr("2")), mr(" − 5x + 64")), mg(msup(mr("x"), mr("2")), mr(" − 11x + 30"))), mr(" ≤ "), mfrac(mr("10"), mr("5 − x"))),
        "v": [
            mg(mr("(−∞; 5) ∪ (6; +∞)")),
            mg(mr("(−∞; 5] ∪ [6; +∞)")),
            mg(mr("[5; 6]")),
            mg(mr("(5; 6)"))
        ],
        "c": 0
    })
    
    # 95: (2x²-14x+6)/(x²-4x+3) ≥ (3x-8)/(x-3)
    qs.append({
        "num": 95,
        "q": mg(mfrac(mg(mr("2"), msup(mr("x"), mr("2")), mr(" − 14x + 6")), mg(msup(mr("x"), mr("2")), mr(" − 4x + 3"))), mr(" ≥ "), mfrac(mr("3x − 8"), mr("x − 3"))),
        "v": [
            mg(mr("(−∞; 1) ∪ (3; 4]")),
            mg(mr("(−∞; 1] ∪ [3; 4]")),
            mg(mr("[1; 3) ∪ [4; +∞)")),
            mg(mr("(1; 3) ∪ (4; +∞)"))
        ],
        "c": 0
    })
    
    # 96-99: Iqtidorli o'quvchilar uchun
    
    # 96: x⁵+x³+x ≥ 138
    qs.append({
        "num": 96,
        "q": mg(msup(mr("x"), mr("5")), mr(" + "), msup(mr("x"), mr("3")), mr(" + x ≥ 138")),
        "v": [
            mg(mr("[3; +∞)")),
            mg(mr("(3; +∞)")),
            mg(mr("(−∞; 3]")),
            mg(mr("[−3; 3]"))
        ],
        "c": 0
    })
    
    # 97: x⁶+3x⁴+6x² < 10
    qs.append({
        "num": 97,
        "q": mg(msup(mr("x"), mr("6")), mr(" + 3"), msup(mr("x"), mr("4")), mr(" + 6"), msup(mr("x"), mr("2")), mr(" < 10")),
        "v": [
            mg(mr("(−1; 1)")),
            mg(mr("[−1; 1]")),
            mg(mr("(−∞; −1) ∪ (1; +∞)")),
            mg(mr("(−∞; 1)"))
        ],
        "c": 0
    })
    
    # 98: 2/(x²+6x+10) + 5/(x²+6x+14) > 3
    qs.append({
        "num": 98,
        "q": mg(mfrac(mr("2"), mg(msup(mr("x"), mr("2")), mr(" + 6x + 10"))), mr(" + "), mfrac(mr("5"), mg(msup(mr("x"), mr("2")), mr(" + 6x + 14"))), mr(" > 3")),
        "v": [
            mg(mr("∅")),
            mg(mr("ℝ")),
            mg(mr("(−∞; −3)")),
            mg(mr("(−3; +∞)"))
        ],
        "c": 0
    })
    
    print(f"   {len(qs)} ta masala topildi\n")
    return qs


def balance_answers(questions):
    """Balance A/B/C/D distribution"""
    print("⚖️  Javoblarni balanslash (A/B/C/D teng taqsimot)...\n")
    
    random.seed(42)
    for q in questions:
        random.shuffle(q['v'])
    
    target = len(questions) // 4
    counts = [0, 0, 0, 0]
    
    for q in questions:
        counts[q['c']] += 1
    
    for q in questions:
        curr = q['c']
        if counts[curr] > target:
            for i in range(4):
                if i != curr and counts[i] < target:
                    q['v'][curr], q['v'][i] = q['v'][i], q['v'][curr]
                    counts[curr] -= 1
                    counts[i] += 1
                    q['c'] = i
                    break


def create_docx(questions, output):
    """Create Word document"""
    print("📝 Word faylini yaratish...\n")
    
    content = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    content.append(f'<w:document xmlns:w="{W}" xmlns:m="{M}"><w:body>')
    
    # Title
    content.append('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>')
    content.append('<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>')
    content.append('<w:t>RATSIONAL TENGSIZLIKLAR TEST 2</w:t></w:r></w:p>')
    content.append('<w:p/>')
    
    # Questions
    for q in questions:
        content.append('<w:p><w:r><w:rPr><w:b/></w:rPr>')
        content.append(f'<w:t xml:space="preserve">{q["num"]}. </w:t></w:r>')
        content.append(f'<m:oMath>{q["q"]}</m:oMath></w:p>')
        
        for i, var in enumerate(q['v']):
            content.append(f'<w:p><w:r><w:t xml:space="preserve">{chr(65+i)}) </w:t></w:r>')
            content.append(f'<m:oMath>{var}</m:oMath></w:p>')
        
        content.append('<w:p/>')
    
    # Answer key
    content.append('<w:p><w:pPr><w:pageBreakBefore/><w:jc w:val="center"/></w:pPr>')
    content.append('<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>')
    content.append('<w:t>Answer Key</w:t></w:r></w:p><w:p/>')
    
    for q in questions:
        content.append(f'<w:p><w:r><w:t>{q["num"]} — {chr(65 + q["c"])}</w:t></w:r></w:p>')
    
    content.append('</w:body></w:document>')
    
    docx_xml = "\n".join(content)
    
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
        
        z.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
        
        z.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')
        
        z.writestr('word/document.xml', docx_xml)
    
    size = os.path.getsize(output)
    print(f"✅ Fayl yaratildi: {output}")
    print(f"   Savollar soni: {len(questions)}")
    print(f"   Fayl hajmi: {size:,} bayt ({size/1024:.1f} KB)\n")
    
    counts = [0, 0, 0, 0]
    for q in questions:
        counts[q['c']] += 1
    print(f"📊 Javoblar: A={counts[0]} B={counts[1]} C={counts[2]} D={counts[3]}\n")
    
    print("📋 JAVOBLAR KALITI:")
    for i in range(0, len(questions), 10):
        line = "  "
        for j in range(i, min(i+10, len(questions))):
            q = questions[j]
            line += f" {q['num']:2d}→{chr(65+q['c'])}"
        print(line)


def main():
    print("="*70)
    print("  RATSIONAL TENGSIZLIKLAR TEST 2")
    print("  Professional Test Generator - Interval Notation")
    print("="*70)
    print()
    
    questions = generate_questions()
    balance_answers(questions)
    create_docx(questions, "/projects/sandbox/test-yaratish-uchun/Ratsional_Tengsizliklar_2_Test.docx")
    
    print("\n" + "="*70)


if __name__ == "__main__":
    main()
