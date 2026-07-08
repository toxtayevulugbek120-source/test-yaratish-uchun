#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PROFESSIONAL TEST GENERATOR - IKKINCHI VA YUQORI DARAJALI TENGLAMALAR SISTEMASI
Yangi PDF asosida qayta yaratilgan versiya
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

def system_brace(eq1, eq2, eq3=None):
    """System with curly brace"""
    d = f'<m:d xmlns:m="{M}"><m:dPr><m:begChr m:val="{{"/><m:endChr m:val=""/><m:sepChr m:val=""/><m:grow m:val="1"/></m:dPr>'
    eqs = f'<m:e>{eq1}</m:e><m:e>{eq2}</m:e>'
    if eq3:
        eqs += f'<m:e>{eq3}</m:e>'
    return d + eqs + '</m:d>'


def generate_questions():
    """Generate all 25 questions from PDF"""
    qs = []
    
    # 1: x+y=6, xy=8
    qs.append({
        "num": 1,
        "q": system_brace(mg(mr("x + y = 6")), mg(mr("xy = 8"))),
        "v": [
            mg(mr("(2; 4); (4; 2)")),
            mg(mr("(1; 5); (5; 1)")),
            mg(mr("(3; 3)")),
            mg(mr("(−2; −4); (−4; −2)"))
        ],
        "c": 0
    })
    
    # 2: x+y=8, xy=15
    qs.append({
        "num": 2,
        "q": system_brace(mg(mr("x + y = 8")), mg(mr("xy = 15"))),
        "v": [
            mg(mr("(3; 5); (5; 3)")),
            mg(mr("(2; 6); (6; 2)")),
            mg(mr("(4; 4)")),
            mg(mr("(1; 7); (7; 1)"))
        ],
        "c": 0
    })
    
    # 3: x+y=10, xy=21
    qs.append({
        "num": 3,
        "q": system_brace(mg(mr("x + y = 10")), mg(mr("xy = 21"))),
        "v": [
            mg(mr("(3; 7); (7; 3)")),
            mg(mr("(2; 8); (8; 2)")),
            mg(mr("(4; 6); (6; 4)")),
            mg(mr("(5; 5)"))
        ],
        "c": 0
    })
    
    # 4: x-y=3, xy=10
    qs.append({
        "num": 4,
        "q": system_brace(mg(mr("x − y = 3")), mg(mr("xy = 10"))),
        "v": [
            mg(mr("(5; 2); (−2; −5)")),
            mg(mr("(4; 1); (−1; −4)")),
            mg(mr("(6; 3); (−3; −6)")),
            mg(mr("(3; 0); (0; −3)"))
        ],
        "c": 0
    })
    
    # 5: x-y=-2, xy=24
    qs.append({
        "num": 5,
        "q": system_brace(mg(mr("x − y = −2")), mg(mr("xy = 24"))),
        "v": [
            mg(mr("(4; 6); (−6; −4)")),
            mg(mr("(3; 5); (−5; −3)")),
            mg(mr("(2; 4); (−4; −2)")),
            mg(mr("(6; 8); (−8; −6)"))
        ],
        "c": 0
    })
    
    # 6: x-y=8, x²-y³=80
    qs.append({
        "num": 6,
        "q": system_brace(
            mg(mr("x − y = 8")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("3")), mr(" = 80"))
        ),
        "v": [
            mg(mr("(10; 2); (−8; −16)")),
            mg(mr("(9; 1); (−7; −15)")),
            mg(mr("(11; 3); (−9; −17)")),
            mg(mr("(12; 4); (−6; −14)"))
        ],
        "c": 0
    })
    
    # 7: x-y=-3, x²-y²=21
    qs.append({
        "num": 7,
        "q": system_brace(
            mg(mr("x − y = −3")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21"))
        ),
        "v": [
            mg(mr("(−1.5; 1.5)")),
            mg(mr("(−2; 1)")),
            mg(mr("(−3; 0)")),
            mg(mr("(0; 3)"))
        ],
        "c": 0
    })
    
    # 8: x-y=3, x²-y²=21
    qs.append({
        "num": 8,
        "q": system_brace(
            mg(mr("x − y = 3")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21"))
        ),
        "v": [
            mg(mr("(5; 2)")),
            mg(mr("(4; 1)")),
            mg(mr("(6; 3)")),
            mg(mr("(7; 4)"))
        ],
        "c": 0
    })
    
    # 9: x+y=9, x²-y²=9
    qs.append({
        "num": 9,
        "q": system_brace(
            mg(mr("x + y = 9")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 9"))
        ),
        "v": [
            mg(mr("(5; 4)")),
            mg(mr("(6; 3)")),
            mg(mr("(7; 2)")),
            mg(mr("(4.5; 4.5)"))
        ],
        "c": 0
    })
    
    # 10: x+y=7, x²-y²=21
    qs.append({
        "num": 10,
        "q": system_brace(
            mg(mr("x + y = 7")),
            mg(msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21"))
        ),
        "v": [
            mg(mr("(5; 2)")),
            mg(mr("(4; 3)")),
            mg(mr("(6; 1)")),
            mg(mr("(3.5; 3.5)"))
        ],
        "c": 0
    })
    
    # 11: x+y=6, x²+y²=20
    qs.append({
        "num": 11,
        "q": system_brace(
            mg(mr("x + y = 6")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 20"))
        ),
        "v": [
            mg(mr("(2; 4); (4; 2)")),
            mg(mr("(1; 5); (5; 1)")),
            mg(mr("(3; 3)")),
            mg(mr("(0; 6); (6; 0)"))
        ],
        "c": 0
    })
    
    # 12: x-y=5, x²+y²=53
    qs.append({
        "num": 12,
        "q": system_brace(
            mg(mr("x − y = 5")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 53"))
        ),
        "v": [
            mg(mr("(7; 2); (−2; −7)")),
            mg(mr("(6; 1); (−1; −6)")),
            mg(mr("(8; 3); (−3; −8)")),
            mg(mr("(5; 0); (0; −5)"))
        ],
        "c": 0
    })
    
    # 13: x+y=2, x²+y²-2xy=16
    qs.append({
        "num": 13,
        "q": system_brace(
            mg(mr("x + y = 2")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" − 2xy = 16"))
        ),
        "v": [
            mg(mr("(5; −3); (−3; 5)")),
            mg(mr("(4; −2); (−2; 4)")),
            mg(mr("(6; −4); (−4; 6)")),
            mg(mr("(1; 1)"))
        ],
        "c": 0
    })
    
    # 14: x+y=3, x²+y²-2xy=1
    qs.append({
        "num": 14,
        "q": system_brace(
            mg(mr("x + y = 3")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" − 2xy = 1"))
        ),
        "v": [
            mg(mr("(2; 1); (1; 2)")),
            mg(mr("(3; 0); (0; 3)")),
            mg(mr("(1.5; 1.5)")),
            mg(mr("(4; −1); (−1; 4)"))
        ],
        "c": 0
    })
    
    # 15: x-y=6, x²+y²+2xy=16
    qs.append({
        "num": 15,
        "q": system_brace(
            mg(mr("x − y = 6")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" + 2xy = 16"))
        ),
        "v": [
            mg(mr("(5; −1); (−1; −7)")),
            mg(mr("(4; −2); (−2; −8)")),
            mg(mr("(6; 0); (0; −6)")),
            mg(mr("(3; −3); (−3; −9)"))
        ],
        "c": 0
    })
    
    # 16: x-y=6, x²+y²+2xy=16
    qs.append({
        "num": 16,
        "q": system_brace(
            mg(mr("x − y = 6")),
            mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" + 2xy = 16"))
        ),
        "v": [
            mg(mr("(5; −1); (−1; −7)")),
            mg(mr("(4; −2); (−2; −8)")),
            mg(mr("(6; 0); (0; −6)")),
            mg(mr("(7; 1); (1; −5)"))
        ],
        "c": 0
    })
    
    # 17: x-y=7, x²-xy-y²=19
    qs.append({
        "num": 17,
        "q": system_brace(
            mg(mr("x − y = 7")),
            mg(msup(mr("x"), mr("2")), mr(" − xy − "), msup(mr("y"), mr("2")), mr(" = 19"))
        ),
        "v": [
            mg(mr("(6; −1); (−2; −9)")),
            mg(mr("(5; −2); (−3; −10)")),
            mg(mr("(7; 0); (0; −7)")),
            mg(mr("(8; 1); (1; −6)"))
        ],
        "c": 0
    })
    
    # 18: y=x+6, x²+3=4y
    qs.append({
        "num": 18,
        "q": system_brace(
            mg(mr("y = x + 6")),
            mg(msup(mr("x"), mr("2")), mr(" + 3 = 4y"))
        ),
        "v": [
            mg(mr("(3; 9); (−7; −1)")),
            mg(mr("(2; 8); (−8; −2)")),
            mg(mr("(4; 10); (−6; 0)")),
            mg(mr("(5; 11); (−5; 1)"))
        ],
        "c": 0
    })
    
    # 19: y-3x=2, x²=2y+3
    qs.append({
        "num": 19,
        "q": system_brace(
            mg(mr("y − 3x = 2")),
            mg(msup(mr("x"), mr("2")), mr(" = 2y + 3"))
        ),
        "v": [
            mg(mr("(−1; −1); (7; 23)")),
            mg(mr("(0; 2); (6; 20)")),
            mg(mr("(1; 5); (5; 17)")),
            mg(mr("(2; 8); (4; 14)"))
        ],
        "c": 0
    })
    
    # 20: x-2y=1, 3x+y²=10
    qs.append({
        "num": 20,
        "q": system_brace(
            mg(mr("x − 2y = 1")),
            mg(mr("3x + "), msup(mr("y"), mr("2")), mr(" = 10"))
        ),
        "v": [
            mg(mr("(1; 0); (3; 1)")),
            mg(mr("(2; 0.5); (4; 1.5)")),
            mg(mr("(0; −0.5); (5; 2)")),
            mg(mr("(3; 1); (5; 2)"))
        ],
        "c": 0
    })
    
    # 21: x²+xy+3=0, y-3x-2=0
    qs.append({
        "num": 21,
        "q": system_brace(
            mg(msup(mr("x"), mr("2")), mr(" + xy + 3 = 0")),
            mg(mr("y − 3x − 2 = 0"))
        ),
        "v": [
            mg(mr("(−1; −1); (−3; −7)")),
            mg(mr("(0; 2); (−2; −4)")),
            mg(mr("(1; 5); (−4; −10)")),
            mg(mr("(−2; −4); (2; 8)"))
        ],
        "c": 0
    })
    
    # 22: y²-x²=16, x+y=8
    qs.append({
        "num": 22,
        "q": system_brace(
            mg(msup(mr("y"), mr("2")), mr(" − "), msup(mr("x"), mr("2")), mr(" = 16")),
            mg(mr("x + y = 8"))
        ),
        "v": [
            mg(mr("(3; 5)")),
            mg(mr("(2; 6)")),
            mg(mr("(4; 4)")),
            mg(mr("(1; 7)"))
        ],
        "c": 0
    })
    
    # 23: 3 equations system
    qs.append({
        "num": 23,
        "q": system_brace(
            mg(mfrac(mr("x + 3y"), mr("y − 1")), mr(" − "), mfrac(mr("y − x"), mr("2x")), mr(" = 2")),
            mg(mr("y − 1 = 2x")),
            mg(mr("y − x = 4"))
        ),
        "v": [
            mg(mr("(3; 7)")),
            mg(mr("(2; 6)")),
            mg(mr("(4; 8)")),
            mg(mr("(1; 5)"))
        ],
        "c": 0
    })
    
    # 24: x²-xy=-1, y+4x=6
    qs.append({
        "num": 24,
        "q": system_brace(
            mg(msup(mr("x"), mr("2")), mr(" − xy = −1")),
            mg(mr("y + 4x = 6"))
        ),
        "v": [
            mg(mr("(1; 2); (−1; 10)")),
            mg(mr("(0; 6); (2; −2)")),
            mg(mr("(2; −2); (−2; 14)")),
            mg(mr("(0.5; 4); (−0.5; 8)"))
        ],
        "c": 0
    })
    
    # 25: 2x²+xy-14=0, 3x-y-3=0
    qs.append({
        "num": 25,
        "q": system_brace(
            mg(mr("2"), msup(mr("x"), mr("2")), mr(" + xy − 14 = 0")),
            mg(mr("3x − y − 3 = 0"))
        ),
        "v": [
            mg(mr("(2; 3); (−7/3; −10)")),
            mg(mr("(1; 0); (−3; −12)")),
            mg(mr("(3; 6); (−2; −9)")),
            mg(mr("(−1; −6); (4; 9)"))
        ],
        "c": 0
    })
    
    return qs


def balance_answers(questions):
    """Balance A/B/C/D distribution"""
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
    content = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    content.append(f'<w:document xmlns:w="{W}" xmlns:m="{M}"><w:body>')
    
    # Title
    content.append('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>')
    content.append('<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>')
    content.append('<w:t>IKKINCHI VA YUQORI DARAJALI TENGLAMALAR SISTEMASI TEST</w:t></w:r></w:p>')
    content.append('<w:p/>')
    
    # Questions
    for q in questions:
        content.append('<w:p><w:r><w:rPr><w:b/></w:rPr>')
        content.append(f'<w:t xml:space="preserve">{q["num"]}. Sistemani yeching: </w:t></w:r></w:p>')
        content.append(f'<w:p><m:oMath>{q["q"]}</m:oMath></w:p>')
        
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
    print(f"✅ Yaratildi: {output}")
    print(f"   Savollar: {len(questions)}")
    print(f"   Hajmi: {size:,} bayt ({size/1024:.1f} KB)")
    
    counts = [0, 0, 0, 0]
    for q in questions:
        counts[q['c']] += 1
    print(f"\n📊 Javoblar: A={counts[0]} B={counts[1]} C={counts[2]} D={counts[3]}\n")
    
    print("📋 KALITI:")
    for i in range(0, len(questions), 10):
        line = "  "
        for j in range(i, min(i+10, len(questions))):
            q = questions[j]
            line += f" {q['num']:2d}→{chr(65+q['c'])}"
        print(line)


def main():
    print("="*70)
    print("  IKKINCHI VA YUQORI DARAJALI TENGLAMALAR SISTEMASI")
    print("  Professional Test Generator - OMML Format")
    print("="*70)
    print()
    
    questions = generate_questions()
    balance_answers(questions)
    create_docx(questions, "/projects/sandbox/test-yaratish-uchun/Ikkinchi_Yuqori_Darajali_Sistemalar_Test_V2.docx")
    
    print("\n" + "="*70)


if __name__ == "__main__":
    main()
