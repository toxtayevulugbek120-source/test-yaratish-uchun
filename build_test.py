#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional A/B/C/D Math Test Generator
Chiziqli Tenglamalar — 90 savol
Extracts original OMML from source docx, shuffles correct answers,
builds a brand-new Word (.docx) with proper OMML equations.
"""

import zipfile, re, os, random

# ─── Fixed seed for reproducible answer distribution ───
random.seed(42)

# ════════════════════════════════════════════════════════
#  STEP 1 — Extract raw OMML blocks from source docx
# ════════════════════════════════════════════════════════

SRC = 'Chiziqli_Tenglamalar_Test.docx'

with zipfile.ZipFile(SRC) as z:
    src_doc = z.read('word/document.xml').decode('utf-8')

raw_omaths = re.findall(r'<m:oMath[^>]*>(.*?)</m:oMath>', src_doc, re.DOTALL)
assert len(raw_omaths) == 450, f"Expected 450 oMath blocks, got {len(raw_omaths)}"

# Each question: blocks[i*5+0]=question, [i*5+1..4]=variants (pos 0 = correct answer)
def get_block(qi, vi):
    """Return cleaned inner OMML XML for question qi (0-based), variant vi (0=question,1-4=variants)."""
    return raw_omaths[qi * 5 + vi].strip()


# ════════════════════════════════════════════════════════
#  STEP 2 — Build shuffled question list
#  correct_pos: 0=A,1=B,2=C,3=D  (spread evenly)
# ════════════════════════════════════════════════════════

# Pre-assign correct positions to ensure even distribution: 90/4 = 22-23 each
# A=23, B=23, C=22, D=22
positions = [0]*23 + [1]*23 + [2]*22 + [3]*22
random.shuffle(positions)

QUESTIONS = []  # list of (q_omml, [A_omml, B_omml, C_omml, D_omml], correct_idx)

for qi in range(90):
    q_omml    = get_block(qi, 0)
    correct   = get_block(qi, 1)   # original slot 1 = correct answer
    distract1 = get_block(qi, 2)
    distract2 = get_block(qi, 3)
    distract3 = get_block(qi, 4)

    wrong = [distract1, distract2, distract3]
    cp = positions[qi]  # where correct answer goes (0=A,1=B,2=C,3=D)

    variants = wrong[:]
    variants.insert(cp, correct)   # insert correct at position cp
    # now variants[0]=A, [1]=B, [2]=C, [3]=D

    QUESTIONS.append((q_omml, variants, cp))

# Verify distribution
from collections import Counter
dist = Counter(cp for _, _, cp in QUESTIONS)
print("Answer distribution:", {['A','B','C','D'][k]: v for k,v in sorted(dist.items())})


# ════════════════════════════════════════════════════════
#  STEP 3 — Word (.docx) XML builder helpers
# ════════════════════════════════════════════════════════

WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
MNS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
RNSP = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def esc(t):
    return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def w_run(text, bold=False, sz=24, italic=False):
    b  = '<w:b/><w:bCs/>' if bold else ''
    it = '<w:i/><w:iCs/>' if italic else ''
    return (
        f'<w:r><w:rPr>{b}{it}'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        f'</w:rPr>'
        f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
    )

def omath_para(inner_omml, prefix_runs='', center=False, space_after=80, indent=0):
    jc  = '<w:jc w:val="center"/>' if center else ''
    ind = f'<w:ind w:left="{indent}"/>' if indent else ''
    spc = f'<w:spacing w:after="{space_after}" w:line="276" w:lineRule="auto"/>'
    omath = (
        f'<m:oMath {MNS} {WNS}>'
        + inner_omml +
        '</m:oMath>'
    )
    return (
        f'<w:p {WNS} {MNS}>'
        f'<w:pPr>{jc}{spc}{ind}</w:pPr>'
        + prefix_runs
        + omath +
        '</w:p>'
    )

def text_para(runs, center=False, space_after=120, indent=0):
    jc  = '<w:jc w:val="center"/>' if center else ''
    ind = f'<w:ind w:left="{indent}"/>' if indent else ''
    spc = f'<w:spacing w:after="{space_after}" w:line="276" w:lineRule="auto"/>'
    return (
        f'<w:p {WNS}>'
        f'<w:pPr>{jc}{spc}{ind}</w:pPr>'
        + ''.join(runs) +
        '</w:p>'
    )

def page_break():
    return f'<w:p {WNS}><w:r><w:br w:type="page"/></w:r></w:p>'

def divider():
    return text_para([w_run('─' * 68, sz=18)], center=True, space_after=60)

VARIANT_LETTERS = ['A', 'B', 'C', 'D']

def question_block(num, q_omml, variants, correct_idx):
    """Build XML paragraphs for one question."""
    parts = []

    # ── Question line ──────────────────────────────────────
    prefix = w_run(f'{num}.  ', bold=True, sz=26)
    parts.append(omath_para(q_omml, prefix_runs=prefix, space_after=60))

    # ── Variants — one per line with letter label ──────────
    for i, (letter, var_omml) in enumerate(zip(VARIANT_LETTERS, variants)):
        is_correct = (i == correct_idx)
        # Bold the label only (not the omath) — helps visual layout
        label_run = w_run(f'   {letter})  ', bold=False, sz=24)
        parts.append(omath_para(var_omml, prefix_runs=label_run,
                                space_after=50, indent=360))

    # small gap after each question
    parts.append(text_para([], space_after=80))
    return ''.join(parts)


# ════════════════════════════════════════════════════════
#  STEP 4 — Document XML
# ════════════════════════════════════════════════════════

def build_document_xml():
    body = []

    # ── TITLE ──
    body.append(text_para(
        [w_run('25. CHIZIQLI TENGLAMALAR', bold=True, sz=32)],
        center=True, space_after=60
    ))
    body.append(text_para(
        [w_run('Professional A / B / C / D  Test  (90 savol)', bold=False, sz=24)],
        center=True, space_after=60
    ))
    body.append(text_para(
        [w_run("To'g'ri javobni tanlang.", bold=False, sz=24, italic=True)],
        center=True, space_after=100
    ))
    body.append(divider())

    # ── SECTIONS ──
    sections = [
        (1,  10,  "1–10. Oddiy chiziqli tenglamalar"),
        (11, 20,  "11–20. Kasr va murakkab tenglamalar"),
        (21, 30,  "21–30. Qavsli va ko'p qadamli tenglamalar"),
        (31, 40,  "31–40. Kasrli tenglamalar (sodda)"),
        (41, 50,  "41–50. Kasrli va nisbat tenglamalari"),
        (51, 60,  "51–60. Nisbat, ichma-ich kasr tenglamalari"),
        (61, 65,  "61–65. Identik va hech-yechimli tenglamalar"),
        (66, 70,  "66–70. Parametrli tenglamalar — a = ? (x berilgan)"),
        (71, 75,  "71–75. x ni y orqali ifodalash"),
        (76, 80,  "76–80. Yechimga ega bo'lmaydigan tenglamalar (a = ?)"),
        (81, 85,  "81–85. Cheksiz ko'p yechimli tenglamalar (a = ?)"),
        (86, 90,  "86–90. Bir ta yechimli tenglamalar (a = ?)"),
    ]

    for (s_start, s_end, s_title) in sections:
        body.append(text_para(
            [w_run(s_title + ':', bold=True, sz=26)],
            space_after=60
        ))
        for qnum in range(s_start, s_end + 1):
            qi = qnum - 1
            q_omml, variants, correct_idx = QUESTIONS[qi]
            body.append(question_block(qnum, q_omml, variants, correct_idx))
        body.append(page_break())

    # ── ANSWER KEY ──
    body.append(text_para(
        [w_run('JAVOBLAR KALITI  (ANSWER KEY)', bold=True, sz=30)],
        center=True, space_after=80
    ))
    body.append(divider())

    # Print in rows of 5
    for row_start in range(0, 90, 5):
        cells = []
        for qi in range(row_start, min(row_start + 5, 90)):
            letter = VARIANT_LETTERS[QUESTIONS[qi][2]]
            cells.append(f'{qi+1:2d} — {letter}')
        line = '        '.join(cells)
        body.append(text_para([w_run(line, sz=24)], center=True, space_after=70))

    # ── SECTION PROPERTIES ──
    body.append(
        f'<w:sectPr {WNS}>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1440"/>'
        '</w:sectPr>'
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {WNS} {MNS} {RNSP}>'
        '<w:body>' + ''.join(body) + '</w:body>'
        '</w:document>'
    )


# ════════════════════════════════════════════════════════
#  STEP 5 — Static docx skeleton files
# ════════════════════════════════════════════════════════

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
      <w:lang w:val="uz-Latn-UZ"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
</w:styles>'''


# ════════════════════════════════════════════════════════
#  STEP 6 — Write .docx
# ════════════════════════════════════════════════════════

def write_docx(out_path):
    doc_xml = build_document_xml()
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',          CONTENT_TYPES)
        z.writestr('_rels/.rels',                  ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', WORD_RELS)
        z.writestr('word/document.xml',            doc_xml)
        z.writestr('word/styles.xml',              STYLES)
        z.writestr('word/settings.xml',            SETTINGS)
    size = os.path.getsize(out_path)
    print(f'✓  Saved: {out_path}  ({size:,} bytes)')
    return doc_xml


if __name__ == '__main__':
    out = 'Chiziqli_Tenglamalar_PROFESSIONAL_TEST.docx'
    doc_xml = write_docx(out)

    # ── Quick validation ──
    import xml.etree.ElementTree as ET
    ET.fromstring(doc_xml)
    print('✓  XML valid')

    omath_n  = doc_xml.count('<m:oMath ')
    para_n   = doc_xml.count('<w:p ')
    print(f'✓  oMath elements : {omath_n}')
    print(f'✓  Paragraphs     : {para_n}')

    # Verify all 90 question numbers present
    missing = [i for i in range(1, 91) if f'>{i}.' not in doc_xml and f'>{i:2d}.' not in doc_xml]
    if missing:
        print(f'✗  Missing questions: {missing}')
    else:
        print('✓  All 90 questions present')

    print('\nFull answer key:')
    for row in range(0, 90, 10):
        line = '   '.join(
            f'{qi+1:2d}={["A","B","C","D"][QUESTIONS[qi][2]]}'
            for qi in range(row, min(row+10, 90))
        )
        print(' ', line)
