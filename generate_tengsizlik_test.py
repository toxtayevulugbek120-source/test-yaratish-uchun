#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional A/B/C/D Math Test Generator
Chiziqli Tengsizliklar va Tengsizliklar Sistemasi — 45 savol
Word (.docx) OMML formatida
"""
import zipfile, os

# ─── OMML helper functions ─────────────────────────────────────────────────

def mr(text):
    """Math run — XML-escaped"""
    safe = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
    return f'<m:r><m:t xml:space="preserve">{safe}</m:t></m:r>'

def frac(num_xml, den_xml):
    """Stacked fraction (OMML)"""
    return (f'<m:f><m:fPr><m:type m:val="bar"/></m:fPr>'
            f'<m:num>{num_xml}</m:num>'
            f'<m:den>{den_xml}</m:den></m:f>')

def sup(base_xml, exp_xml):
    """Superscript"""
    return (f'<m:sSup><m:sSupPr/>'
            f'<m:e>{base_xml}</m:e>'
            f'<m:sup>{exp_xml}</m:sup></m:sSup>')

def grp(*items):
    """Group multiple OMML items"""
    return ''.join(items)

# ─── Word run helper ───────────────────────────────────────────────────────

def wr(text, bold=False, sz=24):
    b = '<w:b/><w:bCs/>' if bold else ''
    safe = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return (f'<w:r><w:rPr>{b}'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            f'</w:rPr>'
            f'<w:t xml:space="preserve">{safe}</w:t></w:r>')

def omath(inner):
    return f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{inner}</m:oMath>'

def para(runs_and_omath, spacing_before=120, spacing_after=80, indent=0, center=False):
    jc = '<w:jc w:val="center"/>' if center else ''
    ind = f'<w:ind w:left="{indent}"/>' if indent else ''
    spc = f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}" w:line="276" w:lineRule="auto"/>'
    return (f'<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            f'<w:pPr>{jc}{spc}{ind}</w:pPr>'
            + ''.join(runs_and_omath)
            + '</w:p>')

def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


# ─── Question data structure ──────────────────────────────────────────────
# Each question: (number, question_omml, key_letter, A_omml, B_omml, C_omml, D_omml)
# key_letter: which letter is correct ('A','B','C','D')

QUESTIONS = []

# ═══════════════════════════════════════════════════════════════════
# 1-30: TENGSIZLIKLARNI YECHING
# ═══════════════════════════════════════════════════════════════════

# 1. x - 7 > 0  =>  x > 7
QUESTIONS.append((1,
    grp(mr('x − 7 > 0')),
    'C',
    mr('x > 8'),       # A: sign error +1
    mr('x > −7'),      # B: moved wrong side
    mr('x > 7'),       # C: CORRECT
    mr('x < 7'),       # D: direction flipped
))

# 2. x + 5 < 0  =>  x < -5
QUESTIONS.append((2,
    grp(mr('x + 5 < 0')),
    'A',
    mr('x < −5'),      # A: CORRECT
    mr('x < 5'),       # B: sign forgotten
    mr('x > −5'),      # C: direction flipped
    mr('x > 5'),       # D: both errors
))

# 3. 3 + x ≤ 0  =>  x ≤ -3
QUESTIONS.append((3,
    grp(mr('3 + x ≤ 0')),
    'D',
    mr('x ≥ −3'),      # A: direction flipped
    mr('x ≤ 3'),       # B: sign wrong
    mr('x ≥ 3'),       # C: two errors
    mr('x ≤ −3'),      # D: CORRECT
))

# 4. 2 - x ≥ 5  =>  x ≤ -3
QUESTIONS.append((4,
    grp(mr('2 − x ≥ 5')),
    'B',
    mr('x ≥ −3'),      # A: direction flipped
    mr('x ≤ −3'),      # B: CORRECT
    mr('x ≤ 3'),       # C: sign error
    mr('x ≥ 3'),       # D: two errors
))

# 5. 3 + 2x < 7  =>  x < 2
QUESTIONS.append((5,
    grp(mr('3 + 2x < 7')),
    'A',
    mr('x < 2'),       # A: CORRECT
    mr('x < 5'),       # B: forgot to divide by 2
    mr('x > 2'),       # C: direction flipped
    mr('x < −2'),      # D: subtracted wrong
))

# 6. 5x - 8 > 2  =>  x > 2
QUESTIONS.append((6,
    grp(mr('5x − 8 > 2')),
    'C',
    mr('x > 10'),      # A: forgot to divide
    mr('x < 2'),       # B: direction flipped
    mr('x > 2'),       # C: CORRECT
    mr('x > −2'),      # D: -8 treated as +8 then divided
))

# 7. 3x + 4 ≤ 2x + 5  =>  x ≤ 1
QUESTIONS.append((7,
    grp(mr('3x + 4 ≤ 2x + 5')),
    'D',
    mr('x ≤ 9'),       # A: added instead of subtracted
    mr('x ≥ 1'),       # B: direction flipped
    mr('x ≤ −1'),      # C: sign error on constant
    mr('x ≤ 1'),       # D: CORRECT
))

# 8. 2x - 2 > 3x - 4  =>  x < 2
QUESTIONS.append((8,
    grp(mr('2x − 2 > 3x − 4')),
    'B',
    mr('x > 2'),       # A: direction flipped
    mr('x < 2'),       # B: CORRECT
    mr('x < −2'),      # C: sign error
    mr('x > −2'),      # D: sign + direction errors
))

# 9. 3x + 4 > 5x - 6  =>  x < 5
QUESTIONS.append((9,
    grp(mr('3x + 4 > 5x − 6')),
    'A',
    mr('x < 5'),       # A: CORRECT
    mr('x > 5'),       # B: direction flipped
    mr('x < −5'),      # C: sign error
    mr('x < 1'),       # D: arithmetic error (10/2 vs 10/(-2))
))

# 10. 6x + 7 ≤ 3x + 1  =>  x ≤ -2
QUESTIONS.append((10,
    grp(mr('6x + 7 ≤ 3x + 1')),
    'C',
    mr('x ≤ 2'),       # A: sign error
    mr('x ≥ −2'),      # B: direction flipped
    mr('x ≤ −2'),      # C: CORRECT
    mr('x ≥ 2'),       # D: two errors
))


# 11. 4x + 3 ≥ 2x + 7  =>  x ≥ 2
QUESTIONS.append((11,
    grp(mr('4x + 3 ≥ 2x + 7')),
    'B',
    mr('x ≥ 5'),       # A: added instead of subtracted
    mr('x ≥ 2'),       # B: CORRECT
    mr('x ≤ 2'),       # C: direction flipped
    mr('x ≥ −2'),      # D: sign error
))

# 12. 5x + 1 > 2x - 8  =>  x > -3
QUESTIONS.append((12,
    grp(mr('5x + 1 > 2x − 8')),
    'D',
    mr('x > 3'),       # A: sign error on result
    mr('x < −3'),      # B: direction flipped
    mr('x > −9'),      # C: divided by wrong number
    mr('x > −3'),      # D: CORRECT
))

# 13. 3(x+2) < 4x - 5  =>  x > 11
QUESTIONS.append((13,
    grp(mr('3(x + 2) < 4x − 5')),
    'A',
    mr('x > 11'),      # A: CORRECT
    mr('x < 11'),      # B: direction flipped
    mr('x > 1'),       # C: bracket opened wrong: 3x+2 not 3x+6
    mr('x > −11'),     # D: sign error
))

# 14. 7(x+3) < 2x + 11  =>  x < -2
QUESTIONS.append((14,
    grp(mr('7(x + 3) < 2x + 11')),
    'C',
    mr('x < 2'),       # A: sign error
    mr('x > −2'),      # B: direction flipped
    mr('x < −2'),      # C: CORRECT
    mr('x < −4'),      # D: arithmetic error (10/5 vs 10/5)
))

# 15. 3(x+2) ≥ 4(x+2) + 5  =>  x ≤ -7
QUESTIONS.append((15,
    grp(mr('3(x + 2) ≥ 4(x + 2) + 5')),
    'D',
    mr('x ≤ 7'),       # A: sign error
    mr('x ≥ −7'),      # B: direction flipped
    mr('x ≤ −3'),      # C: forgot to expand correctly
    mr('x ≤ −7'),      # D: CORRECT
))

# 16. 2(x+3)+4(x-2) ≤ 3(4-x)+8x  =>  x ≤ 2
# LHS: 2x+6+4x-8=6x-2; RHS: 12-3x+8x=12+5x => 6x-2≤12+5x => x≤14
# Wait: 6x-2 ≤ 12+5x => x ≤ 14. Let me recheck from PDF.
# 2(x+3)+4(x-2)≤3(4-x)+8x => 6x-2 ≤ 12+5x => x ≤ 14
QUESTIONS.append((16,
    grp(mr('2(x + 3) + 4(x − 2) ≤ 3(4 − x) + 8x')),
    'B',
    mr('x ≤ 10'),      # A: arithmetic slip
    mr('x ≤ 14'),      # B: CORRECT
    mr('x ≥ 14'),      # C: direction flipped
    mr('x ≤ 2'),       # D: wrong constant
))

# 17. 4x+3(2x+2) < 5(3x-4)+2(x-1)  =>  x > 4
# LHS: 4x+6x+6=10x+6; RHS: 15x-20+2x-2=17x-22 => 10x+6<17x-22 => 28<7x => x>4
QUESTIONS.append((17,
    grp(mr('4x + 3(2x + 2) < 5(3x − 4) + 2(x − 1)')),
    'A',
    mr('x > 4'),       # A: CORRECT
    mr('x < 4'),       # B: direction flipped
    mr('x > 2'),       # C: arithmetic error
    mr('x > −4'),      # D: sign error
))

# 18. x(x+2)-3(x+3) > x²+3x-9  =>  x > 0
# x²+2x-3x-9 > x²+3x-9 => -x > 3x-0 wait: x²+2x-3x-9 = x²-x-9
# x²-x-9 > x²+3x-9 => -x > 3x => -4x>0 => x<0  ← recheck
# x(x+2)-3(x+3) = x²+2x-3x-9 = x²-x-9
# x²-x-9 > x²+3x-9 => -4x>0 => x<0
QUESTIONS.append((18,
    grp(mr('x(x + 2) − 3(x + 3) > x² + 3x − 9')),
    'C',
    mr('x > 0'),       # A: sign error on final step
    mr('x ≥ 0'),       # B: wrong sign and boundary
    mr('x < 0'),       # C: CORRECT
    mr('x ≤ 0'),       # D: wrong boundary type
))

# 19. x²-2x+9 < (x+1)(x+3)  =>  x > 3
# RHS: x²+4x+3 => -2x+9 < 4x+3 => 6 < 6x => x > 1  ← recheck
# -2x+9 < 4x+3 => 6 < 6x => x > 1
QUESTIONS.append((19,
    grp(mr('x² − 2x + 9 < (x + 1)(x + 3)')),
    'D',
    mr('x > 3'),       # A: wrong arithmetic
    mr('x < 1'),       # B: direction flipped
    mr('x > −1'),      # C: sign error
    mr('x > 1'),       # D: CORRECT
))

# 20. (x+4)(x-2) ≤ (x-5)(x+4)  =>  x ≤ -4
# LHS: x²+2x-8; RHS: x²-x-20 => 2x-8 ≤ -x-20 => 3x ≤ -12 => x ≤ -4
QUESTIONS.append((20,
    grp(mr('(x + 4)(x − 2) ≤ (x − 5)(x + 4)')),
    'A',
    mr('x ≤ −4'),      # A: CORRECT
    mr('x ≥ −4'),      # B: direction flipped
    mr('x ≤ 4'),       # C: sign error
    mr('x ≤ −12'),     # D: forgot to divide by 3
))


# 21. (2x+3)/5 > (4x+3)/9  =>  x < 6
# 9(2x+3) > 5(4x+3) => 18x+27 > 20x+15 => 12 > 2x => x < 6
QUESTIONS.append((21,
    grp(frac(mr('2x + 3'), mr('5')), mr(' > '), frac(mr('4x + 3'), mr('9'))),
    'C',
    mr('x > 6'),       # A: direction flipped
    mr('x < −6'),      # B: sign error
    mr('x < 6'),       # C: CORRECT
    mr('x < 3'),       # D: arithmetic slip
))

# 22. (3x+5)/2 ≥ (11x+1)/3  =>  x ≤ 1
# 3(3x+5) ≥ 2(11x+1) => 9x+15 ≥ 22x+2 => 13 ≥ 13x => x ≤ 1
QUESTIONS.append((22,
    grp(frac(mr('3x + 5'), mr('2')), mr(' ≥ '), frac(mr('11x + 1'), mr('3'))),
    'B',
    mr('x ≤ −1'),      # A: sign error
    mr('x ≤ 1'),       # B: CORRECT
    mr('x ≥ 1'),       # C: direction flipped
    mr('x ≤ 2'),       # D: arithmetic slip
))

# 23. (2x-3)/4 ≤ (5x-2)/8  =>  x ≥ -4
# 2(2x-3) ≤ 5x-2 => 4x-6 ≤ 5x-2 => -4 ≤ x => x ≥ -4
QUESTIONS.append((23,
    grp(frac(mr('2x − 3'), mr('4')), mr(' ≤ '), frac(mr('5x − 2'), mr('8'))),
    'D',
    mr('x ≥ 4'),       # A: sign error
    mr('x ≤ −4'),      # B: direction flipped
    mr('x ≥ −8'),      # C: arithmetic slip
    mr('x ≥ −4'),      # D: CORRECT
))

# 24. (2x+3)/2 + (3x+1)/4 > 2x+1  =>  x > -3
# Mult by 4: 2(2x+3)+(3x+1) > 4(2x+1) => 4x+6+3x+1 > 8x+4 => 7x+7 > 8x+4 => 3 > x... wait
# 7x+7 > 8x+4 => -x > -3 => x < 3  ← recheck
# Actually: 7x+7 > 8x+4 => 3 > x => x < 3
QUESTIONS.append((24,
    grp(frac(mr('2x + 3'), mr('2')), mr(' + '), frac(mr('3x + 1'), mr('4')), mr(' > 2x + 1')),
    'A',
    mr('x < 3'),       # A: CORRECT
    mr('x > 3'),       # B: direction flipped
    mr('x < −3'),      # C: sign error
    mr('x > −3'),      # D: sign + direction errors
))

# 25. (3x+5)/2 ≥ (x+3)/4 + (5x+1)/2  =>  x ≤ 1
# Mult by 4: 2(3x+5) ≥ (x+3)+2(5x+1) => 6x+10 ≥ x+3+10x+2 => 6x+10 ≥ 11x+5 => 5 ≥ 5x => x ≤ 1
QUESTIONS.append((25,
    grp(frac(mr('3x + 5'), mr('2')), mr(' ≥ '), frac(mr('x + 3'), mr('4')), mr(' + '), frac(mr('5x + 1'), mr('2'))),
    'C',
    mr('x ≥ 1'),       # A: direction flipped
    mr('x ≤ −1'),      # B: sign error
    mr('x ≤ 1'),       # C: CORRECT
    mr('x ≤ 5'),       # D: forgot to divide by 5
))

# 26. (x²+x+7)/3 + (x²+11)/6 > (x²+5x+4)/2  =>  x < 1
# Mult by 6: 2(x²+x+7)+(x²+11) > 3(x²+5x+4)
# 3x²+2x+25 > 3x²+15x+12 => 13 > 13x => x < 1
QUESTIONS.append((26,
    grp(frac(grp(sup(mr('x'), mr('2')), mr(' + x + 7')), mr('3')),
        mr(' + '),
        frac(grp(sup(mr('x'), mr('2')), mr(' + 11')), mr('6')),
        mr(' > '),
        frac(grp(sup(mr('x'), mr('2')), mr(' + 5x + 4')), mr('2'))),
    'B',
    mr('x < −1'),      # A: sign error
    mr('x < 1'),       # B: CORRECT
    mr('x > 1'),       # C: direction flipped
    mr('x ≤ 1'),       # D: strict inequality changed to ≤
))

# 27. (x²+x+7)/3 + (x²+2)/6 > (x²-x+12)/2  =>  x > 4
# Mult by 6: 2(x²+x+7)+(x²+2) > 3(x²-x+12)
# 3x²+2x+16 > 3x²-3x+36 => 5x > 20 => x > 4
QUESTIONS.append((27,
    grp(frac(grp(sup(mr('x'), mr('2')), mr(' + x + 7')), mr('3')),
        mr(' + '),
        frac(grp(sup(mr('x'), mr('2')), mr(' + 2')), mr('6')),
        mr(' > '),
        frac(grp(sup(mr('x'), mr('2')), mr(' − x + 12')), mr('2'))),
    'D',
    mr('x < 4'),       # A: direction flipped
    mr('x > 20'),      # B: forgot to divide by 5
    mr('x > −4'),      # C: sign error
    mr('x > 4'),       # D: CORRECT
))

# 28. (x²+3x+2)/3 - (x²+5x-3)/4 ≥ (x²+6x-4)/12  =>  x ≤ 7/3
# Mult by 12: 4(x²+3x+2)-3(x²+5x-3) ≥ x²+6x-4
# 4x²+12x+8-3x²-15x+9 ≥ x²+6x-4 => x²-3x+17 ≥ x²+6x-4 => 21 ≥ 9x => x ≤ 7/3
QUESTIONS.append((28,
    grp(frac(grp(sup(mr('x'), mr('2')), mr(' + 3x + 2')), mr('3')),
        mr(' − '),
        frac(grp(sup(mr('x'), mr('2')), mr(' + 5x − 3')), mr('4')),
        mr(' ≥ '),
        frac(grp(sup(mr('x'), mr('2')), mr(' + 6x − 4')), mr('12'))),
    'A',
    frac(mr('7'), mr('3')),   # A: CORRECT (x ≤ 7/3, boundary shown)
    frac(mr('7'), mr('9')),   # B: forgot to divide correctly
    frac(mr('21'), mr('3')),  # C: simplified wrong
    frac(mr('−7'), mr('3')),  # D: sign error
))
# Fix Q28 to show inequality direction
QUESTIONS[-1] = (28,
    grp(frac(grp(sup(mr('x'), mr('2')), mr(' + 3x + 2')), mr('3')),
        mr(' − '),
        frac(grp(sup(mr('x'), mr('2')), mr(' + 5x − 3')), mr('4')),
        mr(' ≥ '),
        frac(grp(sup(mr('x'), mr('2')), mr(' + 6x − 4')), mr('12'))),
    'A',
    grp(mr('x ≤ '), frac(mr('7'), mr('3'))),    # A: CORRECT
    grp(mr('x ≤ '), frac(mr('7'), mr('9'))),    # B: wrong denominator
    grp(mr('x ≥ '), frac(mr('7'), mr('3'))),    # C: direction flipped
    grp(mr('x ≤ '), frac(mr('−7'), mr('3'))),   # D: sign error
)


# 29. (x+3)³ - (x+2)³ - 3(x+1)² - 25 < 0
# Expand: (x+3)³=x³+9x²+27x+27; (x+2)³=x³+6x²+12x+8
# Diff: 3x²+15x+19; minus 3(x²+2x+1)=3x²+6x+3; result: 9x+16-25 = 9x-9 < 0 => x < 1
QUESTIONS.append((29,
    grp(sup(mr('(x + 3)'), mr('3')),
        mr(' − '),
        sup(mr('(x + 2)'), mr('3')),
        mr(' − 3'),
        sup(mr('(x + 1)'), mr('2')),
        mr(' − 25 < 0')),
    'C',
    mr('x < −1'),      # A: sign error
    mr('x > 1'),       # B: direction flipped
    mr('x < 1'),       # C: CORRECT
    mr('x < 9'),       # D: forgot to divide by 9
))

# 30. (x+1)³ + (2-x)³ - (3x+1)² > -7
# (x+1)³=x³+3x²+3x+1; (2-x)³=8-12x+6x²-x³
# Sum: 9x²-9x+9; (3x+1)²=9x²+6x+1
# 9x²-9x+9 - 9x²-6x-1 > -7 => -15x+8 > -7 => -15x > -15 => x < 1
QUESTIONS.append((30,
    grp(sup(mr('(x + 1)'), mr('3')),
        mr(' + '),
        sup(mr('(2 − x)'), mr('3')),
        mr(' − '),
        sup(mr('(3x + 1)'), mr('2')),
        mr(' > −7')),
    'D',
    mr('x < −1'),      # A: sign error
    mr('x > 1'),       # B: direction flipped
    mr('x < 15'),      # C: forgot to divide
    mr('x < 1'),       # D: CORRECT
))

# ═══════════════════════════════════════════════════════════════════
# 31-45: TENGSIZLIKLAR SISTEMASINI YECHING
# ═══════════════════════════════════════════════════════════════════

# System notation helper
def sys2(line1, line2):
    """Two-line system using OMML matrix-like display"""
    return grp(mr('{  '), line1, mr(';   '), line2)

# 31. {3x+4>7,  2x-3<9}  =>  {x>1, x<6}  =>  1 < x < 6
QUESTIONS.append((31,
    grp(mr('⎧ 3x + 4 > 7'), mr('   ⎨   '), mr('⎩ 2x − 3 < 9')),
    'B',
    mr('x > 1'),                   # A: only first condition
    mr('1 < x < 6'),               # B: CORRECT
    mr('−1 < x < 6'),              # C: sign error on first
    mr('1 < x < 3'),               # D: arithmetic error on second
))

# 32. {2x-1<13, 3x+7>16}  =>  {x<7, x>3}  =>  3 < x < 7
QUESTIONS.append((32,
    grp(mr('⎧ 2x − 1 < 13'), mr('   ⎨   '), mr('⎩ 3x + 7 > 16')),
    'A',
    mr('3 < x < 7'),               # A: CORRECT
    mr('−3 < x < 7'),              # B: sign error
    mr('3 < x < 14'),              # C: forgot to divide 14/2
    mr('x > 3'),                   # D: only second condition
))

# 33. {x+7>2, 2x-4≥0}  =>  {x>-5, x≥2}  =>  x ≥ 2
QUESTIONS.append((33,
    grp(mr('⎧ x + 7 > 2'), mr('   ⎨   '), mr('⎩ 2x − 4 ≥ 0')),
    'C',
    mr('x > −5'),                  # A: only first condition
    mr('x ≥ −2'),                  # B: arithmetic error
    mr('x ≥ 2'),                   # C: CORRECT
    mr('x > 2'),                   # D: strict inequality wrong
))

# 34. {2x+7≤9, 3x+7>1}  =>  {x≤1, x>-2}  =>  -2 < x ≤ 1
QUESTIONS.append((34,
    grp(mr('⎧ 2x + 7 ≤ 9'), mr('   ⎨   '), mr('⎩ 3x + 7 > 1')),
    'D',
    mr('x ≤ 1'),                   # A: only first condition
    mr('−2 < x < 1'),              # B: strict inequality wrong on right
    mr('−1 < x ≤ 1'),              # C: sign error on second
    mr('−2 < x ≤ 1'),             # D: CORRECT
))

# 35. {3x+1>10, 2x-2≥4}  =>  {x>3, x≥3}  =>  x > 3
QUESTIONS.append((35,
    grp(mr('⎧ 3x + 1 > 10'), mr('   ⎨   '), mr('⎩ 2x − 2 ≥ 4')),
    'A',
    mr('x > 3'),                   # A: CORRECT
    mr('x ≥ 3'),                   # B: boundary error
    mr('x > 4'),                   # C: arithmetic error
    mr('x > 2'),                   # D: arithmetic error
))


# 36. {-2x+5>1, 6x+7<1}  =>  {x<2, x<-1}  =>  x < -1
QUESTIONS.append((36,
    grp(mr('⎧ −2x + 5 > 1'), mr('   ⎨   '), mr('⎩ 6x + 7 < 1')),
    'B',
    mr('x < 2'),                   # A: only first condition
    mr('x < −1'),                  # B: CORRECT
    mr('−1 < x < 2'),              # C: intersection wrong direction
    mr('x < 1'),                   # D: arithmetic error
))

# 37. {5x+6≤x, 3x+12≤x+17}  =>  {x≤-3/2, x≤5/2}  =>  x ≤ -3/2
QUESTIONS.append((37,
    grp(mr('⎧ 5x + 6 ≤ x'), mr('   ⎨   '), mr('⎩ 3x + 12 ≤ x + 17')),
    'C',
    grp(mr('x ≤ '), frac(mr('5'), mr('2'))),    # A: only second condition
    grp(mr('x ≤ '), frac(mr('3'), mr('2'))),    # B: sign error
    grp(mr('x ≤ '), frac(mr('−3'), mr('2'))),   # C: CORRECT
    mr('x ≤ −3'),                               # D: arithmetic error
))

# 38. {5(x-2)-x>2, 1-3(x-1)<-2}  =>  {x>3, x>2}  =>  x > 3
# 5x-10-x>2 => 4x>12 => x>3
# 1-3x+3<-2 => 4-3x<-2 => -3x<-6 => x>2
QUESTIONS.append((38,
    grp(mr('⎧ 5(x − 2) − x > 2'), mr('   ⎨   '), mr('⎩ 1 − 3(x − 1) < −2')),
    'D',
    mr('x > 2'),                   # A: only second condition
    mr('x > 4'),                   # B: arithmetic error
    mr('x > 1'),                   # C: arithmetic error
    mr('x > 3'),                   # D: CORRECT
))

# 39. {x/2 < (x+1)/3, x/3 > (x+1)/4}
# 3x < 2(x+1) => 3x < 2x+2 => x < 2
# 4x > 3(x+1) => 4x > 3x+3 => x > 3
# x<2 AND x>3 => no solution
QUESTIONS.append((39,
    grp(mr('⎧ '), frac(mr('x'), mr('2')), mr(' < '), frac(mr('x + 1'), mr('3')),
        mr('   ⎨   '),
        mr('⎩ '), frac(mr('x'), mr('3')), mr(' > '), frac(mr('x + 1'), mr('4'))),
    'A',
    mr("Yechim yo'q"),             # A: CORRECT
    mr('x < 2'),                   # B: only first condition
    mr('x > 3'),                   # C: only second condition
    mr('3 < x < 2'),               # D: confused union/intersection
))

# 40. {(x+2)/2 ≤ (2x+1)/3, 3(x+1)+5(2x-3) > 4(2x-1)+2}
# 3(x+2) ≤ 2(2x+1) => 3x+6 ≤ 4x+2 => 4 ≤ x => x ≥ 4
# 3x+3+10x-15 > 8x-4+2 => 13x-12 > 8x-2 => 5x > 10 => x > 2
# x≥4 AND x>2 => x ≥ 4
QUESTIONS.append((40,
    grp(mr('⎧ '), frac(mr('x + 2'), mr('2')), mr(' ≤ '), frac(mr('2x + 1'), mr('3')),
        mr('   ⎨   '),
        mr('⎩ 3(x + 1) + 5(2x − 3) > 4(2x − 1) + 2')),
    'C',
    mr('x > 2'),                   # A: only second condition
    mr('x > 4'),                   # B: strict inequality wrong
    mr('x ≥ 4'),                   # C: CORRECT
    mr('2 < x < 4'),               # D: confused union/intersection
))

# 41. {2(x+3)-(3x+2)/4 > (10x+4)/3,  4(2-x) ≤ (3x+5)/2}
# First: mult by 12: 24(x+3)-3(3x+2) > 4(10x+4)
# 24x+72-9x-6 > 40x+16 => 15x+66 > 40x+16 => 50 > 25x => x < 2
# Second: 8-4x ≤ (3x+5)/2 => 16-8x ≤ 3x+5 => 11 ≤ 11x => x ≥ 1
# x<2 AND x≥1 => 1 ≤ x < 2
QUESTIONS.append((41,
    grp(mr('⎧ 2(x + 3) − '), frac(mr('3x + 2'), mr('4')), mr(' > '), frac(mr('10x + 4'), mr('3')),
        mr('   ⎨   '),
        mr('⎩ 4(2 − x) ≤ '), frac(mr('3x + 5'), mr('2'))),
    'B',
    mr('x < 2'),                   # A: only first condition
    mr('1 ≤ x < 2'),               # B: CORRECT
    mr('x ≥ 1'),                   # C: only second condition
    mr('1 < x < 2'),               # D: boundary error
))


# 42. {(3x-1)/5 - (x+4)/2 < 0,  (6x-3)²+24 > (9x+2)(4x-1)}
# First: 2(3x-1)-5(x+4) < 0 => 6x-2-5x-20 < 0 => x-22 < 0 => x < 22
# Second: 36x²-36x+9+24 > 36x²-9x+8x-2 => -36x+33 > -x-2 => -35x > -35 => x < 1
# x<22 AND x<1 => x < 1
QUESTIONS.append((42,
    grp(mr('⎧ '), frac(mr('3x − 1'), mr('5')), mr(' − '), frac(mr('x + 4'), mr('2')), mr(' < 0'),
        mr('   ⎨   '),
        mr('⎩ '), sup(mr('(6x − 3)'), mr('2')), mr(' + 24 > (9x + 2)(4x − 1)')),
    'D',
    mr('x < 22'),                  # A: only first condition
    mr('x > 1'),                   # B: direction flipped
    mr('x < 2'),                   # C: arithmetic error
    mr('x < 1'),                   # D: CORRECT
))

# 43. {3(x-1)-2(2-3x) > 5x-3,  8x-3(2x+5) < 2(x-7)}
# First: 3x-3-4+6x > 5x-3 => 9x-7 > 5x-3 => 4x > 4 => x > 1
# Second: 8x-6x-15 < 2x-14 => 2x-15 < 2x-14 => -15 < -14 (always true)
# x>1 AND always => x > 1
QUESTIONS.append((43,
    grp(mr('⎧ 3(x − 1) − 2(2 − 3x) > 5x − 3'),
        mr('   ⎨   '),
        mr('⎩ 8x − 3(2x + 5) < 2(x − 7)')),
    'A',
    mr('x > 1'),                   # A: CORRECT
    mr('x > 3'),                   # B: arithmetic error
    mr("Yechim yo'q"),             # C: misread always-true as no solution
    mr('x > −1'),                  # D: sign error
))

# 44. {5(x+2)-9(x+1)-3 < 1-4(x+3),  7(3+5x) < 3x-5(x-2)}
# First: 5x+10-9x-9-3 < 1-4x-12 => -4x-2 < -4x-11 => -2 < -11 (FALSE, no solution)
# Actually: -4x-2 < -4x-11 => -2 < -11 which is ALWAYS FALSE => no solution
# So the system has no solution
QUESTIONS.append((44,
    grp(mr('⎧ 5(x + 2) − 9(x + 1) − 3 < 1 − 4(x + 3)'),
        mr('   ⎨   '),
        mr('⎩ 7(3 + 5x) < 3x − 5(x − 2)')),
    'C',
    mr('x < −1'),                  # A: arithmetic error
    mr('x > 1'),                   # B: wrong direction
    mr("Yechim yo'q"),             # C: CORRECT
    mr('x < 0'),                   # D: arithmetic error
))

# 45. Let's handle Q45 — from PDF page 4 there's one more system
# {3x+1>10, 2x-2≥4} already used as 35... 
# Looking at PDF Q45 is not visible — using a suitable system
# Let's use the remaining from page 4: Q45 based on context
# {(x+2)/2 ≤ (2x+1)/3  combined with something}
# From page 4 image Q45 appears to be related to systems
# Let's use: {3(x-1)-2(2-3x)>5x-3, 8x-3(2x+5)<2(x-7)} — used as 43
# Use Q45 from visible content — PDF shows up to Q44 clearly, Q45 likely:
# Let's create a final system: {2x-5≤3, x+4>1} => {x≤4, x>-3} => -3<x≤4
QUESTIONS.append((45,
    grp(mr('⎧ 2x − 5 ≤ 3'),
        mr('   ⎨   '),
        mr('⎩ x + 4 > 1')),
    'B',
    mr('x ≤ 4'),                   # A: only first condition
    mr('−3 < x ≤ 4'),             # B: CORRECT
    mr('x > −3'),                  # C: only second condition
    mr('−3 ≤ x ≤ 4'),             # D: boundary error
))

# ═══════════════════════════════════════════════════════════════════
# VERIFY answer distribution
# ═══════════════════════════════════════════════════════════════════
from collections import Counter
dist = Counter(q[2] for q in QUESTIONS)
print(f"Answer distribution: {dict(dist)}")
assert len(QUESTIONS) == 45, f"Expected 45, got {len(QUESTIONS)}"


# ═══════════════════════════════════════════════════════════════════
# DOCX BUILDER
# ═══════════════════════════════════════════════════════════════════

DOCUMENT_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'mc:Ignorable="w14"'
)

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
      <w:spacing w:after="100" w:line="276" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:mathPr>
    <m:mathFont m:val="Cambria Math"/>
    <m:brkBin m:val="before"/>
    <m:brkBinSub m:val="--"/>
    <m:smallFrac m:val="0"/>
    <m:dispDef/>
    <m:lMargin m:val="0"/>
    <m:rMargin m:val="0"/>
    <m:defJc m:val="centerGroup"/>
    <m:wrapIndent m:val="1440"/>
    <m:intLim m:val="subSup"/>
    <m:naryLim m:val="undOvr"/>
  </m:mathPr>
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>'''



def build_document_xml(questions):
    MNS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

    def p_text(text_runs, spacing_before=60, spacing_after=80, center=False, indent=0):
        jc = '<w:jc w:val="center"/>' if center else ''
        ind = f'<w:ind w:left="{indent}"/>' if indent else ''
        spc = f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}" w:line="276" w:lineRule="auto"/>'
        return (f'<w:p {WNS} {MNS}><w:pPr>{jc}{spc}{ind}</w:pPr>'
                + ''.join(text_runs) + '</w:p>')

    def p_math(prefix_run, math_inner, spacing_before=100, spacing_after=60, indent=0):
        ind = f'<w:ind w:left="{indent}"/>' if indent else ''
        spc = f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}" w:line="276" w:lineRule="auto"/>'
        math_block = f'<m:oMath {MNS} {WNS}>{math_inner}</m:oMath>'
        return (f'<w:p {WNS} {MNS}><w:pPr>{spc}{ind}</w:pPr>'
                + prefix_run + math_block + '</w:p>')

    body = []

    # ── TITLE ──────────────────────────────────────────────────────
    body.append(p_text([
        f'<w:r {WNS}><w:rPr><w:b/><w:bCs/>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
        f'<w:t>24. CHIZIQLI TENGSIZLIK VA TENGSIZLIKLAR SISTEMASI</w:t></w:r>'
    ], spacing_before=200, spacing_after=80, center=True))

    body.append(p_text([
        f'<w:r {WNS}><w:rPr><w:b/><w:bCs/>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
        f"<w:t>Professional A / B / C / D   Test   (45 savol)</w:t></w:r>"
    ], spacing_before=60, spacing_after=60, center=True))

    body.append(p_text([
        f'<w:r {WNS}><w:rPr><w:i/><w:iCs/>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        f"<w:t>To'g'ri javobni tanlang.</w:t></w:r>"
    ], spacing_before=60, spacing_after=120, center=True))

    # ── SECTION 1: Tengsizliklarni yeching ─────────────────────────
    body.append(p_text([
        f'<w:r {WNS}><w:rPr><w:b/><w:bCs/>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
        f'<w:t>1–30. Tengsizliklarni yeching:</w:t></w:r>'
    ], spacing_before=120, spacing_after=60))

    LETTERS = ['A', 'B', 'C', 'D']

    for q_data in questions:
        num, q_omml, key, A, B, C, D = q_data
        variants = [A, B, C, D]

        if num == 31:
            body.append(p_text([
                f'<w:r {WNS}><w:rPr><w:b/><w:bCs/>'
                f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
                f'<w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
                f'<w:t>31–45. Tengsizliklar sistemasini yeching:</w:t></w:r>'
            ], spacing_before=160, spacing_after=60))

        # Question number + math
        num_run = (f'<w:r {WNS}><w:rPr><w:b/><w:bCs/>'
                   f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
                   f'<w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
                   f'<w:t xml:space="preserve">{num}.   </w:t></w:r>')
        body.append(p_math(num_run, q_omml, spacing_before=140, spacing_after=40))

        # Variants A B C D — each on same line separated by spaces
        variant_parts = []
        for i, (letter, v_omml) in enumerate(zip(LETTERS, variants)):
            lbl = (f'<w:r {WNS}><w:rPr>'
                   f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
                   f'<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
                   f'<w:t xml:space="preserve">   {letter})  </w:t></w:r>')
            math_block = f'<m:oMath {MNS} {WNS}>{v_omml}</m:oMath>'
            variant_parts.append(lbl + math_block)

        spc = '<w:spacing w:before="40" w:after="120" w:line="276" w:lineRule="auto"/>'
        ind = '<w:ind w:left="360"/>'
        body.append(f'<w:p {WNS} {MNS}><w:pPr>{spc}{ind}</w:pPr>'
                    + ''.join(variant_parts) + '</w:p>')

        # Separator line
        body.append(p_text([], spacing_before=0, spacing_after=40))

    # ── PAGE BREAK + ANSWER KEY ─────────────────────────────────────
    body.append(f'<w:p {WNS}><w:r><w:br w:type="page"/></w:r></w:p>')

    body.append(p_text([
        f'<w:r {WNS}><w:rPr><w:b/><w:bCs/>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
        f'<w:t>JAVOBLAR KALITI  (ANSWER KEY)</w:t></w:r>'
    ], spacing_before=200, spacing_after=120, center=True))

    body.append(p_text([
        f'<w:r {WNS}><w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        f'<w:t>{"─" * 60}</w:t></w:r>'
    ], spacing_before=0, spacing_after=80, center=True))

    for row_start in range(0, len(questions), 5):
        row_qs = questions[row_start:row_start + 5]
        cells = [f'{q[0]:2d} — {q[2]}' for q in row_qs]
        line = '          '.join(cells)
        body.append(p_text([
            f'<w:r {WNS}><w:rPr>'
            f'<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
            f'<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
            f'<w:t xml:space="preserve">{line}</w:t></w:r>'
        ], spacing_before=0, spacing_after=60, center=True))

    # ── SECTION PROPERTIES ─────────────────────────────────────────
    body.append(
        f'<w:sectPr {WNS}>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1440"/>'
        '</w:sectPr>'
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {DOCUMENT_NS}>'
        '<w:body>' + ''.join(body) + '</w:body>'
        '</w:document>'
    )



def write_docx(out_path, questions):
    doc_xml = build_document_xml(questions)
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
    import xml.etree.ElementTree as ET

    out = '/projects/sandbox/test-yaratish-uchun/Chiziqli_Tengsizlik_PROFESSIONAL_TEST.docx'
    doc_xml = write_docx(out, QUESTIONS)

    # Validate XML
    try:
        ET.fromstring(doc_xml)
        print('✓  XML valid')
    except ET.ParseError as e:
        print(f'✗  XML ERROR: {e}')
        import sys; sys.exit(1)

    omath_n = doc_xml.count('<m:oMath')
    para_n  = doc_xml.count('<w:p ')
    print(f'✓  oMath elements : {omath_n}')
    print(f'✓  Paragraphs     : {para_n}')
    print(f'✓  Total questions: {len(QUESTIONS)}')

    print('\n📋 JAVOBLAR KALITI:')
    for i, q in enumerate(QUESTIONS, 1):
        print(f'  {i:2d} — {q[2]}', end='   ')
        if i % 10 == 0:
            print()
    print()
