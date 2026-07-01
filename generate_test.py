#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional Math Test Generator - Linear Equations
Creates a .docx file with OMML equations for all 90 questions
"""

import zipfile
import os
import re

def omml(xml):
    """Wrap raw OMML math XML in proper namespace declarations"""
    return xml

def make_run(text, bold=False, size=24):
    b = "<w:b/><w:bCs/>" if bold else ""
    return f"""<w:r><w:rPr>{b}<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"""

def math_run(omml_inner):
    """Wrap OMML inner content in oMath tags"""
    return f"""<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{omml_inner}</m:oMath>"""

def para(content, style="Normal", indent=0):
    ind = f'<w:ind w:left="{indent}"/>' if indent else ""
    return f"""<w:p><w:pPr><w:pStyle w:val="{style}"/><w:spacing w:after="120"/>{ind}</w:pPr>{content}</w:p>"""

def page_break():
    return """<w:p><w:r><w:br w:type="page"/></w:r></w:p>"""

# OMML helper functions
def frac(num, den):
    return f"""<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num><m:r><m:t>{num}</m:t></m:r></m:num><m:den><m:r><m:t>{den}</m:t></m:r></m:den></m:f>"""

def frac_expr(num_xml, den_xml):
    return f"""<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>{num_xml}</m:num><m:den>{den_xml}</m:den></m:f>"""

def mrun(text):
    return f"""<m:r><m:t xml:space="preserve">{text}</m:t></m:r>"""

def sup_script(base, exp):
    return f"""<m:sSup><m:sSupPr/><m:e><m:r><m:t>{base}</m:t></m:r></m:e><m:sup><m:r><m:t>{exp}</m:t></m:r></m:sup></m:sSup>"""

def paren_run(inner):
    return f"""<m:d><m:dPr><m:begChr m:val="("/>
<m:endChr m:val=")"/></m:dPr><m:e>{inner}</m:e></m:d>"""



# ============================================================
# QUESTIONS DATA: (question_text_omml, correct_answer, A, B, C, D)
# correct_answer is 'A','B','C','D'
# Each variant is OMML inner XML
# ============================================================

def q(num, label, omml_q, key, A, B, C, D):
    return {"num": num, "label": label, "q": omml_q, "key": key,
            "A": A, "B": B, "C": C, "D": D}

questions = []

# ---------- Q1: 19x = 57 => x=3 ----------
questions.append(q(1, "19x = 57",
    f"{mrun('19x = 57')}",
    "C",
    mrun("x = 1"), mrun("x = 2"), mrun("x = 3"), mrun("x = 4")
))

# ---------- Q2: -48x = 96 => x=-2 ----------
questions.append(q(2, "-48x = 96",
    f"{mrun('−48x = 96')}",
    "B",
    mrun("x = 2"), mrun("x = −2"), mrun("x = −3"), mrun("x = 3")
))

# ---------- Q3: -23x = 69 => x=-3 ----------
questions.append(q(3, "-23x = 69",
    f"{mrun('−23x = 69')}",
    "D",
    mrun("x = 3"), mrun("x = −2"), mrun("x = −4"), mrun("x = −3")
))

# ---------- Q4: -105 : x = -35 => x=3 ----------
questions.append(q(4, "-105 : x = -35",
    f"{mrun('−105 ÷ x = −35')}",
    "A",
    mrun("x = 3"), mrun("x = −3"), mrun("x = 5"), mrun("x = −5")
))

# ---------- Q5: 10 : (-x) = 2 => x=-5 ----------
questions.append(q(5, "10:(-x)=2",
    f"{mrun('10 ÷ (−x) = 2')}",
    "C",
    mrun("x = 5"), mrun("x = 2"), mrun("x = −5"), mrun("x = −2")
))

# ---------- Q6: 50:(-x)=5 => x=-10 ----------
questions.append(q(6, "50:(-x)=5",
    f"{mrun('50 ÷ (−x) = 5')}",
    "B",
    mrun("x = 10"), mrun("x = −10"), mrun("x = −5"), mrun("x = 5")
))

# ---------- Q7: 12+x=-9 => x=-21 ----------
questions.append(q(7, "12+x=-9",
    f"{mrun('12 + x = −9')}",
    "D",
    mrun("x = 21"), mrun("x = −3"), mrun("x = 3"), mrun("x = −21")
))

# ---------- Q8: -8+x=10 => x=18 ----------
questions.append(q(8, "-8+x=10",
    f"{mrun('−8 + x = 10')}",
    "A",
    mrun("x = 18"), mrun("x = 2"), mrun("x = −18"), mrun("x = −2")
))

# ---------- Q9: 12x+14=38 => x=2 ----------
questions.append(q(9, "12x+14=38",
    f"{mrun('12x + 14 = 38')}",
    "B",
    mrun("x = 3"), mrun("x = 2"), mrun("x = 4"), mrun("x = 1")
))

# ---------- Q10: 24-3x=15 => x=3 ----------
questions.append(q(10, "24-3x=15",
    f"{mrun('24 − 3x = 15')}",
    "C",
    mrun("x = 13"), mrun("x = −3"), mrun("x = 3"), mrun("x = 5")
))



# ---------- Q11: 6x-12=24 => x=6 ----------
questions.append(q(11, "6x-12=24",
    f"{mrun('6x − 12 = 24')}",
    "A",
    mrun("x = 6"), mrun("x = 2"), mrun("x = −6"), mrun("x = 4")
))

# ---------- Q12: 8x+4=40 => x=4.5 ----------
questions.append(q(12, "8x+4=40",
    f"{mrun('8x + 4 = 40')}",
    "D",
    mrun("x = 5"), mrun("x = 3"), mrun("x = 6"), mrun("x = 4.5")
))

# ---------- Q13: (-4x+4)/4=-4 => x=5 ----------
questions.append(q(13, "(-4x+4)/4=-4",
    f"{frac(mrun('−4x + 4'), mrun('4'))}{mrun(' = −4')}",
    "B",
    mrun("x = 3"), mrun("x = 5"), mrun("x = −5"), mrun("x = 4")
))

# ---------- Q14: (-x-5)/2=-3 => x=1 ----------
questions.append(q(14, "(-x-5)/2=-3",
    f"{frac(mrun('−x − 5'), mrun('2'))}{mrun(' = −3')}",
    "C",
    mrun("x = −1"), mrun("x = 11"), mrun("x = 1"), mrun("x = −11")
))

# ---------- Q15: 4(-7x+2)=-20 => x=4/7... actually x=1 ----------
# 4(-7x+2)=-20 => -7x+2=-5 => -7x=-7 => x=1
questions.append(q(15, "4(-7x+2)=-20",
    f"{mrun('4(−7x + 2) = −20')}",
    "A",
    mrun("x = 1"), mrun("x = −1"), mrun("x = 2"), mrun("x = 3")
))

# ---------- Q16: 6(-x-5)=-42 => x=2 ----------
# -x-5=-7 => -x=-2 => x=2
questions.append(q(16, "6(-x-5)=-42",
    f"{mrun('6(−x − 5) = −42')}",
    "D",
    mrun("x = −2"), mrun("x = 7"), mrun("x = −7"), mrun("x = 2")
))

# ---------- Q17: -(x+4)-4(3x+1)=-21 => x=1 ----------
# -x-4-12x-4=-21 => -13x-8=-21 => -13x=-13 => x=1
questions.append(q(17, "-(x+4)-4(3x+1)=-21",
    f"{mrun('−(x + 4) − 4(3x + 1) = −21')}",
    "B",
    mrun("x = −1"), mrun("x = 1"), mrun("x = 2"), mrun("x = −2")
))

# ---------- Q18: 5(-3x+3)-2(-3x-2)=19 => x=0 ----------
# -15x+15+6x+4=19 => -9x+19=19 => x=0
questions.append(q(18, "5(-3x+3)-2(-3x-2)=19",
    f"{mrun('5(−3x + 3) − 2(−3x − 2) = 19')}",
    "C",
    mrun("x = 1"), mrun("x = −1"), mrun("x = 0"), mrun("x = 2")
))

# ---------- Q19: 3x+2(x-4)=12 => x=4 ----------
# 3x+2x-8=12 => 5x=20 => x=4
questions.append(q(19, "3x+2(x-4)=12",
    f"{mrun('3x + 2(x − 4) = 12')}",
    "A",
    mrun("x = 4"), mrun("x = 2"), mrun("x = −4"), mrun("x = 5")
))

# ---------- Q20: 5(1-2x)-12=35-9x => x=-46 ----------
# 5-10x-12=35-9x => -7-10x=35-9x => -x=42 => x=-42
questions.append(q(20, "5(1-2x)-12=35-9x",
    f"{mrun('5(1 − 2x) − 12 = 35 − 9x')}",
    "D",
    mrun("x = 42"), mrun("x = −46"), mrun("x = 46"), mrun("x = −42")
))



# ---------- Q21: 3(2x-1)-2(1-2x)=9(x-2) => x=1 ----------
# 6x-3-2+4x=9x-18 => 10x-5=9x-18 => x=-13
questions.append(q(21, "3(2x-1)-2(1-2x)=9(x-2)",
    f"{mrun('3(2x − 1) − 2(1 − 2x) = 9(x − 2)')}",
    "B",
    mrun("x = 13"), mrun("x = −13"), mrun("x = −7"), mrun("x = 7")
))

# ---------- Q22: 2(3x+1)+4(2x-3)=3(x-1) => x=1 ----------
# 6x+2+8x-12=3x-3 => 14x-10=3x-3 => 11x=7 => x=7/11
questions.append(q(22, "2(3x+1)+4(2x-3)=3(x-1)",
    f"{mrun('2(3x + 1) + 4(2x − 3) = 3(x − 1)')}",
    "A",
    f"{frac(mrun('7'), mrun('11'))}", mrun("x = 1"), mrun("x = −1"), f"{frac(mrun('−7'), mrun('11'))}"
))

# ---------- Q23: 7(3x+2)-3(4x-2)=5(2x-1) => x=13/3 ----------
# 21x+14-12x+6=10x-5 => 9x+20=10x-5 => x=25
questions.append(q(23, "7(3x+2)-3(4x-2)=5(2x-1)",
    f"{mrun('7(3x + 2) − 3(4x − 2) = 5(2x − 1)')}",
    "C",
    mrun("x = 20"), mrun("x = −25"), mrun("x = 25"), mrun("x = 5")
))

# ---------- Q24: 3(2-4x)+1=-19-2(1-x) => x=3 ----------
# 6-12x+1=-19-2+2x => 7-12x=-21+2x => 28=14x => x=2
questions.append(q(24, "3(2-4x)+1=-19-2(1-x)",
    f"{mrun('3(2 − 4x) + 1 = −19 − 2(1 − x)')}",
    "D",
    mrun("x = −2"), mrun("x = 3"), mrun("x = −3"), mrun("x = 2")
))

# ---------- Q25: -3(4x+4)-(3x+2)=(4x-3)+(-5x+1)+16 => x=-2 ----------
# -12x-12-3x-2=4x-3-5x+1+16 => -15x-14=-x+14 => -14x=28 => x=-2
questions.append(q(25, "-3(4x+4)-(3x+2)=(4x-3)+(-5x+1)+16",
    f"{mrun('−3(4x + 4) − (3x + 2) = (4x − 3) + (−5x + 1) + 16')}",
    "A",
    mrun("x = −2"), mrun("x = 2"), mrun("x = −1"), mrun("x = 1")
))

# ---------- Q26: -4(5x-3)-5(2x+2)=3(-5x+4)-(5x+1)-9 => x=0 ----------
# -20x+12-10x-10=-15x+12-5x-1-9 => -30x+2=-20x+2 => -10x=0 => x=0
questions.append(q(26, "-4(5x-3)-5(2x+2)=3(-5x+4)-(5x+1)-9",
    f"{mrun('−4(5x − 3) − 5(2x + 2) = 3(−5x + 4) − (5x + 1) − 9')}",
    "C",
    mrun("x = 1"), mrun("x = −1"), mrun("x = 0"), mrun("x = 2")
))

# ---------- Q27: 4(x-2)+3x=6(3/2 x-4)+12 => x=2 ----------
# 4x-8+3x=9x-24+12 => 7x-8=9x-12 => -2x=-4 => x=2
_q27_inner = frac(mrun('3'), mrun('2')) + mrun('x − 4')
_q27_paren = paren_run(_q27_inner)
_q27_expr = mrun('4(x − 2) + 3x = 6') + _q27_paren + mrun(' + 12')
questions.append(q(27, "4(x-2)+3x=6(3/2 x-4)+12",
    _q27_expr,
    "B",
    mrun("x = 4"), mrun("x = 2"), mrun("x = −2"), mrun("x = 8")
))

# ---------- Q28: 2.5(3-x)-1.5=4(x+1)-11 => x=1 ----------
# 7.5-2.5x-1.5=4x+4-11 => 6-2.5x=4x-7 => 13=6.5x => x=2
questions.append(q(28, "2.5(3-x)-1.5=4(x+1)-11",
    f"{mrun('2,5(3 − x) − 1,5 = 4(x + 1) − 11')}",
    "D",
    mrun("x = −2"), mrun("x = 1"), mrun("x = 3"), mrun("x = 2")
))

# ---------- Q29: x/2-3=1/2 => x=7 ----------
questions.append(q(29, "x/2-3=1/2",
    f"{frac(mrun('x'), mrun('2'))}{mrun(' − 3 = ')}{frac(mrun('1'), mrun('2'))}",
    "A",
    mrun("x = 7"), mrun("x = 5"), mrun("x = 3"), mrun("x = 9")
))

# ---------- Q30: (x+2)/3-2=1/3 => x=3 ----------
questions.append(q(30, "(x+2)/3-2=1/3",
    f"{frac(mrun('x + 2'), mrun('3'))}{mrun(' − 2 = ')}{frac(mrun('1'), mrun('3'))}",
    "C",
    mrun("x = 5"), mrun("x = 1"), mrun("x = 3"), mrun("x = −1")
))



# ---------- Q31: (2x+1)/3-3=3/4 => x=57/8 ----------
# (2x+1)/3=3+3/4=15/4 => 2x+1=45/4 => 2x=41/4 => x=41/8
questions.append(q(31, "(2x+1)/3-3=3/4",
    f"{frac(mrun('2x + 1'), mrun('3'))}{mrun(' − 3 = ')}{frac(mrun('3'), mrun('4'))}",
    "B",
    f"{frac(mrun('35'), mrun('8'))}", f"{frac(mrun('41'), mrun('8'))}", f"{frac(mrun('45'), mrun('8'))}", f"{frac(mrun('37'), mrun('8'))}"
))

# ---------- Q32: (2x+5)/6-3/4=1/2 => x=1 ----------
# (2x+5)/6=1/2+3/4=5/4 => 2x+5=30/4=7.5 => 2x=2.5 => x=1.25
questions.append(q(32, "(2x+5)/6-3/4=1/2",
    f"{frac(mrun('2x + 5'), mrun('6'))}{mrun(' − ')}{frac(mrun('3'), mrun('4'))}{mrun(' = ')}{frac(mrun('1'), mrun('2'))}",
    "D",
    mrun("x = 2"), mrun("x = 3"), mrun("x = 1"), f"{frac(mrun('5'), mrun('4'))}"
))

# ---------- Q33: (x-1)/2+(3x-5)/3=5/6 => x=2 ----------
# 3(x-1)+2(3x-5)=5 => 3x-3+6x-10=5 => 9x-13=5 => 9x=18 => x=2
questions.append(q(33, "(x-1)/2+(3x-5)/3=5/6",
    f"{frac(mrun('x − 1'), mrun('2'))}{mrun(' + ')}{frac(mrun('3x − 5'), mrun('3'))}{mrun(' = ')}{frac(mrun('5'), mrun('6'))}",
    "A",
    mrun("x = 2"), mrun("x = 3"), mrun("x = 1"), mrun("x = −2")
))

# ---------- Q34: (-3x-7)/6+(4x+7)/6=5 => x=23 ----------
# (-3x-7+4x+7)/6=5 => x/6=5 => x=30
questions.append(q(34, "(-3x-7)/6+(4x+7)/6=5",
    f"{frac(mrun('−3x − 7'), mrun('6'))}{mrun(' + ')}{frac(mrun('4x + 7'), mrun('6'))}{mrun(' = 5')}",
    "C",
    mrun("x = 23"), mrun("x = 25"), mrun("x = 30"), mrun("x = 35")
))

# ---------- Q35: (3x+4)/(-2)-(−x+6)/2=-1 => x=1 ----------
# -(3x+4)/2+(x-6)/2=-1 => (-3x-4+x-6)/2=-1 => -2x-10=-2 => -2x=8 => x=-4
questions.append(q(35, "(3x+4)/(-2)-(−x+6)/2=-1",
    f"{frac(mrun('3x + 4'), mrun('−2'))}{mrun(' − ')}{frac(mrun('−x + 6'), mrun('2'))}{mrun(' = −1')}",
    "B",
    mrun("x = 4"), mrun("x = −4"), mrun("x = 1"), mrun("x = −1")
))

# ---------- Q36: (2x+1)/3+(5x+1)/6=(7x-3)/4+1 => x=9 ----------
# LCD=12: 4(2x+1)+2(5x+1)=3(7x-3)+12 => 8x+4+10x+2=21x-9+12 => 18x+6=21x+3 => -3x=-3 => x=1
questions.append(q(36, "(2x+1)/3+(5x+1)/6=(7x-3)/4+1",
    f"{frac(mrun('2x + 1'), mrun('3'))}{mrun(' + ')}{frac(mrun('5x + 1'), mrun('6'))}{mrun(' = ')}{frac(mrun('7x − 3'), mrun('4'))}{mrun(' + 1')}",
    "D",
    mrun("x = 9"), mrun("x = 3"), mrun("x = −1"), mrun("x = 1")
))

# ---------- Q37: (x+4)/9+(4-x)/6=1/2+(2x-2)/9 => x=0 ----------
# LCD=18: 2(x+4)+3(4-x)=9+2(2x-2) => 2x+8+12-3x=9+4x-4 => -x+20=4x+5 => -5x=-15 => x=3
questions.append(q(37, "(x+4)/9+(4-x)/6=1/2+(2x-2)/9",
    f"{frac(mrun('x + 4'), mrun('9'))}{mrun(' + ')}{frac(mrun('4 − x'), mrun('6'))}{mrun(' = ')}{frac(mrun('1'), mrun('2'))}{mrun(' + ')}{frac(mrun('2x − 2'), mrun('9'))}",
    "A",
    mrun("x = 3"), mrun("x = −3"), mrun("x = 1"), mrun("x = −1")
))

# ---------- Q38: (x-3)/6-(x-2)/3=5/12-(x-1)/4 => x=2 ----------
# LCD=12: 2(x-3)-4(x-2)=5-3(x-1) => 2x-6-4x+8=5-3x+3 => -2x+2=-3x+8 => x=6
questions.append(q(38, "(x-3)/6-(x-2)/3=5/12-(x-1)/4",
    f"{frac(mrun('x − 3'), mrun('6'))}{mrun(' − ')}{frac(mrun('x − 2'), mrun('3'))}{mrun(' = ')}{frac(mrun('5'), mrun('12'))}{mrun(' − ')}{frac(mrun('x − 1'), mrun('4'))}",
    "C",
    mrun("x = 2"), mrun("x = −6"), mrun("x = 6"), mrun("x = −2")
))

# ---------- Q39: 2(x-4)/3-(x-2)/2=5/3-x/2 => x=4 ----------
# LCD=6: 4(x-4)-3(x-2)=10-3x => 4x-16-3x+6=10-3x => x-10=10-3x => 4x=20 => x=5
_q39_num = mrun('2(x − 4)')
_q39_expr = frac(_q39_num, mrun('3')) + mrun(' − ') + frac(mrun('x − 2'), mrun('2')) + mrun(' = ') + frac(mrun('5'), mrun('3')) + mrun(' − ') + frac(mrun('x'), mrun('2'))
questions.append(q(39, "2(x-4)/3-(x-2)/2=5/3-x/2",
    _q39_expr,
    "B",
    mrun("x = 4"), mrun("x = 5"), mrun("x = 3"), mrun("x = 6")
))

# ---------- Q40: (7x+4)/(-2)-((-6x+4)/4)=5 => x=? ----------
# -(7x+4)/2+(6x-4)/4=5 => LCD=4: -2(7x+4)+(6x-4)=20 => -14x-8+6x-4=20 => -8x=32 => x=-4
questions.append(q(40, "(7x+4)/(-2)-(-6x+4)/4=5",
    f"{frac(mrun('7x + 4'), mrun('−2'))}{mrun(' − ')}{frac(mrun('−6x + 4'), mrun('4'))}{mrun(' = 5')}",
    "D",
    mrun("x = 4"), mrun("x = −2"), mrun("x = 2"), mrun("x = −4")
))



# ---------- Q41: (2x+3)/(3x+2)=1 => x=1 ----------
# 2x+3=3x+2 => x=1
questions.append(q(41, "(2x+3)/(3x+2)=1",
    f"{frac(mrun('2x + 3'), mrun('3x + 2'))}{mrun(' = 1')}",
    "A",
    mrun("x = 1"), mrun("x = −1"), mrun("x = 2"), mrun("x = −2")
))

# ---------- Q42: (-4x-6)/(-x-3)=10 => x=? ----------
# -4x-6=10(-x-3)=-10x-30 => 6x=-24 => x=-4
questions.append(q(42, "(-4x-6)/(-x-3)=10",
    f"{frac(mrun('−4x − 6'), mrun('−x − 3'))}{mrun(' = 10')}",
    "C",
    mrun("x = 4"), mrun("x = −2"), mrun("x = −4"), mrun("x = 2")
))

# ---------- Q43: -2/(-3x+6)=8/(-5x-7) => x=? ----------
# -2(-5x-7)=8(-3x+6) => 10x+14=-24x+48 => 34x=34 => x=1
questions.append(q(43, "-2/(-3x+6)=8/(-5x-7)",
    f"{frac(mrun('−2'), mrun('−3x + 6'))}{mrun(' = ')}{frac(mrun('8'), mrun('−5x − 7'))}",
    "B",
    mrun("x = −1"), mrun("x = 1"), mrun("x = 2"), mrun("x = −2")
))

# ---------- Q44: 1/(-5x-6)=2/(3x-3) => x=? ----------
# 3x-3=-2(5x+6)=-10x-12 => 13x=-9 => x=-9/13
questions.append(q(44, "1/(-5x-6)=2/(3x-3)",
    f"{frac(mrun('1'), mrun('−5x − 6'))}{mrun(' = ')}{frac(mrun('2'), mrun('3x − 3'))}",
    "D",
    f"{frac(mrun('9'), mrun('13'))}", f"{frac(mrun('−7'), mrun('13'))}", f"{frac(mrun('7'), mrun('13'))}", f"{frac(mrun('−9'), mrun('13'))}"
))

# ---------- Q45: 4x+(5x-(6x-(7x-(8x-9))))=10 => x=? ----------
# innermost: 8x-9; 7x-(8x-9)=7x-8x+9=-x+9; 6x-(-x+9)=7x-9; 5x-(7x-9)=-2x+9; 4x+(-2x+9)=2x+9=10 => x=0.5
questions.append(q(45, "4x+(5x-(6x-(7x-(8x-9))))=10",
    f"{mrun('4x + (5x − (6x − (7x − (8x − 9)))) = 10')}",
    "A",
    f"{frac(mrun('1'), mrun('2'))}", mrun("x = 1"), mrun("x = 2"), mrun("x = −1")
))

# ---------- Q46: 5x-(3x+(2x-(4x+(6x-12))))=18 => x=? ----------
# 6x-12; 4x+(6x-12)=10x-12; 2x-(10x-12)=-8x+12; 3x+(-8x+12)=-5x+12; 5x-(-5x+12)=10x-12=18 => 10x=30 => x=3
questions.append(q(46, "5x-(3x+(2x-(4x+(6x-12))))=18",
    f"{mrun('5x − (3x + (2x − (4x + (6x − 12)))) = 18')}",
    "C",
    mrun("x = 1"), mrun("x = 2"), mrun("x = 3"), mrun("x = 4")
))

# ---------- Q47: (x+2)^2-(x-1)^2=21 => x=? ----------
# x²+4x+4-(x²-2x+1)=21 => 6x+3=21 => 6x=18 => x=3
questions.append(q(47, "(x+2)^2-(x-1)^2=21",
    f"{sup_script('(x+2)', '2')}{mrun('−')}{sup_script('(x−1)', '2')}{mrun(' = 21')}",
    "B",
    mrun("x = 2"), mrun("x = 3"), mrun("x = 4"), mrun("x = −3")
))

# ---------- Q48: (2x-1)^2-(2x-3)(2x+3)=10 => x=? ----------
# 4x²-4x+1-(4x²-9)=10 => -4x+10=10 => x=0
questions.append(q(48, "(2x-1)^2-(2x-3)(2x+3)=10",
    f"{sup_script('(2x−1)', '2')}{mrun(' − (2x − 3)(2x + 3) = 10')}",
    "D",
    mrun("x = 1"), mrun("x = −1"), mrun("x = 2"), mrun("x = 0")
))

# ---------- Q49: (x-1)(x+1)-(x+2)^2=-9 => x=? ----------
# x²-1-(x²+4x+4)=-9 => -4x-5=-9 => -4x=-4 => x=1
questions.append(q(49, "(x-1)(x+1)-(x+2)^2=-9",
    f"{mrun('(x − 1)(x + 1) − ')}{sup_script('(x+2)', '2')}{mrun(' = −9')}",
    "A",
    mrun("x = 1"), mrun("x = −1"), mrun("x = 2"), mrun("x = 3")
))

# ---------- Q50: (x+1)^2=(111-(1-x))x-80 => x=? ----------
# x²+2x+1=(111-1+x)x-80=(110+x)x-80=110x+x²-80 => 2x+1=110x-80 => -108x=-81 => x=81/108=3/4
questions.append(q(50, "(x+1)^2=(111-(1-x))x-80",
    f"{sup_script('(x+1)', '2')}{mrun(' = (111 − (1 − x))x − 80')}",
    "C",
    f"{frac(mrun('1'), mrun('2'))}", f"{frac(mrun('2'), mrun('3'))}", f"{frac(mrun('3'), mrun('4'))}", f"{frac(mrun('4'), mrun('5'))}"
))



# ---------- Q51: (2x-3)/(4x+5)=(3x-1)/(6x-1) => x=? ----------
# (2x-3)(6x-1)=(3x-1)(4x+5) => 12x²-2x-18x+3=12x²+15x-4x-5 => -20x+3=11x-5 => -31x=-8 => x=8/31
questions.append(q(51, "(2x-3)/(4x+5)=(3x-1)/(6x-1)",
    f"{frac(mrun('2x − 3'), mrun('4x + 5'))}{mrun(' = ')}{frac(mrun('3x − 1'), mrun('6x − 1'))}",
    "B",
    f"{frac(mrun('−8'), mrun('31'))}", f"{frac(mrun('8'), mrun('31'))}", f"{frac(mrun('7'), mrun('31'))}", f"{frac(mrun('−7'), mrun('31'))}"
))

# ---------- Q52: (x-1)/(3x-2)=(5x-3)/(15x-1) => x=? ----------
# (x-1)(15x-1)=(5x-3)(3x-2) => 15x²-x-15x+1=15x²-10x-9x+6 => -16x+1=-19x+6 => 3x=5 => x=5/3
questions.append(q(52, "(x-1)/(3x-2)=(5x-3)/(15x-1)",
    f"{frac(mrun('x − 1'), mrun('3x − 2'))}{mrun(' = ')}{frac(mrun('5x − 3'), mrun('15x − 1'))}",
    "D",
    f"{frac(mrun('−5'), mrun('3'))}", f"{frac(mrun('1'), mrun('3'))}", f"{frac(mrun('−1'), mrun('3'))}", f"{frac(mrun('5'), mrun('3'))}"
))

# ---------- Q53: (3x-1)/(6x-5)=(x+1)/(2x+3) => x=? ----------
# (3x-1)(2x+3)=(x+1)(6x-5) => 6x²+9x-2x-3=6x²-5x+6x-5 => 7x-3=x-5 => 6x=-2 => x=-1/3
questions.append(q(53, "(3x-1)/(6x-5)=(x+1)/(2x+3)",
    f"{frac(mrun('3x − 1'), mrun('6x − 5'))}{mrun(' = ')}{frac(mrun('x + 1'), mrun('2x + 3'))}",
    "A",
    f"{frac(mrun('−1'), mrun('3'))}", f"{frac(mrun('1'), mrun('3'))}", f"{frac(mrun('−2'), mrun('3'))}", f"{frac(mrun('2'), mrun('3'))}"
))

# ---------- Q54: (5x²+x+5)/(x-4)=5x+3 => x=? ----------
_q54_num = mrun('5') + sup_script('x', '2') + mrun(' + x + 5')
_q54_expr = frac(_q54_num, mrun('x − 4')) + mrun(' = 5x + 3')
questions.append(q(54, "(5x^2+x+5)/(x-4)=5x+3",
    _q54_expr,
    "C",
    frac(mrun('17'), mrun('18')), frac(mrun('−15'), mrun('18')), frac(mrun('−17'), mrun('18')), frac(mrun('15'), mrun('18'))
))

# ---------- Q55: (x+2 2/9)/(3 1/4)=4 => x=? ----------
_q55_num = mrun('x + 2') + frac(mrun('2'), mrun('9'))
_q55_den = mrun('3') + frac(mrun('1'), mrun('4'))
_q55_expr = frac(_q55_num, _q55_den) + mrun(' = 4')
questions.append(q(55, "(x+2 2/9)/(3 1/4)=4",
    _q55_expr,
    "B",
    mrun('10') + frac(mrun('3'), mrun('9')),
    mrun('10') + frac(mrun('7'), mrun('9')),
    mrun('9') + frac(mrun('7'), mrun('9')),
    mrun('11') + frac(mrun('7'), mrun('9'))
))

# ---------- Q56: (5 4/5+x)/(5 1/3)=3 => x=? ----------
_q56_num = mrun('5') + frac(mrun('4'), mrun('5')) + mrun(' + x')
_q56_den = mrun('5') + frac(mrun('1'), mrun('3'))
_q56_expr = frac(_q56_num, _q56_den) + mrun(' = 3')
questions.append(q(56, "(5 4/5+x)/(5 1/3)=3",
    _q56_expr,
    "D",
    mrun('10') + frac(mrun('4'), mrun('5')),
    mrun('9') + frac(mrun('1'), mrun('5')),
    mrun('11') + frac(mrun('1'), mrun('5')),
    mrun('10') + frac(mrun('1'), mrun('5'))
))

# ---------- Q57: (1+1/0.2)/(1+1/x)=36 => x=? ----------
_q57_num = mrun('1 + ') + frac(mrun('1'), mrun('0,2'))
_q57_den = mrun('1 + ') + frac(mrun('1'), mrun('x'))
_q57_expr = frac(_q57_num, _q57_den) + mrun(' = 36')
questions.append(q(57, "(1+1/0.2)/(1+1/x)=36",
    _q57_expr,
    "A",
    frac(mrun('−6'), mrun('5')), frac(mrun('6'), mrun('5')), frac(mrun('−5'), mrun('6')), frac(mrun('5'), mrun('6'))
))

# ---------- Q58: 5-(5/x+1)/(3-1/x)=2 => x=? ----------
_q58_num = frac(mrun('5'), mrun('x')) + mrun(' + 1')
_q58_den = mrun('3 − ') + frac(mrun('1'), mrun('x'))
_q58_expr = mrun('5 − ') + frac(_q58_num, _q58_den) + mrun(' = 2')
questions.append(q(58, "5-(5/x+1)/(3-1/x)=2",
    _q58_expr,
    "C",
    mrun("x = −1"), mrun("x = 2"), mrun("x = 1"), mrun("x = −2")
))



# ---------- Q59: 100/(25-25/x)=5 => x=? ----------
# 25-25/x=20 => 25/x=5 => x=5
_q59_den = mrun('25 − ') + frac(mrun('25'), mrun('x'))
questions.append(q(59, "100/(25-25/x)=5",
    frac(mrun('100'), _q59_den) + mrun(' = 5'),
    "B",
    mrun("x = 4"), mrun("x = 5"), mrun("x = 6"), mrun("x = −5")
))

# ---------- Q60: 300/(7-75/x)=75 => x=? ----------
# 7-75/x=4 => 75/x=3 => x=25
_q60_den = mrun('7 − ') + frac(mrun('75'), mrun('x'))
questions.append(q(60, "300/(7-75/x)=75",
    frac(mrun('300'), _q60_den) + mrun(' = 75'),
    "D",
    mrun("x = 15"), mrun("x = 20"), mrun("x = 30"), mrun("x = 25")
))

# ---------- Q61: 2x-3=2(x+1)-5 => all x (cheksiz) => identity ----------
# 2x-3=2x+2-5=2x-3 => identity
questions.append(q(61, "2x-3=2(x+1)-5",
    f"{mrun('2x − 3 = 2(x + 1) − 5')}",
    "A",
    mrun("Har qanday x"), mrun("x = 0"), mrun("x = 1"), mrun("x = −1")
))

# ---------- Q62: 3x+2(x-3)=6x-(x+4) => x=2 ----------
# 3x+2x-6=6x-x-4 => 5x-6=5x-4 => -6=-4 => yechim yo'q
questions.append(q(62, "3x+2(x-3)=6x-(x+4)",
    f"{mrun('3x + 2(x − 3) = 6x − (x + 4)')}",
    "C",
    mrun("x = 2"), mrun("x = −2"), mrun("Yechim yo'q"), mrun("Har qanday x")
))

# ---------- Q63: (x+2)/3+(x-5)/6=(2x-3)/4+15 => x=? ----------
# LCD=12: 4(x+2)+2(x-5)=3(2x-3)+180 => 4x+8+2x-10=6x-9+180 => 6x-2=6x+171 => -2=171 => yechim yo'q
questions.append(q(63, "(x+2)/3+(x-5)/6=(2x-3)/4+15",
    f"{frac(mrun('x + 2'), mrun('3'))}{mrun(' + ')}{frac(mrun('x − 5'), mrun('6'))}{mrun(' = ')}{frac(mrun('2x − 3'), mrun('4'))}{mrun(' + 15')}",
    "B",
    mrun("x = 15"), mrun("Yechim yo'q"), mrun("x = 0"), mrun("Har qanday x")
))

# ---------- Q64: (2x-1)/3-(4-x)/2-x=1+(x-3)/6 => x=? ----------
# LCD=6: 2(2x-1)-3(4-x)-6x=6+( x-3) => 4x-2-12+3x-6x=6+x-3 => x-14=x+3 => -14=3 => yechim yo'q
questions.append(q(64, "(2x-1)/3-(4-x)/2-x=1+(x-3)/6",
    f"{frac(mrun('2x − 1'), mrun('3'))}{mrun(' − ')}{frac(mrun('4 − x'), mrun('2'))}{mrun(' − x = 1 + ')}{frac(mrun('x − 3'), mrun('6'))}",
    "D",
    mrun("x = 3"), mrun("x = −3"), mrun("Har qanday x"), mrun("Yechim yo'q")
))

# ---------- Q65: 2(x-3)-3(2-x)+8=5x-4 => x=? ----------
# 2x-6-6+3x+8=5x-4 => 5x-4=5x-4 => Har qanday x
questions.append(q(65, "2(x-3)-3(2-x)+8=5x-4",
    f"{mrun('2(x − 3) − 3(2 − x) + 8 = 5x − 4')}",
    "A",
    mrun("Har qanday x"), mrun("x = 0"), mrun("Yechim yo'q"), mrun("x = 4")
))



# ============================================================
# PARAMETRIC EQUATIONS (Q66-Q90)
# ============================================================

# ---------- Q66: 3xa-7=7a+5x => x=1 bo'lganda a=? ----------
# 3a-7=7a+5 => -4a=12 => a=-3
questions.append(q(66, "3xa−7=7a+5x => x=1 bo'lganda a=?",
    f"{mrun('3xa − 7 = 7a + 5x,  x = 1 da  a = ?')}",
    "C",
    mrun("a = 3"), mrun("a = −1"), mrun("a = −3"), mrun("a = 1")
))

# ---------- Q67: 2x+5a=4x-3+ax => x=12 bo'lganda a=? ----------
# 24+5a=48-3+12a => 24+5a=45+12a => -7a=21 => a=-3
questions.append(q(67, "2x+5a=4x−3+ax => x=12 bo'lganda a=?",
    f"{mrun('2x + 5a = 4x − 3 + ax,  x = 12 da  a = ?')}",
    "B",
    mrun("a = 3"), mrun("a = −3"), mrun("a = 2"), mrun("a = −2")
))

# ---------- Q68: a(x+1)-8=x(a-2)+3a => x=5 bo'lganda a=? ----------
# 5a+a-8=5a-10+3a => 6a-8=8a-10 => -2a=-2 => a=1
questions.append(q(68, "a(x+1)−8=x(a−2)+3a => x=5 bo'lganda a=?",
    f"{mrun('a(x + 1) − 8 = x(a − 2) + 3a,  x = 5 da  a = ?')}",
    "D",
    mrun("a = −1"), mrun("a = 2"), mrun("a = −2"), mrun("a = 1")
))

# ---------- Q69: 5x(a-x)+3x=8a+4+2x => x=2 bo'lganda a=? ----------
# 10(a-2)+6=8a+4+4 => 10a-20+6=8a+8 => 10a-14=8a+8 => 2a=22 => a=11
questions.append(q(69, "5x(a−x)+3x=8a+4+2x => x=2 bo'lganda a=?",
    f"{mrun('5x(a − x) + 3x = 8a + 4 + 2x,  x = 2 da  a = ?')}",
    "A",
    mrun("a = 11"), mrun("a = −11"), mrun("a = 9"), mrun("a = −9")
))

# ---------- Q70: 2x(x+a)-ax=a+3x => x=2 bo'lganda a=? ----------
# 4(2+a)-2a=a+6 => 8+4a-2a=a+6 => 8+2a=a+6 => a=-2
questions.append(q(70, "2x(x+a)−ax=a+3x => x=2 bo'lganda a=?",
    f"{mrun('2x(x + a) − ax = a + 3x,  x = 2 da  a = ?')}",
    "C",
    mrun("a = 2"), mrun("a = 4"), mrun("a = −2"), mrun("a = −4")
))

# ---------- Q71: y+2x=6+3y => x ni y orqali ----------
# 2x=6+2y => x=3+y
questions.append(q(71, "y+2x=6+3y (x ni y orqali)",
    f"{mrun('y + 2x = 6 + 3y')}",
    "B",
    mrun("x = 3 − y"), mrun("x = 3 + y"), mrun("x = 6 + y"), mrun("x = 6 − y")
))

# ---------- Q72: yx+3=5x+y => x ni y orqali ----------
# x(y-5)=y-3 => x=(y-3)/(y-5)
questions.append(q(72, "yx+3=5x+y (x ni y orqali)",
    f"{mrun('yx + 3 = 5x + y')}",
    "D",
    f"{frac(mrun('y + 3'), mrun('y − 5'))}", f"{frac(mrun('y − 3'), mrun('y + 5'))}", f"{frac(mrun('y + 3'), mrun('y + 5'))}", f"{frac(mrun('y − 3'), mrun('y − 5'))}"
))

# ---------- Q73: y=(3x+1)/(2x-1) => x ni y orqali ----------
# y(2x-1)=3x+1 => 2xy-y=3x+1 => x(2y-3)=y+1 => x=(y+1)/(2y-3)
questions.append(q(73, "y=(3x+1)/(2x-1) (x ni y orqali)",
    f"{mrun('y = ')}{frac(mrun('3x + 1'), mrun('2x − 1'))}",
    "A",
    f"{frac(mrun('y + 1'), mrun('2y − 3'))}", f"{frac(mrun('y − 1'), mrun('2y − 3'))}", f"{frac(mrun('y + 1'), mrun('2y + 3'))}", f"{frac(mrun('y − 1'), mrun('2y + 3'))}"
))

# ---------- Q74: y=(2x+3)/(x+1) => x ni y orqali ----------
# y(x+1)=2x+3 => xy+y=2x+3 => x(y-2)=3-y => x=(3-y)/(y-2)
questions.append(q(74, "y=(2x+3)/(x+1) (x ni y orqali)",
    f"{mrun('y = ')}{frac(mrun('2x + 3'), mrun('x + 1'))}",
    "C",
    f"{frac(mrun('3 + y'), mrun('y − 2'))}", f"{frac(mrun('3 + y'), mrun('y + 2'))}", f"{frac(mrun('3 − y'), mrun('y − 2'))}", f"{frac(mrun('3 − y'), mrun('y + 2'))}"
))

# ---------- Q75: y=(5x+1)/(x+2) => x ni y orqali ----------
# y(x+2)=5x+1 => xy+2y=5x+1 => x(y-5)=1-2y => x=(1-2y)/(y-5)
questions.append(q(75, "y=(5x+1)/(x+2) (x ni y orqali)",
    f"{mrun('y = ')}{frac(mrun('5x + 1'), mrun('x + 2'))}",
    "B",
    f"{frac(mrun('1 + 2y'), mrun('y − 5'))}", f"{frac(mrun('1 − 2y'), mrun('y − 5'))}", f"{frac(mrun('2y − 1'), mrun('y − 5'))}", f"{frac(mrun('1 − 2y'), mrun('y + 5'))}"
))



# ---------- Q76: ax+5=4x-b yechimga ega bo'lmaydi => a=4, b≠-5 ----------
# ax-4x=-b-5 => (a-4)x=-b-5 => yechim yo'q: a=4, b≠-5
questions.append(q(76, "ax+5=4x−b: yechimga ega bo'lmaydi => a,b=?",
    f"{mrun('ax + 5 = 4x − b  (yechimga ega bo`lmaydi)')}",
    "A",
    mrun("a=4, b≠−5"), mrun("a≠4, b=5"), mrun("a=4, b=−5"), mrun("a=−4, b=5")
))

# ---------- Q77: 5x+2=ax-4 yechimga ega bo'lmaydi => a=5 ----------
# (5-a)x=-6 => yechim yo'q: a=5
questions.append(q(77, "5x+2=ax−4: yechimga ega bo'lmaydi => a=?",
    f"{mrun('5x + 2 = ax − 4  (yechimga ega bo`lmaydi)')}",
    "C",
    mrun("a = −5"), mrun("a = 4"), mrun("a = 5"), mrun("a = −4")
))

# ---------- Q78: a(x+5)=3(x+1) yechimga ega bo'lmaydi => a=3 ----------
# ax+5a=3x+3 => (a-3)x=3-5a => yechim yo'q: a=3, 3-15≠0 => 3≠15, ok
questions.append(q(78, "a(x+5)=3(x+1): yechimga ega bo'lmaydi => a=?",
    f"{mrun('a(x + 5) = 3(x + 1)  (yechimga ega bo`lmaydi)')}",
    "B",
    mrun("a = −3"), mrun("a = 3"), mrun("a = 1"), mrun("a = −1")
))

# ---------- Q79: a=(3x-1)/(x-2) yechimga ega bo'lmaydi => a=? ----------
# a(x-2)=3x-1 => ax-2a=3x-1 => (a-3)x=2a-1 => yechim yo'q: a=3
questions.append(q(79, "a=(3x−1)/(x−2): yechimga ega bo'lmaydi => a=?",
    f"{mrun('a = ')}{frac(mrun('3x − 1'), mrun('x − 2'))}{mrun('  (yechimga ega bo`lmaydi)')}",
    "D",
    mrun("a = 1"), mrun("a = −1"), mrun("a = −3"), mrun("a = 3")
))

# ---------- Q80: a=(2x-3)/(x+10) yechimga ega bo'lmaydi => a=2 ----------
# a(x+10)=2x-3 => ax+10a=2x-3 => (a-2)x=-10a-3 => yechim yo'q: a=2
questions.append(q(80, "a=(2x−3)/(x+10): yechimga ega bo'lmaydi => a=?",
    f"{mrun('a = ')}{frac(mrun('2x − 3'), mrun('x + 10'))}{mrun('  (yechimga ega bo`lmaydi)')}",
    "A",
    mrun("a = 2"), mrun("a = −2"), mrun("a = 3"), mrun("a = −3")
))

# ---------- Q81: ax+5=4x-b cheksiz ko'p yechim => a=4, b=-5 ----------
questions.append(q(81, "ax+5=4x−b: cheksiz ko'p yechim => a,b=?",
    f"{mrun('ax + 5 = 4x − b  (cheksiz ko`p yechim)')}",
    "C",
    mrun("a=4, b=5"), mrun("a=−4, b=−5"), mrun("a=4, b=−5"), mrun("a=−4, b=5")
))

# ---------- Q82: ax-3=7x-b cheksiz ko'p yechim => a=7, b=3 ----------
questions.append(q(82, "ax−3=7x−b: cheksiz ko'p yechim => a,b=?",
    f"{mrun('ax − 3 = 7x − b  (cheksiz ko`p yechim)')}",
    "B",
    mrun("a=7, b=−3"), mrun("a=7, b=3"), mrun("a=−7, b=3"), mrun("a=−7, b=−3")
))

# ---------- Q83: (3a-6)x+2b=6x+8 cheksiz ko'p yechim => a=4, b=4 ----------
# 3a-6=6 => a=4; 2b=8 => b=4
questions.append(q(83, "(3a−6)x+2b=6x+8: cheksiz ko'p yechim => a,b=?",
    f"{mrun('(3a − 6)x + 2b = 6x + 8  (cheksiz ko`p yechim)')}",
    "D",
    mrun("a=2, b=4"), mrun("a=4, b=2"), mrun("a=2, b=2"), mrun("a=4, b=4")
))

# ---------- Q84: a²x+2=4x+a cheksiz ko'p yechim => a=2 or a=-2 ----------
# a²=4 => a=±2; 2=a => a=2
questions.append(q(84, "a²x+2=4x+a: cheksiz ko'p yechim => a=?",
    f"{sup_script('a', '2')}{mrun('x + 2 = 4x + a  (cheksiz ko`p yechim)')}",
    "A",
    mrun("a = 2"), mrun("a = −2"), mrun("a = 4"), mrun("a = −4")
))

# ---------- Q85: a²x-a=9x+3 cheksiz ko'p yechim => a=? ----------
# a²=9 => a=±3; -a=3 => a=-3
questions.append(q(85, "a²x−a=9x+3: cheksiz ko'p yechim => a=?",
    f"{sup_script('a', '2')}{mrun('x − a = 9x + 3  (cheksiz ko`p yechim)')}",
    "C",
    mrun("a = 3"), mrun("a = 9"), mrun("a = −3"), mrun("a = −9")
))

# ---------- Q86: ax+x=3(x-4) 1 ta yechim => a≠2 ----------
# x(a+1)=3x-12 => x(a+1-3)=-12 => x(a-2)=-12 => 1 yechim: a≠2
questions.append(q(86, "ax+x=3(x−4): 1 ta yechim => a=?",
    f"{mrun('ax + x = 3(x − 4)  (1 ta yechim)')}",
    "B",
    mrun("a = 2"), mrun("a ≠ 2"), mrun("a = −2"), mrun("a ≠ −2")
))

# ---------- Q87: 3a(x-4)=9x-2 1 ta yechim => a≠3 ----------
# 3ax-12a=9x-2 => (3a-9)x=12a-2 => 1 yechim: a≠3
questions.append(q(87, "3a(x−4)=9x−2: 1 ta yechim => a=?",
    f"{mrun('3a(x − 4) = 9x − 2  (1 ta yechim)')}",
    "D",
    mrun("a = 3"), mrun("a = −3"), mrun("a = 9"), mrun("a ≠ 3")
))

# ---------- Q88: a=(x-2)/(2x-1) 1 ta yechim => a≠1/2 ----------
# a(2x-1)=x-2 => 2ax-a=x-2 => (2a-1)x=a-2 => 1 yechim: a≠1/2
questions.append(q(88, "a=(x−2)/(2x−1): 1 ta yechim => a=?",
    f"{mrun('a = ')}{frac(mrun('x − 2'), mrun('2x − 1'))}{mrun('  (1 ta yechim)')}",
    "A",
    f"{mrun('a ≠ ')}{frac(mrun('1'), mrun('2'))}", f"{mrun('a = ')}{frac(mrun('1'), mrun('2'))}", f"{mrun('a ≠ ')}{frac(mrun('−1'), mrun('2'))}", f"{mrun('a = ')}{frac(mrun('−1'), mrun('2'))}"
))

# ---------- Q89: a=(2x-1)/(x-2) 1 ta yechim => a≠2 ----------
# a(x-2)=2x-1 => ax-2a=2x-1 => (a-2)x=2a-1 => 1 yechim: a≠2
questions.append(q(89, "a=(2x−1)/(x−2): 1 ta yechim => a=?",
    f"{mrun('a = ')}{frac(mrun('2x − 1'), mrun('x − 2'))}{mrun('  (1 ta yechim)')}",
    "C",
    mrun("a = 2"), mrun("a = −2"), mrun("a ≠ 2"), mrun("a ≠ −2")
))

# ---------- Q90: a(x+2)-2=2x 1 ta yechim => a≠2 ----------
# ax+2a-2=2x => (a-2)x=2-2a => 1 yechim: a≠2
questions.append(q(90, "a(x+2)−2=2x: 1 ta yechim => a=?",
    f"{mrun('a(x + 2) − 2 = 2x  (1 ta yechim)')}",
    "B",
    mrun("a = 2"), mrun("a ≠ 2"), mrun("a = −2"), mrun("a ≠ −2")
))



# ============================================================
# DOCX BUILDER
# ============================================================

NS = {
    'wpc': "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    'cx': "http://schemas.microsoft.com/office/drawing/2014/chartex",
    'mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
    'aink': "http://schemas.microsoft.com/office/drawing/2016/ink",
    'am3d': "http://schemas.microsoft.com/office/drawing/2017/model3d",
    'o': "urn:schemas-microsoft-com:office:office",
    'oel': "http://schemas.microsoft.com/office/2019/extlst",
    'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    'm': "http://schemas.openxmlformats.org/officeDocument/2006/math",
    'v': "urn:schemas-microsoft-com:vml",
    'wp14': "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    'wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    'w10': "urn:schemas-microsoft-com:office:word",
    'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    'w14': "http://schemas.microsoft.com/office/word/2010/wordml",
    'w15': "http://schemas.microsoft.com/office/word/2012/wordml",
    'w16cex': "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    'w16cid': "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    'w16': "http://schemas.microsoft.com/office/word/2018/wordml",
    'w16sdtdh': "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    'w16se': "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    'wpg': "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    'wpi': "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    'wne': "http://schemas.microsoft.com/office/word/2006/wordml",
    'wps': "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}

def build_document_xml(questions):
    body_parts = []

    # Title
    body_parts.append(f"""<w:p>
  <w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr>
  <w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
    <w:t>CHIZIQLI TENGLAMALAR — TEST</w:t>
  </w:r>
</w:p>""")

    body_parts.append(f"""<w:p>
  <w:pPr><w:jc w:val="center"/><w:spacing w:after="300"/></w:pPr>
  <w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
    <w:t>To'g'ri javobni tanlang</w:t>
  </w:r>
</w:p>""")

    for item in questions:
        num = item["num"]
        key = item["key"]
        q_omml = item["q"]
        variants = {"A": item["A"], "B": item["B"], "C": item["C"], "D": item["D"]}

        # Question paragraph
        body_parts.append(f"""<w:p>
  <w:pPr><w:spacing w:before="160" w:after="60"/><w:ind w:left="0"/></w:pPr>
  <w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
    <w:t xml:space="preserve">{num}. </w:t>
  </w:r>
  <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    {q_omml}
  </m:oMath>
</w:p>""")

        # Variants paragraph
        variant_parts = []
        for letter in ["A", "B", "C", "D"]:
            v_omml = variants[letter]
            variant_parts.append(f"""<w:r><w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
    <w:t xml:space="preserve">  {letter})  </w:t>
  </w:r>
  <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    {v_omml}
  </m:oMath>""")

        body_parts.append(f"""<w:p>
  <w:pPr><w:spacing w:after="80"/><w:ind w:left="360"/></w:pPr>
  {''.join(variant_parts)}
</w:p>""")

    # PAGE BREAK before answer key
    body_parts.append("""<w:p><w:r><w:br w:type="page"/></w:r></w:p>""")

    # Answer Key title
    body_parts.append(f"""<w:p>
  <w:pPr><w:jc w:val="center"/><w:spacing w:after="300"/></w:pPr>
  <w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
    <w:t>ANSWER KEY — JAVOBLAR KALITI</w:t>
  </w:r>
</w:p>""")

    # Answer key rows — 3 columns
    keys_per_row = 3
    rows = []
    i = 0
    while i < len(questions):
        row_items = questions[i:i+keys_per_row]
        row_text = "     ".join([f"{it['num']}  —  {it['key']}" for it in row_items])
        rows.append(f"""<w:p>
  <w:pPr><w:spacing w:after="80"/></w:pPr>
  <w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
    <w:t xml:space="preserve">{row_text}</w:t>
  </w:r>
</w:p>""")
        i += keys_per_row

    body_parts.extend(rows)

    return "\n".join(body_parts)




def build_docx(output_path, questions):
    import zipfile, os

    body_content = build_document_xml(questions)

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  mc:Ignorable="w14 w15">
<w:body>
{body_content}
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
           w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

    word_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
        <w:lang w:val="uz-Latn-UZ" w:eastAsia="uz-Latn-UZ" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr>
      <w:spacing w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>"""

    settings_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</w:settings>"""

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('word/_rels/document.xml.rels', word_rels)
        zf.writestr('word/styles.xml', styles_xml)
        zf.writestr('word/settings.xml', settings_xml)

    print(f"✅ Created: {output_path}")
    print(f"   Total questions: {len(questions)}")
    print(f"   File size: {os.path.getsize(output_path):,} bytes")


if __name__ == "__main__":
    output = "/projects/sandbox/test-yaratish-uchun/Chiziqli_Tenglamalar_Test.docx"
    build_docx(output, questions)
    
    # Print answer key summary
    print("\n📋 ANSWER KEY:")
    for i, item in enumerate(questions, 1):
        print(f"  {i:2d} — {item['key']}", end="   ")
        if i % 10 == 0:
            print()
    print()
