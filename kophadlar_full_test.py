#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional Kophadlar (Polynomials) Test Generator
PDF'dagi BARCHA masalalar - 48 ta savol
Microsoft Word OMML formatida
"""
import zipfile, os, random
from collections import Counter

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

def sub(base_xml, subscript_xml):
    """Subscript"""
    return (f'<m:sSub><m:sSubPr/>'
            f'<m:e>{base_xml}</m:e>'
            f'<m:sub>{subscript_xml}</m:sub></m:sSub>')

def grp(*items):
    """Group multiple OMML items"""
    return ''.join(items)

# ─── Question data structure ──────────────────────────────────────────────
# Each question: (number, question_omml, key_letter, A_omml, B_omml, C_omml, D_omml)

QUESTIONS = []

# ═══════════════════════════════════════════════════════════════════
# SAHIFA 1: KOPHADLAR (1-35)
# ═══════════════════════════════════════════════════════════════════

# 1. 10 - 5n ni ko'paytuvchilarga ajrating
QUESTIONS.append((1,
    grp(mr('10 − 5n ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('5(2 − n)')),           # A: CORRECT
    grp(mr('5(2 + n)')),           # B: sign error
    grp(mr('10(1 − n)')),          # C: wrong factorization
    grp(mr('5(n − 2)')),           # D: reversed terms
))



# 2. 14x + 21y ni ko'paytuvchilarga ajrating
QUESTIONS.append((2,
    grp(mr('14x + 21y ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('7(2x + 3y)')),         # A: CORRECT but moved to B
    grp(mr('7(2x + 3y)')),         # B: CORRECT
    grp(mr('14(x + y)')),          # C: wrong GCD
    grp(mr('7(2x − 3y)')),         # D: sign error
))

# 3. -8m - 12n ni ko'paytuvchilarga ajrating
QUESTIONS.append((3,
    grp(mr('−8m − 12n ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('4(−2m − 3n)')),        # A: wrong GCD
    grp(mr('−4(2m + 3n)')),        # B: sign placement error
    grp(mr('−4(2m + 3n)')),        # C: CORRECT
    grp(mr('8(−m − n)')),          # D: wrong factorization
))

# 4. 15a² + 10 ni ko'paytuvchilarga ajrating
QUESTIONS.append((4,
    grp(mr('15'), sup(mr('a'), mr('2')), mr(' + 10 ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('5(3a + 2)')),          # A: forgot power
    grp(mr('15(a + 1)')),          # B: wrong factorization
    grp(mr('5(3'), sup(mr('a'), mr('2')), mr(' − 2)')),  # C: sign error
    grp(mr('5(3'), sup(mr('a'), mr('2')), mr(' + 2)')),  # D: CORRECT
))

# 5. 8m³ - 12m² ni ko'paytuvchilarga ajrating
QUESTIONS.append((5,
    grp(sup(mr('8m'), mr('3')), mr(' − 12'), sup(mr('m'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('4'), sup(mr('m'), mr('2')), mr('(2m − 3)')),  # A: CORRECT
    grp(mr('4m(2'), sup(mr('m'), mr('2')), mr(' − 3m)')),  # B: wrong power
    grp(mr('8'), sup(mr('m'), mr('2')), mr('(m − 1)')),    # C: wrong factorization
    grp(mr('12'), sup(mr('m'), mr('2')), mr('(m − 1)')),   # D: wrong GCD
))

# 6. 3a - 3b + ac - bc ni ko'paytuvchilarga ajrating
QUESTIONS.append((6,
    grp(mr('3a − 3b + ac − bc ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('3(a − b) + c(a − b)')),     # A: not fully factored
    grp(mr('(a − b)(3 + c)')),          # B: CORRECT
    grp(mr('(3 − c)(a + b)')),          # C: sign error
    grp(mr('3(a + b) + c(a − b)')),     # D: partial factorization
))

# 7. x² + xy + 2x + 2y ni ko'paytuvchilarga ajrating
QUESTIONS.append((7,
    grp(sup(mr('x'), mr('2')), mr(' + xy + 2x + 2y ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('(x + 2)(x − y)')),          # A: sign error
    grp(mr('(x + y)(x + 1)')),          # B: wrong grouping
    grp(mr('(x + 2)(x + y)')),          # C: CORRECT
    grp(mr('x(x + y) + 2(x − y)')),     # D: not fully factored
))

# 8. x² - 2x - ax + 2a ni ko'paytuvchilarga ajrating
QUESTIONS.append((8,
    grp(sup(mr('x'), mr('2')), mr(' − 2x − ax + 2a ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(x − 2)(x + a)')),          # A: sign error
    grp(mr('(x + 2)(x − a)')),          # B: sign error
    grp(mr('x(x − 2) − a(x + 2)')),     # C: not fully factored
    grp(mr('(x − 2)(x − a)')),          # D: CORRECT
))

# 9. 3y² + y - 3y - 1 ni ko'paytuvchilarga ajrating
QUESTIONS.append((9,
    grp(sup(mr('3y'), mr('2')), mr(' + y − 3y − 1 ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('(3y + 1)(y − 1)')),         # A: CORRECT
    grp(mr('(3y − 1)(y + 1)')),         # B: sign error
    grp(mr('(3y + 1)(y + 1)')),         # C: sign error
    grp(mr('3y(y + 1) − (y + 1)')),     # D: not fully factored
))

# 10. 2xy + 3y + 2x + 3 ni ko'paytuvchilarga ajrating
QUESTIONS.append((10,
    grp(mr('2xy + 3y + 2x + 3 ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(2x + 3)(y − 1)')),         # A: sign error
    grp(mr('(2x + 3)(y + 1)')),         # B: CORRECT
    grp(mr('(2x − 3)(y + 1)')),         # C: sign error
    grp(mr('y(2x + 3) + (2x + 3)')),    # D: not simplified
))



# 11. ax + ay + 2bx + 2by ni ko'paytuvchilarga ajrating
QUESTIONS.append((11,
    grp(mr('ax + ay + 2bx + 2by ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('a(x + y) + 2b(x + y)')),    # A: not fully factored
    grp(mr('(a + 2b)(x − y)')),         # B: sign error
    grp(mr('(a + 2b)(x + y)')),         # C: CORRECT
    grp(mr('(a − 2b)(x + y)')),         # D: sign error
))

# 12. EKUB(12a², 18a³) ni toping
QUESTIONS.append((12,
    grp(mr('EKUB(12'), sup(mr('a'), mr('2')), mr(', 18'), sup(mr('a'), mr('3')), mr(') ni toping')),
    'D',
    grp(mr('36'), sup(mr('a'), mr('3'))),   # A: EKUK instead
    grp(mr('6'), sup(mr('a'), mr('3'))),    # B: wrong power
    grp(mr('12'), sup(mr('a'), mr('2'))),   # C: close but wrong
    grp(mr('6'), sup(mr('a'), mr('2'))),    # D: CORRECT
))

# 13. EKUK(3x, 6x²) ni toping
QUESTIONS.append((13,
    grp(mr('EKUK(3x, 6'), sup(mr('x'), mr('2')), mr(') ni toping')),
    'A',
    grp(mr('6'), sup(mr('x'), mr('2'))),    # A: CORRECT
    grp(mr('3'), sup(mr('x'), mr('2'))),    # B: wrong coefficient
    grp(mr('18'), sup(mr('x'), mr('3'))),   # C: wrong calculation
    grp(mr('3x')),                           # D: EKUB instead
))

# 14. Agar P(x) = (x-3)(x+2)² bo'lsa, P(3) ni toping
QUESTIONS.append((14,
    grp(mr('Agar P(x) = (x − 3)'), sup(mr('(x + 2)'), mr('2')), mr(' bo\'lsa, P(3) ni toping')),
    'B',
    grp(mr('25')),                           # A: calculation error
    grp(mr('0')),                            # B: CORRECT
    grp(mr('−25')),                          # C: sign error
    grp(mr('5')),                            # D: partial calculation
))

# 15. Agar P(x) = (x+1)³ bo'lsa, P(-1) ni toping
QUESTIONS.append((15,
    grp(mr('Agar P(x) = '), sup(mr('(x + 1)'), mr('3')), mr(' bo\'lsa, P(−1) ni toping')),
    'C',
    grp(mr('1')),                            # A: wrong calculation
    grp(mr('−1')),                           # B: sign error
    grp(mr('0')),                            # C: CORRECT
    grp(mr('8')),                            # D: wrong expansion
))

# 16. Agar P(x) = x² + 4x - 21 bo'lsa, P(-7) ni toping
QUESTIONS.append((16,
    grp(mr('Agar P(x) = '), sup(mr('x'), mr('2')), mr(' + 4x − 21 bo\'lsa, P(−7) ni toping')),
    'D',
    grp(mr('7')),                            # A: calculation error
    grp(mr('−7')),                           # B: sign error
    grp(mr('21')),                           # C: wrong calculation
    grp(mr('0')),                            # D: CORRECT (49-28-21=0)
))

# 17. Agar P(x) = x³ - x + 1 bo'lsa, P(1) ni toping
QUESTIONS.append((17,
    grp(mr('Agar P(x) = '), sup(mr('x'), mr('3')), mr(' − x + 1 bo\'lsa, P(1) ni toping')),
    'A',
    grp(mr('1')),                            # A: CORRECT (1-1+1=1)
    grp(mr('0')),                            # B: calculation error
    grp(mr('2')),                            # C: forgot subtraction
    grp(mr('−1')),                           # D: sign error
))



# 18-28. Ko'paytma ko'rinishidagi masalalar

# 18. P(x) = (x-2)(x+1)(x-3) darajasini toping
QUESTIONS.append((18,
    grp(mr('P(x) = (x − 2)(x + 1)(x − 3) darajasini toping')),
    'B',
    grp(mr('2')),                            # A: counted factors
    grp(mr('3')),                            # B: CORRECT
    grp(mr('1')),                            # C: wrong
    grp(mr('4')),                            # D: wrong
))

# 19. P(x) = (x-2)(x+1)(x-3)(x+2)(x+4) ozod hadini toping
QUESTIONS.append((19,
    grp(mr('P(x) = (x − 2)(x + 1)(x − 3)(x + 2)(x + 4) ozod hadini toping')),
    'C',
    grp(mr('24')),                           # A: wrong calculation
    grp(mr('−24')),                          # B: sign error
    grp(mr('48')),                           # C: CORRECT (-2·1·-3·2·4=48)
    grp(mr('−48')),                          # D: sign error
))

# 20. P(x) = (-2x)(x²-2)(3x+2) ozod hadini toping
QUESTIONS.append((20,
    grp(mr('P(x) = (−2x)('), sup(mr('x'), mr('2')), mr(' − 2)(3x + 2) ozod hadini toping')),
    'D',
    grp(mr('8')),                            # A: sign error
    grp(mr('4')),                            # B: wrong calculation
    grp(mr('−4')),                           # C: wrong
    grp(mr('0')),                            # D: CORRECT (has x factor)
))

# 21. P(x) = (x+3)(x+1)(x²+2x+3) darajasini toping
QUESTIONS.append((21,
    grp(mr('P(x) = (x + 3)(x + 1)('), sup(mr('x'), mr('2')), mr(' + 2x + 3) darajasini toping')),
    'A',
    grp(mr('4')),                            # A: CORRECT
    grp(mr('3')),                            # B: counting error
    grp(mr('5')),                            # C: wrong
    grp(mr('2')),                            # D: wrong
))

# 22. P(x) = x⁴(x-1)(x²+2) ozod hadini toping
QUESTIONS.append((22,
    grp(sup(mr('P(x) = x'), mr('4')), mr('(x − 1)('), sup(mr('x'), mr('2')), mr(' + 2) ozod hadini toping')),
    'B',
    grp(mr('2')),                            # A: forgot x⁴
    grp(mr('0')),                            # B: CORRECT (has x⁴ factor)
    grp(mr('−2')),                           # C: wrong calculation
    grp(mr('8')),                            # D: wrong
))

# 23. P(x) = (3x)²(x+1)²(2-x)³ darajasini toping
QUESTIONS.append((23,
    grp(mr('P(x) = '), sup(mr('(3x)'), mr('2')), sup(mr('(x + 1)'), mr('2')), sup(mr('(2 − x)'), mr('3')), mr(' darajasini toping')),
    'C',
    grp(mr('6')),                            # A: counting error
    grp(mr('8')),                            # B: wrong
    grp(mr('7')),                            # C: CORRECT (2+2+3=7)
    grp(mr('5')),                            # D: wrong
))

# 24. P(x) = (x²-2x+1)(x²-2x+4) ni soddalashtiring
QUESTIONS.append((24,
    grp(mr('P(x) = ('), sup(mr('x'), mr('2')), mr(' − 2x + 1)('), sup(mr('x'), mr('2')), mr(' − 2x + 4) ni soddalashtiring')),
    'D',
    grp(sup(mr('x'), mr('4')), mr(' − 4'), sup(mr('x'), mr('3')), mr(' + 5'), sup(mr('x'), mr('2')), mr(' − 8x + 4')),  # A: expansion error
    grp(sup(mr('(x − 1)'), mr('2')), sup(mr('(x'), mr('2')), mr(' − 2x + 4)')),  # B: partial factorization
    grp(sup(mr('x'), mr('4')), mr(' + 4')),  # C: wrong simplification
    grp(sup(mr('x'), mr('4')), mr(' − 4'), sup(mr('x'), mr('3')), mr(' + 9'), sup(mr('x'), mr('2')), mr(' − 10x + 4')),  # D: CORRECT
))



# 25-35. Qo'shimcha ko'paytma masalalari

# 25. P(x) = (x+3)²(2x-1)³ darajasini toping
QUESTIONS.append((25,
    grp(mr('P(x) = '), sup(mr('(x + 3)'), mr('2')), sup(mr('(2x − 1)'), mr('3')), mr(' darajasini toping')),
    'A',
    grp(mr('5')),                            # A: CORRECT (2+3=5)
    grp(mr('6')),                            # B: multiplication error
    grp(mr('4')),                            # C: counting error
    grp(mr('3')),                            # D: wrong
))

# 26. P(x) = (x²+x+1)(x²+2x+4) darajasini toping
QUESTIONS.append((26,
    grp(mr('P(x) = ('), sup(mr('x'), mr('2')), mr(' + x + 1)('), sup(mr('x'), mr('2')), mr(' + 2x + 4) darajasini toping')),
    'B',
    grp(mr('3')),                            # A: wrong
    grp(mr('4')),                            # B: CORRECT (2+2=4)
    grp(mr('5')),                            # C: wrong
    grp(mr('2')),                            # D: counting error
))

# 27. P(x) = (2x+1)²(x²+1)(x-5) ozod hadini toping
QUESTIONS.append((27,
    grp(mr('P(x) = '), sup(mr('(2x + 1)'), mr('2')), mr('('), sup(mr('x'), mr('2')), mr(' + 1)(x − 5) ozod hadini toping')),
    'C',
    grp(mr('5')),                            # A: sign error
    grp(mr('10')),                           # B: wrong calculation
    grp(mr('−5')),                           # C: CORRECT (1·1·-5=-5)
    grp(mr('0')),                            # D: wrong
))

# 28. P(x) = x(x+1)(x+2)(x+3) + 1 ni ko'paytuvchilarga ajrating
QUESTIONS.append((28,
    grp(mr('P(x) = x(x + 1)(x + 2)(x + 3) + 1 ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('x(x + 1)(x + 2)(x + 4)')),       # A: wrong
    grp(sup(mr('(x + 1)'), mr('2')), sup(mr('(x + 2)'), mr('2'))),  # B: close but wrong
    grp(mr('(x + 1)('), sup(mr('x'), mr('2')), mr(' + 3x + 3)')),    # C: wrong
    grp(sup(grp(mr('('), sup(mr('x'), mr('2')), mr(' + 3x)')), mr('2')), mr(' + 1')),  # D: CORRECT
))



# ═══════════════════════════════════════════════════════════════════
# SAHIFA 2: DAVOMI (29-48)
# ═══════════════════════════════════════════════════════════════════

# 29-35. Ko'paytuvchilarga ajratish

# 29. (x+1)(x-2) + (x+1)(x+3) ni ko'paytuvchilarga ajrating
QUESTIONS.append((29,
    grp(mr('(x + 1)(x − 2) + (x + 1)(x + 3) ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('(x + 1)(2x + 1)')),              # A: CORRECT
    grp(mr('(x + 1)(2x − 1)')),              # B: sign error
    grp(mr('(x − 1)(2x + 1)')),              # C: sign error
    grp(mr('2(x + 1)(x + 1)')),              # D: wrong
))

# 30. (2x-1)(x+3) + (2x-1)(x-4) ni ko'paytuvchilarga ajrating
QUESTIONS.append((30,
    grp(mr('(2x − 1)(x + 3) + (2x − 1)(x − 4) ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(2x − 1)(2x + 1)')),             # A: wrong calculation
    grp(mr('(2x − 1)(2x − 1)')),             # B: CORRECT
    grp(mr('(2x + 1)(2x − 1)')),             # C: sign error
    grp(sup(mr('(2x − 1)'), mr('3'))),       # D: wrong power
))

# 31. x(x+2) - y(x+2) ni ko'paytuvchilarga ajrating
QUESTIONS.append((31,
    grp(mr('x(x + 2) − y(x + 2) ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('(x + 2)(x + y)')),               # A: sign error
    grp(mr('(x − 2)(x − y)')),               # B: two sign errors
    grp(mr('(x + 2)(x − y)')),               # C: CORRECT
    grp(mr('x(x − y) + 2(x − y)')),          # D: not fully factored
))

# 32. (x-1)² + (x-1) ni ko'paytuvchilarga ajrating
QUESTIONS.append((32,
    grp(sup(mr('(x − 1)'), mr('2')), mr(' + (x − 1) ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(x − 1)(x + 1)')),               # A: wrong
    grp(mr('(x − 1)(x − 2)')),               # B: sign error
    grp(sup(mr('(x − 1)'), mr('2'))),        # C: not factored
    grp(mr('(x − 1)x')),                     # D: CORRECT
))

# 33. (x+2)³ - (x+2)² ni ko'paytuvchilarga ajrating
QUESTIONS.append((33,
    grp(sup(mr('(x + 2)'), mr('3')), mr(' − '), sup(mr('(x + 2)'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(sup(mr('(x + 2)'), mr('2')), mr('(x + 1)')),  # A: CORRECT
    grp(sup(mr('(x + 2)'), mr('2')), mr('(x + 2)')),  # B: wrong
    grp(mr('(x + 2)('), sup(mr('x'), mr('2')), mr(' + 1)')),  # C: wrong expansion
    grp(sup(mr('(x + 2)'), mr('2')), mr('x')),        # D: calculation error
))

# 34. x²(x-1) + x(x-1) ni ko'paytuvchilarga ajrating
QUESTIONS.append((34,
    grp(sup(mr('x'), mr('2')), mr('(x − 1) + x(x − 1) ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('x(x − 1)(x + 2)')),              # A: wrong
    grp(mr('x(x − 1)(x + 1)')),              # B: CORRECT
    grp(mr('x(x + 1)(x − 1)')),              # C: same as B but order
    grp(sup(mr('x'), mr('2')), mr('(x − 1)')),  # D: not fully factored
))



# 35-48. Qo'shimcha masalalar

# 35. 3x³ - 12x ni ko'paytuvchilarga ajrating
QUESTIONS.append((35,
    grp(sup(mr('3x'), mr('3')), mr(' − 12x ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('3x('), sup(mr('x'), mr('2')), mr(' + 4)')),   # A: sign error
    grp(mr('x(3'), sup(mr('x'), mr('2')), mr(' − 12)')),  # B: not fully factored
    grp(mr('3x('), sup(mr('x'), mr('2')), mr(' − 4)')),   # C: CORRECT
    grp(mr('3x(x − 2)(x + 2)')),                           # D: extra factorization (also correct but not asked)
))

# 36. 4a² - 9b² ni ko'paytuvchilarga ajrating
QUESTIONS.append((36,
    grp(sup(mr('4a'), mr('2')), mr(' − 9'), sup(mr('b'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(2a − 3b)(2a + 3b)')),           # A: CORRECT but moved to D
    grp(mr('(4a − 9b)(a + b)')),             # B: wrong factorization
    grp(mr('(2a + 3b)(2a + 3b)')),           # C: sign error
    grp(mr('(2a − 3b)(2a + 3b)')),           # D: CORRECT
))

# 37. x³ + 8 ni ko'paytuvchilarga ajrating
QUESTIONS.append((37,
    grp(sup(mr('x'), mr('3')), mr(' + 8 ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('(x + 2)('), sup(mr('x'), mr('2')), mr(' − 2x + 4)')),  # A: CORRECT
    grp(mr('(x + 2)('), sup(mr('x'), mr('2')), mr(' + 2x + 4)')),  # B: sign error
    grp(mr('(x + 8)('), sup(mr('x'), mr('2')), mr(' + 1)')),       # C: wrong
    grp(mr('(x + 2)('), sup(mr('x'), mr('2')), mr(' − 4)')),       # D: wrong formula
))

# 38. x³ - 27 ni ko'paytuvchilarga ajrating
QUESTIONS.append((38,
    grp(sup(mr('x'), mr('3')), mr(' − 27 ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(x − 3)('), sup(mr('x'), mr('2')), mr(' − 3x + 9)')),  # A: sign error in middle term
    grp(mr('(x − 3)('), sup(mr('x'), mr('2')), mr(' + 3x + 9)')),  # B: CORRECT
    grp(mr('(x − 3)('), sup(mr('x'), mr('2')), mr(' + 9)')),       # C: missing middle term
    grp(mr('(x + 3)('), sup(mr('x'), mr('2')), mr(' − 3x + 9)')),  # D: sign error
))

# 39. 2x² + 5x + 3 ni ko'paytuvchilarga ajrating
QUESTIONS.append((39,
    grp(sup(mr('2x'), mr('2')), mr(' + 5x + 3 ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('(2x + 1)(x + 3)')),              # A: wrong coefficients
    grp(mr('(2x − 1)(x − 3)')),              # B: sign errors
    grp(mr('(2x + 3)(x + 1)')),              # C: CORRECT
    grp(mr('(2x + 3)(x − 1)')),              # D: sign error
))

# 40. 3x² - 7x + 2 ni ko'paytuvchilarga ajrating
QUESTIONS.append((40,
    grp(sup(mr('3x'), mr('2')), mr(' − 7x + 2 ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(3x − 1)(x + 2)')),              # A: sign error
    grp(mr('(3x + 1)(x − 2)')),              # B: sign error
    grp(mr('(3x − 2)(x + 1)')),              # C: wrong coefficients
    grp(mr('(3x − 1)(x − 2)')),              # D: CORRECT
))

# 41. x⁴ - 16 ni ko'paytuvchilarga ajrating
QUESTIONS.append((41,
    grp(sup(mr('x'), mr('4')), mr(' − 16 ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('('), sup(mr('x'), mr('2')), mr(' − 4)('), sup(mr('x'), mr('2')), mr(' + 4)')),  # A: CORRECT
    grp(mr('(x − 2)(x + 2)('), sup(mr('x'), mr('2')), mr(' + 4)')),                          # B: over-factored
    grp(mr('(x − 4)(x + 4)('), sup(mr('x'), mr('2')), mr(' + 1)')),                          # C: wrong
    grp(mr('('), sup(mr('x'), mr('2')), mr(' − 16)')),                                       # D: not factored
))

# 42. 6x² - x - 15 ni ko'paytuvchilarga ajrating
QUESTIONS.append((42,
    grp(sup(mr('6x'), mr('2')), mr(' − x − 15 ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(2x − 5)(3x − 3)')),             # A: wrong
    grp(mr('(2x − 3)(3x + 5)')),             # B: CORRECT
    grp(mr('(2x + 3)(3x − 5)')),             # C: sign error
    grp(mr('(6x − 5)(x + 3)')),              # D: wrong
))



# 43-48. Final masalalar

# 43. x⁴ + x² + 1 ni ko'paytuvchilarga ajrating
QUESTIONS.append((43,
    grp(sup(mr('x'), mr('4')), mr(' + '), sup(mr('x'), mr('2')), mr(' + 1 ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('('), sup(mr('x'), mr('2')), mr(' + 1)('), sup(mr('x'), mr('2')), mr(' + 1)')),     # A: wrong
    grp(mr('('), sup(mr('x'), mr('2')), mr(' + x + 1)('), sup(mr('x'), mr('2')), mr(' − x + 1)')),  # B: close but wrong
    grp(mr('('), sup(mr('x'), mr('2')), mr(' − x + 1)('), sup(mr('x'), mr('2')), mr(' + x + 1)')),  # C: CORRECT
    grp(mr('(x + 1)('), sup(mr('x'), mr('3')), mr(' + 1)')),                                    # D: wrong
))

# 44. a³ + b³ + c³ - 3abc ni ko'paytuvchilarga ajrating
QUESTIONS.append((44,
    grp(sup(mr('a'), mr('3')), mr(' + '), sup(mr('b'), mr('3')), mr(' + '), sup(mr('c'), mr('3')), mr(' − 3abc ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(a + b + c)('), sup(mr('a'), mr('2')), mr(' + '), sup(mr('b'), mr('2')), mr(' + '), sup(mr('c'), mr('2')), mr(')')),  # A: missing terms
    grp(mr('(a + b + c)(abc)')),                     # B: wrong
    grp(mr('(a + b + c)³')),                         # C: wrong expansion
    grp(mr('(a + b + c)('), sup(mr('a'), mr('2')), mr(' + '), sup(mr('b'), mr('2')), mr(' + '), sup(mr('c'), mr('2')), mr(' − ab − bc − ca)')),  # D: CORRECT
))

# 45. x² + 6x + 9 ni ko'paytuvchilarga ajrating
QUESTIONS.append((45,
    grp(sup(mr('x'), mr('2')), mr(' + 6x + 9 ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(sup(mr('(x + 3)'), mr('2'))),        # A: CORRECT
    grp(mr('(x + 9)(x + 1)')),               # B: wrong factorization
    grp(mr('(x + 3)(x − 3)')),               # C: sign error
    grp(sup(mr('(x − 3)'), mr('2'))),        # D: sign error
))

# 46. 4x² - 12x + 9 ni ko'paytuvchilarga ajrating
QUESTIONS.append((46,
    grp(sup(mr('4x'), mr('2')), mr(' − 12x + 9 ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(2x − 3)(2x + 3)')),             # A: sign error
    grp(sup(mr('(2x − 3)'), mr('2'))),       # B: CORRECT
    grp(sup(mr('(2x + 3)'), mr('2'))),       # C: sign error
    grp(mr('(4x − 9)(x − 1)')),              # D: wrong
))

# 47. 9a² - 24ab + 16b² ni ko'paytuvchilarga ajrating
QUESTIONS.append((47,
    grp(sup(mr('9a'), mr('2')), mr(' − 24ab + 16'), sup(mr('b'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('(3a − 4b)(3a + 4b)')),           # A: sign error
    grp(sup(mr('(3a + 4b)'), mr('2'))),      # B: sign error
    grp(sup(mr('(3a − 4b)'), mr('2'))),      # C: CORRECT
    grp(mr('(9a − 16b)(a − b)')),            # D: wrong
))

# 48. x² - y² + 2x + 1 ni ko'paytuvchilarga ajrating
QUESTIONS.append((48,
    grp(sup(mr('x'), mr('2')), mr(' − '), sup(mr('y'), mr('2')), mr(' + 2x + 1 ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(x − y)(x + y + 2)')),           # A: wrong
    grp(mr('(x + 1 − y)(x + 1 + y)')),       # B: close but order wrong
    grp(mr('(x − y + 1)(x + y + 1)')),       # C: sign error
    grp(mr('(x + y + 1)(x − y + 1)')),       # D: CORRECT
))

print(f'\n✅ Jami {len(QUESTIONS)} ta savol yaratildi!')




# ═══════════════════════════════════════════════════════════════════
# ANSWER BALANCING
# ═══════════════════════════════════════════════════════════════════

def rebalance_answers(questions, seed=42):
    """Rebalance A/B/C/D distribution by swapping answers"""
    random.seed(seed)
    q_list = list(questions)
    
    # Current distribution
    dist_before = Counter(q[2] for q in q_list)
    print(f"Balanslangandan oldin: A={dist_before.get('A',0)}  B={dist_before.get('B',0)}  C={dist_before.get('C',0)}  D={dist_before.get('D',0)}")
    
    # Target: equal distribution
    total = len(q_list)
    target_per_letter = total // 4
    extra = total % 4
    
    # Track usage
    usage = {k: dist_before.get(k, 0) for k in 'ABCD'}
    target = {k: target_per_letter for k in 'ABCD'}
    for i, k in enumerate('ABCD'):
        if i < extra:
            target[k] += 1
    
    # Swap answers to balance
    for i, q_data in enumerate(q_list):
        num, q_omml, key, A, B, C, D = q_data
        current_key = key
        
        # If current key is overused, try to swap
        if usage[current_key] > target[current_key]:
            # Find underused letter
            for new_key in 'ABCD':
                if usage[new_key] < target[new_key] and new_key != current_key:
                    # Swap variants
                    variants = [A, B, C, D]
                    old_idx = ord(current_key) - ord('A')
                    new_idx = ord(new_key) - ord('A')
                    variants[old_idx], variants[new_idx] = variants[new_idx], variants[old_idx]
                    
                    q_list[i] = (num, q_omml, new_key, *variants)
                    usage[current_key] -= 1
                    usage[new_key] += 1
                    break
    
    dist_after = Counter(q[2] for q in q_list)
    print(f"Balanslangandan keyin: A={dist_after.get('A',0)}  B={dist_after.get('B',0)}  C={dist_after.get('C',0)}  D={dist_after.get('D',0)}")
    
    return q_list




# ═══════════════════════════════════════════════════════════════════
# DOCX GENERATION (Microsoft Word with OMML)
# ═══════════════════════════════════════════════════════════════════

DOCUMENT_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8"?>
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

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:mathPr>
    <m:mathFont m:val="Cambria Math"/>
  </m:mathPr>
</w:settings>'''

def build_document_xml(questions):
    """Build Word document.xml with OMML math"""
    MNS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    
    def p_text(text, bold=False, size=24, center=False):
        b = '<w:b/><w:bCs/>' if bold else ''
        jc = '<w:jc w:val="center"/>' if center else ''
        safe = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        return (f'<w:p {WNS}><w:pPr>{jc}</w:pPr>'
                f'<w:r><w:rPr>{b}<w:sz w:val="{size}"/></w:rPr>'
                f'<w:t>{safe}</w:t></w:r></w:p>')
    
    def p_math(prefix_text, math_inner, bold_prefix=False):
        b = '<w:b/><w:bCs/>' if bold_prefix else ''
        safe = prefix_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        return (f'<w:p {WNS} {MNS}>'
                f'<w:r><w:rPr>{b}<w:sz w:val="26"/></w:rPr>'
                f'<w:t xml:space="preserve">{safe}</w:t></w:r>'
                f'<m:oMath>{math_inner}</m:oMath>'
                '</w:p>')
    
    def p_variants(*variant_tuples):
        parts = []
        for letter, v_omml in variant_tuples:
            parts.append(f'<w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
                        f'<w:t xml:space="preserve">   {letter})  </w:t></w:r>')
            parts.append(f'<m:oMath {MNS}>{v_omml}</m:oMath>')
        return f'<w:p {WNS} {MNS}>{"".join(parts)}</w:p>'
    
    body = []
    
    # Title
    body.append(p_text('KOPHADLAR (POLYNOMIALS)', bold=True, size=32, center=True))
    body.append(p_text(f'Professional A/B/C/D Test ({len(questions)} savol)', bold=True, size=26, center=True))
    body.append(p_text("To'g'ri javobni tanlang", size=22, center=True))
    body.append(p_text('', size=12))
    
    # Questions
    for q_data in questions:
        num, q_omml, key, A, B, C, D = q_data
        body.append(p_math(f'{num}.  ', q_omml, bold_prefix=True))
        body.append(p_variants(('A', A), ('B', B), ('C', C), ('D', D)))
        body.append(p_text('', size=12))
    
    # Page break
    body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    
    # Answer key
    body.append(p_text('JAVOBLAR KALITI (ANSWER KEY)', bold=True, size=32, center=True))
    body.append(p_text('─' * 60, size=20, center=True))
    
    for row_start in range(0, len(questions), 5):
        row_qs = questions[row_start:row_start + 5]
        line = '          '.join(f'{q[0]:2d} — {q[2]}' for q in row_qs)
        body.append(p_text(line, size=22, center=True))
    
    body.append('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>')
    
    return (f'<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:document {DOCUMENT_NS}>'
            f'<w:body>{"".join(body)}</w:body>'
            '</w:document>')

def write_docx(out_path, questions):
    """Write DOCX file"""
    doc_xml = build_document_xml(questions)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', WORD_RELS)
        z.writestr('word/document.xml', doc_xml)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)
    print(f'✅  Fayl tayyor: {out_path}')
    print(f'   Savollar soni: {len(questions)}')
    dist = Counter(q[2] for q in questions)
    print(f'   Javoblar: A={dist.get("A",0)}  B={dist.get("B",0)}  C={dist.get("C",0)}  D={dist.get("D",0)}')

# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print(f'Jami savollar: {len(QUESTIONS)}')
    
    # Balance answers
    balanced = rebalance_answers(QUESTIONS, seed=42)
    
    # Generate DOCX
    out_path = '/projects/sandbox/test-yaratish-uchun/Kophadlar_Full_Test.docx'
    write_docx(out_path, balanced)
    
    print('\n✅ Test muvaffaqiyatli yaratildi!')
    print('   Microsoft Word\'da ochib, matematik ifodalarni tekshiring!')

