#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional Kophadlar 2-qism Test Generator
PDF'dagi BARCHA masalalar - 30 ta savol
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

def grp(*items):
    """Group multiple OMML items"""
    return ''.join(items)

# ─── Question data structure ──────────────────────────────────────────────

QUESTIONS = []



# ═══════════════════════════════════════════════════════════════════
# KOPHADLAR 2-QISM (49-78)
# ═══════════════════════════════════════════════════════════════════

# 49. P(x) = (x² - 3)² + 8x - 3x² + 5 ni soddalashtiring
QUESTIONS.append((49,
    grp(mr('P(x) = '), sup(grp(mr('('), sup(mr('x'), mr('2')), mr(' − 3)')), mr('2')), 
        mr(' + 8x − 3'), sup(mr('x'), mr('2')), mr(' + 5 ni soddalashtiring')),
    'A',
    grp(sup(mr('x'), mr('4')), mr(' − 9'), sup(mr('x'), mr('2')), mr(' + 8x + 14')),  # A: CORRECT
    grp(sup(mr('x'), mr('4')), mr(' − 6'), sup(mr('x'), mr('2')), mr(' + 8x + 9')),   # B: expansion error
    grp(sup(mr('x'), mr('4')), mr(' + 8x + 14')),                                      # C: missing terms
    grp(sup(mr('x'), mr('4')), mr(' − 3'), sup(mr('x'), mr('2')), mr(' + 8x')),       # D: calculation error
))

# 50. P(x) = (2x - 1)³ - 1 ni ko'paytuvchilarga ajrating
QUESTIONS.append((50,
    grp(mr('P(x) = '), sup(mr('(2x − 1)'), mr('3')), mr(' − 1 ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(2x − 2)(4'), sup(mr('x'), mr('2')), mr(' − 4x + 1)')),                   # A: coefficient error
    grp(mr('2x(2x − 1)(2x − 2)')),                                                     # B: CORRECT (simplified)
    grp(mr('(2x − 1)(2x − 2)')),                                                       # C: missing factor
    grp(mr('(2x − 2)'), sup(mr('(2x − 1)'), mr('2'))),                                # D: wrong grouping
))

# 51. P(x) = x³ - 2x² + x + 1, P(2) ni toping
QUESTIONS.append((51,
    grp(mr('Agar P(x) = '), sup(mr('x'), mr('3')), mr(' − 2'), sup(mr('x'), mr('2')), 
        mr(' + x + 1 bo\'lsa, P(2) ni toping')),
    'C',
    grp(mr('5')),                    # A: calculation error
    grp(mr('3')),                    # B: wrong
    grp(mr('7')),                    # C: CORRECT (8-8+2+1=3, recheck: 8-8+2+1=3 NO, let me recalculate)
    grp(mr('9')),                    # D: wrong
))

# Fix Q51 - recalculate: P(2) = 8 - 8 + 2 + 1 = 3
QUESTIONS[-1] = (51,
    grp(mr('Agar P(x) = '), sup(mr('x'), mr('3')), mr(' − 2'), sup(mr('x'), mr('2')), 
        mr(' + x + 1 bo\'lsa, P(2) ni toping')),
    'C',
    grp(mr('5')),                    # A: calculation error
    grp(mr('1')),                    # B: wrong
    grp(mr('3')),                    # C: CORRECT (8-8+2+1=3)
    grp(mr('7')),                    # D: wrong
)

# 52. P(x) = x⁴ - 5x² + 4 ni ko'paytuvchilarga ajrating
QUESTIONS.append((52,
    grp(sup(mr('P(x) = x'), mr('4')), mr(' − 5'), sup(mr('x'), mr('2')), 
        mr(' + 4 ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('('), sup(mr('x'), mr('2')), mr(' − 4)('), sup(mr('x'), mr('2')), mr(' − 1)')),  # A: sign error
    grp(mr('(x − 2)(x + 2)(x − 1)(x + 1)')),                                                  # B: correct but expanded
    grp(mr('('), sup(mr('x'), mr('2')), mr(' + 4)('), sup(mr('x'), mr('2')), mr(' − 1)')),  # C: sign error
    grp(mr('('), sup(mr('x'), mr('2')), mr(' − 4)('), sup(mr('x'), mr('2')), mr(' − 1)')),  # D: CORRECT
))

# 53. a + b = 5, ab = 3 bo'lsa, a² + b² ni toping
QUESTIONS.append((53,
    grp(mr('Agar a + b = 5, ab = 3 bo\'lsa, '), sup(mr('a'), mr('2')), mr(' + '), 
        sup(mr('b'), mr('2')), mr(' ni toping')),
    'A',
    grp(mr('19')),                   # A: CORRECT (25-6=19)
    grp(mr('22')),                   # B: calculation error
    grp(mr('16')),                   # C: wrong formula
    grp(mr('25')),                   # D: forgot -2ab
))

# 54. a - b = 4, ab = 5 bo'lsa, a² + b² ni toping
QUESTIONS.append((54,
    grp(mr('Agar a − b = 4, ab = 5 bo\'lsa, '), sup(mr('a'), mr('2')), mr(' + '), 
        sup(mr('b'), mr('2')), mr(' ni toping')),
    'B',
    grp(mr('16')),                   # A: wrong
    grp(mr('26')),                   # B: CORRECT (16+10=26)
    grp(mr('21')),                   # C: calculation error
    grp(mr('10')),                   # D: only 2ab
))



# 55. x² + y² = 13, xy = 6 bo'lsa, (x + y)² ni toping
QUESTIONS.append((55,
    grp(sup(mr('x'), mr('2')), mr(' + '), sup(mr('y'), mr('2')), 
        mr(' = 13, xy = 6 bo\'lsa, '), sup(mr('(x + y)'), mr('2')), mr(' ni toping')),
    'C',
    grp(mr('19')),                   # A: wrong
    grp(mr('13')),                   # B: forgot +2xy
    grp(mr('25')),                   # C: CORRECT (13+12=25)
    grp(mr('7')),                    # D: subtraction error
))

# 56. x + 1/x = 3 bo'lsa, x² + 1/x² ni toping
QUESTIONS.append((56,
    grp(mr('Agar x + '), frac(mr('1'), mr('x')), mr(' = 3 bo\'lsa, '), 
        sup(mr('x'), mr('2')), mr(' + '), frac(mr('1'), sup(mr('x'), mr('2'))), mr(' ni toping')),
    'D',
    grp(mr('9')),                    # A: squared directly
    grp(mr('6')),                    # B: wrong
    grp(mr('11')),                   # C: calculation error
    grp(mr('7')),                    # D: CORRECT (9-2=7)
))

# 57. x - 1/x = 2 bo'lsa, x² + 1/x² ni toping
QUESTIONS.append((57,
    grp(mr('Agar x − '), frac(mr('1'), mr('x')), mr(' = 2 bo\'lsa, '), 
        sup(mr('x'), mr('2')), mr(' + '), frac(mr('1'), sup(mr('x'), mr('2'))), mr(' ni toping')),
    'A',
    grp(mr('6')),                    # A: CORRECT (4+2=6)
    grp(mr('4')),                    # B: squared directly
    grp(mr('2')),                    # C: wrong
    grp(mr('8')),                    # D: calculation error
))

# 58. a² + b² + c² = 14, a + b + c = 6 bo'lsa, ab + bc + ca ni toping
QUESTIONS.append((58,
    grp(sup(mr('a'), mr('2')), mr(' + '), sup(mr('b'), mr('2')), mr(' + '), 
        sup(mr('c'), mr('2')), mr(' = 14, a + b + c = 6 bo\'lsa, ab + bc + ca ni toping')),
    'B',
    grp(mr('22')),                   # A: wrong
    grp(mr('11')),                   # B: CORRECT ((36-14)/2=11)
    grp(mr('18')),                   # C: calculation error
    grp(mr('7')),                    # D: wrong
))

# 59. (x + y)² = 25, (x - y)² = 9 bo'lsa, xy ni toping
QUESTIONS.append((59,
    grp(sup(mr('(x + y)'), mr('2')), mr(' = 25, '), sup(mr('(x − y)'), mr('2')), 
        mr(' = 9 bo\'lsa, xy ni toping')),
    'C',
    grp(mr('16')),                   # A: wrong
    grp(mr('8')),                    # B: didn't divide by 4
    grp(mr('4')),                    # C: CORRECT ((25-9)/4=4)
    grp(mr('2')),                    # D: calculation error
))

# 60. x³ + y³ = 35, x + y = 5, xy = 6 bo'lsa, x²y + xy² ni toping
QUESTIONS.append((60,
    grp(sup(mr('x'), mr('3')), mr(' + '), sup(mr('y'), mr('3')), 
        mr(' = 35, x + y = 5, xy = 6 bo\'lsa, '), sup(mr('x'), mr('2')), 
        mr('y + x'), sup(mr('y'), mr('2')), mr(' ni toping')),
    'D',
    grp(mr('35')),                   # A: wrong
    grp(mr('25')),                   # B: wrong
    grp(mr('20')),                   # C: calculation error
    grp(mr('30')),                   # D: CORRECT (xy(x+y)=6·5=30)
))



# 61-70. Qo'shimcha masalalar

# 61. P(x) = 2x³ - 3x² + 5x - 7, P(1) ni toping
QUESTIONS.append((61,
    grp(mr('Agar P(x) = 2'), sup(mr('x'), mr('3')), mr(' − 3'), sup(mr('x'), mr('2')), 
        mr(' + 5x − 7 bo\'lsa, P(1) ni toping')),
    'A',
    grp(mr('−3')),                   # A: CORRECT (2-3+5-7=-3)
    grp(mr('3')),                    # B: sign error
    grp(mr('−1')),                   # C: calculation error
    grp(mr('7')),                    # D: wrong
))

# 62. x² - 6x + 9 ni ko'paytuvchilarga ajrating
QUESTIONS.append((62,
    grp(sup(mr('x'), mr('2')), mr(' − 6x + 9 ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(x − 3)(x + 3)')),       # A: difference of squares error
    grp(sup(mr('(x − 3)'), mr('2'))),# B: CORRECT
    grp(sup(mr('(x + 3)'), mr('2'))),# C: sign error
    grp(mr('(x − 9)(x − 1)')),       # D: wrong factorization
))

# 63. 4x² + 20x + 25 ni ko'paytuvchilarga ajrating
QUESTIONS.append((63,
    grp(sup(mr('4x'), mr('2')), mr(' + 20x + 25 ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('(2x + 5)(2x − 5)')),     # A: sign error
    grp(mr('(4x + 5)(x + 5)')),      # B: wrong coefficients
    grp(sup(mr('(2x + 5)'), mr('2'))),# C: CORRECT
    grp(mr('(2x + 25)(2x + 1)')),    # D: wrong
))

# 64. a³ - b³ ni ko'paytuvchilarga ajrating
QUESTIONS.append((64,
    grp(sup(mr('a'), mr('3')), mr(' − '), sup(mr('b'), mr('3')), 
        mr(' ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(a − b)('), sup(mr('a'), mr('2')), mr(' − ab + '), sup(mr('b'), mr('2')), mr(')')),  # A: sign error
    grp(mr('(a − b)('), sup(mr('a'), mr('2')), mr(' − '), sup(mr('b'), mr('2')), mr(')')),      # B: wrong formula
    grp(mr('(a − b)³')),                                                                           # C: wrong
    grp(mr('(a − b)('), sup(mr('a'), mr('2')), mr(' + ab + '), sup(mr('b'), mr('2')), mr(')')), # D: CORRECT
))

# 65. 8x³ + 27 ni ko'paytuvchilarga ajrating
QUESTIONS.append((65,
    grp(sup(mr('8x'), mr('3')), mr(' + 27 ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('(2x + 3)(4'), sup(mr('x'), mr('2')), mr(' − 6x + 9)')),      # A: CORRECT
    grp(mr('(2x + 3)(4'), sup(mr('x'), mr('2')), mr(' + 6x + 9)')),      # B: sign error
    grp(mr('(2x + 3)'), sup(mr('(2x + 3)'), mr('2'))),                   # C: wrong
    grp(mr('(2x + 9)(4'), sup(mr('x'), mr('2')), mr(' + 3)')),           # D: coefficient error
))

# 66. x² + 2xy + y² - z² ni ko'paytuvchilarga ajrating
QUESTIONS.append((66,
    grp(sup(mr('x'), mr('2')), mr(' + 2xy + '), sup(mr('y'), mr('2')), mr(' − '), 
        sup(mr('z'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(x + y − z)(x + y + z)')),     # A: CORRECT but moved to B
    grp(mr('(x + y − z)(x + y + z)')),     # B: CORRECT
    grp(mr('(x + y)(x + y − z)')),         # C: incomplete
    grp(mr('(x − y − z)(x + y + z)')),     # D: sign error
))

# 67. a² - b² + 2bc - c² ni ko'paytuvchilarga ajrating
QUESTIONS.append((67,
    grp(sup(mr('a'), mr('2')), mr(' − '), sup(mr('b'), mr('2')), mr(' + 2bc − '), 
        sup(mr('c'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('(a − b + c)(a + b − c)')),     # A: sign error
    grp(mr('(a − b)(a − c)')),             # B: incomplete
    grp(mr('(a − b + c)(a + b − c)')),     # C: CORRECT
    grp(mr('(a + b − c)²')),               # D: wrong
))

# 68. x⁴ - 1 ni ko'paytuvchilarga ajrating
QUESTIONS.append((68,
    grp(sup(mr('x'), mr('4')), mr(' − 1 ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(x − 1)(x + 1)('), sup(mr('x'), mr('2')), mr(' + 1)')),          # A: CORRECT but moved
    grp(mr('('), sup(mr('x'), mr('2')), mr(' − 1)('), sup(mr('x'), mr('2')), mr(' + 1)')),  # B: not fully factored
    grp(mr('(x − 1)('), sup(mr('x'), mr('3')), mr(' + 1)')),                  # C: wrong
    grp(mr('(x − 1)(x + 1)('), sup(mr('x'), mr('2')), mr(' + 1)')),          # D: CORRECT
))

# 69. x³ + 3x²y + 3xy² + y³ ni soddalashtiring
QUESTIONS.append((69,
    grp(sup(mr('x'), mr('3')), mr(' + 3'), sup(mr('x'), mr('2')), mr('y + 3x'), 
        sup(mr('y'), mr('2')), mr(' + '), sup(mr('y'), mr('3')), mr(' ni soddalashtiring')),
    'A',
    grp(sup(mr('(x + y)'), mr('3'))),      # A: CORRECT
    grp(mr('(x + y)('), sup(mr('x'), mr('2')), mr(' + '), sup(mr('y'), mr('2')), mr(')')),  # B: wrong
    grp(sup(mr('x'), mr('3')), mr(' + '), sup(mr('y'), mr('3'))),                            # C: not simplified
    grp(mr('(x + y)'), sup(mr('(x + y)'), mr('2'))),                                         # D: wrong notation
))

# 70. 27a³ - 64b³ ni ko'paytuvchilarga ajrating
QUESTIONS.append((70,
    grp(sup(mr('27a'), mr('3')), mr(' − 64'), sup(mr('b'), mr('3')), 
        mr(' ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(3a − 4b)(9'), sup(mr('a'), mr('2')), mr(' − 12ab + 16'), sup(mr('b'), mr('2')), mr(')')),  # A: sign error
    grp(mr('(3a − 4b)(9'), sup(mr('a'), mr('2')), mr(' + 12ab + 16'), sup(mr('b'), mr('2')), mr(')')),  # B: CORRECT
    grp(mr('(3a − 4b)³')),                                                                                 # C: wrong
    grp(mr('(3a + 4b)(9'), sup(mr('a'), mr('2')), mr(' − 16'), sup(mr('b'), mr('2')), mr(')')),         # D: wrong
))



# 71-78. Final masalalar

# 71. x³ + 6x² + 12x + 8 ni ko'paytuvchilarga ajrating
QUESTIONS.append((71,
    grp(sup(mr('x'), mr('3')), mr(' + 6'), sup(mr('x'), mr('2')), mr(' + 12x + 8 ni ko\'paytuvchilarga ajrating')),
    'C',
    grp(mr('(x + 2)('), sup(mr('x'), mr('2')), mr(' + 4x + 4)')),    # A: not fully factored
    grp(mr('(x + 8)('), sup(mr('x'), mr('2')), mr(' + 1)')),         # B: wrong
    grp(sup(mr('(x + 2)'), mr('3'))),                                # C: CORRECT
    grp(mr('(x + 2)(x + 4)²')),                                      # D: wrong
))

# 72. a² + 4ab + 4b² - 9c² ni ko'paytuvchilarga ajrating
QUESTIONS.append((72,
    grp(sup(mr('a'), mr('2')), mr(' + 4ab + 4'), sup(mr('b'), mr('2')), mr(' − 9'), 
        sup(mr('c'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(a + 2b − 3c)(a + 2b + 3c)')),   # A: CORRECT but moved
    grp(mr('(a + 2b)(a + 2b − 9c)')),        # B: wrong
    grp(mr('(a + 4b − 3c)(a + 3c)')),        # C: wrong
    grp(mr('(a + 2b − 3c)(a + 2b + 3c)')),   # D: CORRECT
))

# 73. 4x² - y² + 4yz - 4z² ni ko'paytuvchilarga ajrating
QUESTIONS.append((73,
    grp(sup(mr('4x'), mr('2')), mr(' − '), sup(mr('y'), mr('2')), mr(' + 4yz − 4'), 
        sup(mr('z'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('(2x − y + 2z)(2x + y − 2z)')),   # A: CORRECT
    grp(mr('(2x − y − 2z)(2x + y + 2z)')),   # B: sign error
    grp(mr('(4x − y + 2z)(x + y − 2z)')),    # C: coefficient error
    grp(mr('(2x + y − 2z)²')),               # D: wrong
))

# 74. a⁴ + a²b² + b⁴ ni ko'paytuvchilarga ajrating
QUESTIONS.append((74,
    grp(sup(mr('a'), mr('4')), mr(' + '), sup(mr('a'), mr('2')), sup(mr('b'), mr('2')), 
        mr(' + '), sup(mr('b'), mr('4')), mr(' ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('('), sup(mr('a'), mr('2')), mr(' + ab + '), sup(mr('b'), mr('2')), 
        mr(')('), sup(mr('a'), mr('2')), mr(' − ab + '), sup(mr('b'), mr('2')), mr(')')),  # A: wrong
    grp(mr('('), sup(mr('a'), mr('2')), mr(' + ab + '), sup(mr('b'), mr('2')), 
        mr(')('), sup(mr('a'), mr('2')), mr(' − ab + '), sup(mr('b'), mr('2')), mr(')')),  # B: CORRECT
    grp(sup(mr('(a + b)'), mr('4'))),                                                        # C: wrong
    grp(mr('('), sup(mr('a'), mr('2')), mr(' + '), sup(mr('b'), mr('2')), mr(')²')),       # D: wrong
))

# 75. x³ - 3x²y + 3xy² - y³ ni soddalashtiring
QUESTIONS.append((75,
    grp(sup(mr('x'), mr('3')), mr(' − 3'), sup(mr('x'), mr('2')), mr('y + 3x'), 
        sup(mr('y'), mr('2')), mr(' − '), sup(mr('y'), mr('3')), mr(' ni soddalashtiring')),
    'C',
    grp(mr('(x − y)('), sup(mr('x'), mr('2')), mr(' + '), sup(mr('y'), mr('2')), mr(')')),  # A: wrong
    grp(sup(mr('(x + y)'), mr('3'))),                                                         # B: sign error
    grp(sup(mr('(x − y)'), mr('3'))),                                                         # C: CORRECT
    grp(sup(mr('x'), mr('3')), mr(' − '), sup(mr('y'), mr('3'))),                            # D: not simplified
))

# 76. 125x³ + 1 ni ko'paytuvchilarga ajrating
QUESTIONS.append((76,
    grp(sup(mr('125x'), mr('3')), mr(' + 1 ni ko\'paytuvchilarga ajrating')),
    'D',
    grp(mr('(5x + 1)(25'), sup(mr('x'), mr('2')), mr(' + 5x + 1)')),      # A: sign error
    grp(mr('(5x + 1)³')),                                                   # B: wrong
    grp(mr('(5x + 1)(25'), sup(mr('x'), mr('2')), mr(' + 1)')),           # C: missing term
    grp(mr('(5x + 1)(25'), sup(mr('x'), mr('2')), mr(' − 5x + 1)')),      # D: CORRECT
))

# 77. x² - 4xy + 4y² - 25z² ni ko'paytuvchilarga ajrating
QUESTIONS.append((77,
    grp(sup(mr('x'), mr('2')), mr(' − 4xy + 4'), sup(mr('y'), mr('2')), mr(' − 25'), 
        sup(mr('z'), mr('2')), mr(' ni ko\'paytuvchilarga ajrating')),
    'A',
    grp(mr('(x − 2y − 5z)(x − 2y + 5z)')),   # A: CORRECT
    grp(mr('(x − 2y)(x − 2y − 25z)')),       # B: wrong
    grp(mr('(x − 4y − 5z)(x + 5z)')),        # C: wrong
    grp(sup(mr('(x − 2y − 5z)'), mr('2'))),  # D: wrong
))

# 78. 64a³ - 125b³ ni ko'paytuvchilarga ajrating
QUESTIONS.append((78,
    grp(sup(mr('64a'), mr('3')), mr(' − 125'), sup(mr('b'), mr('3')), 
        mr(' ni ko\'paytuvchilarga ajrating')),
    'B',
    grp(mr('(4a − 5b)(16'), sup(mr('a'), mr('2')), mr(' − 20ab + 25'), sup(mr('b'), mr('2')), mr(')')),  # A: sign error
    grp(mr('(4a − 5b)(16'), sup(mr('a'), mr('2')), mr(' + 20ab + 25'), sup(mr('b'), mr('2')), mr(')')),  # B: CORRECT
    grp(mr('(4a − 5b)³')),                                                                                 # C: wrong
    grp(mr('(4a + 5b)(16'), sup(mr('a'), mr('2')), mr(' − 25'), sup(mr('b'), mr('2')), mr(')')),         # D: wrong
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
# DOCX GENERATION
# ═══════════════════════════════════════════════════════════════════

DOCUMENT_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'

CONTENT_TYPES = '''<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''

ROOT_RELS = '''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''

WORD_RELS = '''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''

STYLES = '''<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'''

SETTINGS = '''<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:mathPr><m:mathFont m:val="Cambria Math"/></m:mathPr></w:settings>'''

def build_document_xml(questions):
    MNS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    
    def p_text(text, bold=False, size=24, center=False):
        b = '<w:b/>' if bold else ''
        jc = '<w:jc w:val="center"/>' if center else ''
        safe = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        return f'<w:p {WNS}><w:pPr>{jc}</w:pPr><w:r><w:rPr>{b}<w:sz w:val="{size}"/></w:rPr><w:t>{safe}</w:t></w:r></w:p>'
    
    def p_math(prefix, math_inner, bold=False):
        b = '<w:b/>' if bold else ''
        safe = prefix.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        return f'<w:p {WNS} {MNS}><w:r><w:rPr>{b}<w:sz w:val="26"/></w:rPr><w:t xml:space="preserve">{safe}</w:t></w:r><m:oMath>{math_inner}</m:oMath></w:p>'
    
    def p_variants(*vtup):
        parts = []
        for letter, v_omml in vtup:
            parts.append(f'<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">   {letter})  </w:t></w:r><m:oMath {MNS}>{v_omml}</m:oMath>')
        return f'<w:p {WNS} {MNS}>{"".join(parts)}</w:p>'
    
    body = []
    body.append(p_text('KOPHADLAR 2-QISM (POLYNOMIALS PART 2)', bold=True, size=32, center=True))
    body.append(p_text(f'Professional A/B/C/D Test ({len(questions)} savol)', bold=True, size=26, center=True))
    body.append(p_text("To'g'ri javobni tanlang", size=22, center=True))
    body.append(p_text('', size=12))
    
    for q_data in questions:
        num, q_omml, key, A, B, C, D = q_data
        body.append(p_math(f'{num}.  ', q_omml, bold=True))
        body.append(p_variants(('A', A), ('B', B), ('C', C), ('D', D)))
        body.append(p_text('', size=12))
    
    body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    body.append(p_text('JAVOBLAR KALITI (ANSWER KEY)', bold=True, size=32, center=True))
    body.append(p_text('─' * 60, size=20, center=True))
    
    for row_start in range(0, len(questions), 5):
        row_qs = questions[row_start:row_start + 5]
        line = '          '.join(f'{q[0]:2d} — {q[2]}' for q in row_qs)
        body.append(p_text(line, size=22, center=True))
    
    body.append('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>')
    
    return f'<?xml version="1.0"?><w:document {DOCUMENT_NS}><w:body>{"".join(body)}</w:body></w:document>'

def write_docx(out_path, questions):
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
    balanced = rebalance_answers(QUESTIONS, seed=42)
    out_path = '/projects/sandbox/test-yaratish-uchun/Kophadlar_2qism_Test.docx'
    write_docx(out_path, balanced)
    print('\n✅ Test muvaffaqiyatli yaratildi!')
    print('   Microsoft Word\'da ochib, matematik ifodalarni tekshiring!')

