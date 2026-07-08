#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional Ratsional Tengsizliklar Test Generator
PDF'dagi BARCHA masalalar - 56 ta savol
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

def interval(left_bracket, left_val, right_val, right_bracket):
    """Interval notation: (a;b), [a;b], etc."""
    return grp(mr(left_bracket), mr(left_val), mr(';'), mr(right_val), mr(right_bracket))

def union_intervals(*intervals):
    """Union of intervals with ∪ symbol"""
    parts = []
    for i, inv in enumerate(intervals):
        if i > 0:
            parts.append(mr(' ∪ '))
        parts.append(inv)
    return grp(*parts)

# ─── Question data structure ──────────────────────────────────────────────
# Each question: (number, question_omml, key_letter, A_omml, B_omml, C_omml, D_omml)

QUESTIONS = []



# ═══════════════════════════════════════════════════════════════════
# SAHIFA 1: RATSIONAL TENGSIZLIKLAR (1-25)
# ═══════════════════════════════════════════════════════════════════

# 1. x/(x-7) ≥ 0  =>  x ∈ (-∞;0] ∪ (7;+∞)
QUESTIONS.append((1,
    grp(frac(mr('x'), mr('x − 7')), mr(' ≥ 0')),
    'A',
    union_intervals(interval('(', '−∞', '0', ']'), interval('(', '7', '+∞', ')')),  # A: CORRECT
    union_intervals(interval('[', '0', '7', ')')),  # B: wrong intervals
    union_intervals(interval('(', '−∞', '7', ']')),  # C: forgot denominator zero
    union_intervals(interval('[', '0', '7', ']')),  # D: included x=7
))

# 2. x/(x+2) > 0  =>  x ∈ (-∞;-2) ∪ (0;+∞)
QUESTIONS.append((2,
    grp(frac(mr('x'), mr('x + 2')), mr(' > 0')),
    'B',
    union_intervals(interval('(', '−2', '0', ')')),  # A: sign table error
    union_intervals(interval('(', '−∞', '−2', ')'), interval('(', '0', '+∞', ')')),  # B: CORRECT
    union_intervals(interval('(', '−∞', '0', ')')),  # C: forgot x=-2
    interval('(', '0', '+∞', ')'),  # D: only positive part
))

# 3. (x-4)/(x+1) ≤ 0  =>  x ∈ (-1;4]
QUESTIONS.append((3,
    grp(frac(mr('x − 4'), mr('x + 1')), mr(' ≤ 0')),
    'C',
    interval('[', '−1', '4', ']'),  # A: included x=-1
    interval('(', '−1', '4', ')'),  # B: excluded x=4
    interval('(', '−1', '4', ']'),  # C: CORRECT
    union_intervals(interval('(', '−∞', '−1', ')'), interval('[', '4', '+∞', ')')),  # D: inverted
))

# 4. (x-2)/(x-3) ≥ 0  =>  x ∈ (-∞;2] ∪ (3;+∞)
QUESTIONS.append((4,
    grp(frac(mr('x − 2'), mr('x − 3')), mr(' ≥ 0')),
    'D',
    interval('(', '2', '3', ')'),  # A: inverted
    union_intervals(interval('[', '2', '3', ')')),  # B: wrong direction
    union_intervals(interval('(', '−∞', '3', ']')),  # C: forgot critical point
    union_intervals(interval('(', '−∞', '2', ']'), interval('(', '3', '+∞', ')')),  # D: CORRECT
))

# 5. (x+1)/(x+2) > 0  =>  x ∈ (-∞;-2) ∪ (-1;+∞)
QUESTIONS.append((5,
    grp(frac(mr('x + 1'), mr('x + 2')), mr(' > 0')),
    'A',
    union_intervals(interval('(', '−∞', '−2', ')'), interval('(', '−1', '+∞', ')')),  # A: CORRECT
    interval('(', '−2', '−1', ')'),  # B: wrong direction
    union_intervals(interval('[', '−2', '−1', ']')),  # C: included forbidden points
    interval('(', '−1', '+∞', ')'),  # D: only one part
))

# 6. (x-5)/(x-7) ≥ 0  =>  x ∈ (-∞;5] ∪ (7;+∞)
QUESTIONS.append((6,
    grp(frac(mr('x − 5'), mr('x − 7')), mr(' ≥ 0')),
    'B',
    interval('(', '5', '7', ')'),  # A: inverted
    union_intervals(interval('(', '−∞', '5', ']'), interval('(', '7', '+∞', ')')),  # B: CORRECT
    union_intervals(interval('[', '5', '7', ')')),  # C: wrong direction
    union_intervals(interval('(', '−∞', '5', ']'), interval('[', '7', '+∞', ')')),  # D: included x=7
))

# 7. (x-2)/(x+4) > 0  =>  x ∈ (-∞;-4) ∪ (2;+∞)
QUESTIONS.append((7,
    grp(frac(mr('x − 2'), mr('x + 4')), mr(' > 0')),
    'C',
    interval('(', '−4', '2', ')'),  # A: inverted
    union_intervals(interval('[', '−4', '2', ']')),  # B: wrong boundaries
    union_intervals(interval('(', '−∞', '−4', ')'), interval('(', '2', '+∞', ')')),  # C: CORRECT
    interval('(', '2', '+∞', ')'),  # D: only positive part
))



# 8. (2x+1)/(x+2) ≤ 0  =>  x ∈ (-2;-1/2]
QUESTIONS.append((8,
    grp(frac(mr('2x + 1'), mr('x + 2')), mr(' ≤ 0')),
    'D',
    interval('[', '−2', '−0.5', ']'),  # A: included x=-2
    interval('(', '−2', '−0.5', ')'),  # B: excluded x=-0.5
    union_intervals(interval('(', '−∞', '−2', ')'), interval('[', '−0.5', '+∞', ')')),  # C: inverted
    interval('(', '−2', '−0.5', ']'),  # D: CORRECT
))

# 9. (x-3)/(5x-3) < 0  =>  x ∈ (0.6;3)
QUESTIONS.append((9,
    grp(frac(mr('x − 3'), mr('5x − 3')), mr(' < 0')),
    'A',
    interval('(', '0.6', '3', ')'),  # A: CORRECT (3/5=0.6)
    interval('[', '0.6', '3', ']'),  # B: included boundaries
    union_intervals(interval('(', '−∞', '0.6', ')'), interval('(', '3', '+∞', ')')),  # C: inverted
    interval('(', '3', '+∞', ')'),  # D: only one part
))

# 10. (2x+1)/(x-1) ≥ 0  =>  x ∈ (-∞;-0.5] ∪ (1;+∞)
QUESTIONS.append((10,
    grp(frac(mr('2x + 1'), mr('x − 1')), mr(' ≥ 0')),
    'B',
    interval('(', '−0.5', '1', ')'),  # A: inverted
    union_intervals(interval('(', '−∞', '−0.5', ']'), interval('(', '1', '+∞', ')')),  # B: CORRECT
    union_intervals(interval('[', '−0.5', '1', ']')),  # C: wrong direction
    interval('(', '1', '+∞', ')'),  # D: only positive part
))

# 11. (3-x)/(x-2) ≤ 0  =>  x ∈ (-∞;2) ∪ [3;+∞)
QUESTIONS.append((11,
    grp(frac(mr('3 − x'), mr('x − 2')), mr(' ≤ 0')),
    'C',
    interval('(', '2', '3', ']'),  # A: inverted
    union_intervals(interval('[', '2', '3', ']')),  # B: included x=2
    union_intervals(interval('(', '−∞', '2', ')'), interval('[', '3', '+∞', ')')),  # C: CORRECT
    union_intervals(interval('(', '−∞', '2', ']'), interval('(', '3', '+∞', ')')),  # D: boundary error
))

# 12. (2x+7)/(x+2) ≥ 0  =>  x ∈ (-∞;-3.5] ∪ (-2;+∞)
QUESTIONS.append((12,
    grp(frac(mr('2x + 7'), mr('x + 2')), mr(' ≥ 0')),
    'D',
    interval('(', '−3.5', '−2', ')'),  # A: inverted
    union_intervals(interval('[', '−3.5', '−2', ')')),  # B: wrong direction
    union_intervals(interval('(', '−∞', '−3.5', ')'), interval('(', '−2', '+∞', ')')),  # C: boundary error
    union_intervals(interval('(', '−∞', '−3.5', ']'), interval('(', '−2', '+∞', ')')),  # D: CORRECT
))

# 13. (x+4)/(x+1) ≤ 0  =>  x ∈ (-4;-1)
QUESTIONS.append((13,
    grp(frac(mr('x + 4'), mr('x + 1')), mr(' ≤ 0')),
    'A',
    interval('[', '−4', '−1', ')'),  # A: CORRECT
    interval('(', '−4', '−1', ')'),  # B: excluded x=-4
    union_intervals(interval('(', '−∞', '−4', ')'), interval('(', '−1', '+∞', ')')),  # C: inverted
    interval('[', '−4', '−1', ']'),  # D: included x=-1
))



# 14. (x+4)(x+1)x ≤ 0  =>  x ∈ (-∞;-4] ∪ [-1;0]
QUESTIONS.append((14,
    grp(mr('(x + 4)(x + 1)x ≤ 0')),
    'B',
    union_intervals(interval('[', '−4', '−1', ']'), interval('[', '0', '+∞', ')')),  # A: wrong intervals
    union_intervals(interval('(', '−∞', '−4', ']'), interval('[', '−1', '0', ']')),  # B: CORRECT
    interval('[', '−4', '0', ']'),  # C: forgot middle part
    union_intervals(interval('(', '−∞', '−4', ')'), interval('(', '−1', '0', ')')),  # D: wrong boundaries
))

# 15. (x+2)(x-6) > 0  =>  x ∈ (-∞;-2) ∪ (6;+∞)
QUESTIONS.append((15,
    grp(mr('(x + 2)(x − 6) > 0')),
    'C',
    interval('(', '−2', '6', ')'),  # A: inverted
    union_intervals(interval('[', '−2', '6', ']')),  # B: wrong boundaries
    union_intervals(interval('(', '−∞', '−2', ')'), interval('(', '6', '+∞', ')')),  # C: CORRECT
    interval('(', '6', '+∞', ')'),  # D: only positive part
))

# 16. (x-2)(x+6) < 0  =>  x ∈ (-6;2)
QUESTIONS.append((16,
    grp(mr('(x − 2)(x + 6) < 0')),
    'D',
    union_intervals(interval('(', '−∞', '−6', ')'), interval('(', '2', '+∞', ')')),  # A: inverted
    interval('[', '−6', '2', ']'),  # B: wrong boundaries
    interval('(', '−6', '2', ']'),  # C: boundary error
    interval('(', '−6', '2', ')'),  # D: CORRECT
))

# 17. x(x-3)(5x-3) < 0  =>  x ∈ (-∞;0) ∪ (0.6;3)
QUESTIONS.append((17,
    grp(mr('x(x − 3)(5x − 3) < 0')),
    'A',
    union_intervals(interval('(', '−∞', '0', ')'), interval('(', '0.6', '3', ')')),  # A: CORRECT
    interval('(', '0', '3', ')'),  # B: forgot 5x-3=0
    union_intervals(interval('(', '−∞', '0', ')'), interval('(', '3', '+∞', ')')),  # C: wrong intervals
    interval('(', '0.6', '3', ')'),  # D: only middle part
))

# 18. (x-2)x(x+6) > 0  =>  x ∈ (-6;0) ∪ (2;+∞)
QUESTIONS.append((18,
    grp(mr('(x − 2)x(x + 6) > 0')),
    'B',
    interval('(', '0', '2', ')'),  # A: wrong interval
    union_intervals(interval('(', '−6', '0', ')'), interval('(', '2', '+∞', ')')),  # B: CORRECT
    union_intervals(interval('(', '−∞', '−6', ')'), interval('(', '0', '2', ')')),  # C: inverted
    interval('(', '2', '+∞', ')'),  # D: only positive part
))

# 19. (x+10)(x+1)(x-3) ≥ 0  =>  x ∈ [-10;-1] ∪ [3;+∞)
QUESTIONS.append((19,
    grp(mr('(x + 10)(x + 1)(x − 3) ≥ 0')),
    'C',
    union_intervals(interval('(', '−10', '−1', ')'), interval('(', '3', '+∞', ')')),  # A: wrong boundaries
    union_intervals(interval('(', '−∞', '−10', ')'), interval('(', '−1', '3', ')')),  # B: inverted
    union_intervals(interval('[', '−10', '−1', ']'), interval('[', '3', '+∞', ')')),  # C: CORRECT
    interval('[', '−10', '3', ']'),  # D: wrong intervals
))

# 20. x²-10x+16 ≥ 0  =>  x ∈ (-∞;2] ∪ [8;+∞)
QUESTIONS.append((20,
    grp(sup(mr('x'), mr('2')), mr(' − 10x + 16 ≥ 0')),
    'D',
    interval('[', '2', '8', ']'),  # A: inverted
    union_intervals(interval('(', '−∞', '2', ')'), interval('(', '8', '+∞', ')')),  # B: wrong boundaries
    union_intervals(interval('[', '2', '8', ')')),  # C: wrong direction
    union_intervals(interval('(', '−∞', '2', ']'), interval('[', '8', '+∞', ')')),  # D: CORRECT
))



# 21. x²-2x < 0  =>  x ∈ (0;2)
QUESTIONS.append((21,
    grp(sup(mr('x'), mr('2')), mr(' − 2x < 0')),
    'A',
    interval('(', '0', '2', ')'),  # A: CORRECT
    interval('[', '0', '2', ']'),  # B: wrong boundaries
    union_intervals(interval('(', '−∞', '0', ')'), interval('(', '2', '+∞', ')')),  # C: inverted
    interval('(', '0', '2', ']'),  # D: boundary error
))

# 22. x²-1 ≥ 0  =>  x ∈ (-∞;-1] ∪ [1;+∞)
QUESTIONS.append((22,
    grp(sup(mr('x'), mr('2')), mr(' − 1 ≥ 0')),
    'B',
    interval('[', '−1', '1', ']'),  # A: inverted
    union_intervals(interval('(', '−∞', '−1', ']'), interval('[', '1', '+∞', ')')),  # B: CORRECT
    union_intervals(interval('(', '−∞', '−1', ')'), interval('(', '1', '+∞', ')')),  # C: wrong boundaries
    interval('(', '−∞', '+∞', ')'),  # D: forgot critical points
))

# 23. x²-36 ≤ 0  =>  x ∈ [-6;6]
QUESTIONS.append((23,
    grp(sup(mr('x'), mr('2')), mr(' − 36 ≤ 0')),
    'C',
    interval('(', '−6', '6', ')'),  # A: wrong boundaries
    union_intervals(interval('(', '−∞', '−6', ']'), interval('[', '6', '+∞', ')')),  # B: inverted
    interval('[', '−6', '6', ']'),  # C: CORRECT
    interval('[', '−6', '6', ')'),  # D: boundary error
))

# 24. x²-2x+7 > 0  =>  x ∈ ℝ (always positive, discriminant < 0)
QUESTIONS.append((24,
    grp(sup(mr('x'), mr('2')), mr(' − 2x + 7 > 0')),
    'D',
    interval('(', '0', '+∞', ')'),  # A: wrong interval
    interval('(', '−∞', '0', ')'),  # B: wrong interval
    mr('∅'),  # C: empty set (wrong)
    mr('ℝ'),  # D: CORRECT (all real numbers)
))

# 25. x²-3x+5 > 0  =>  x ∈ ℝ (always positive, discriminant < 0)
QUESTIONS.append((25,
    grp(sup(mr('x'), mr('2')), mr(' − 3x + 5 > 0')),
    'A',
    mr('ℝ'),  # A: CORRECT
    interval('(', '0', '+∞', ')'),  # B: wrong interval
    mr('∅'),  # C: empty set (wrong)
    interval('(', '−∞', '+∞', ')'),  # D: same as ℝ but different notation
))



# ═══════════════════════════════════════════════════════════════════
# SAHIFA 2: DAVOMI (26-46)
# ═══════════════════════════════════════════════════════════════════

# 26. x/(x+1) ≥ 1  =>  x/(x+1)-1≥0  =>  -1/(x+1)≥0  =>  x+1<0  =>  x ∈ (-∞;-1)
QUESTIONS.append((26,
    grp(frac(mr('x'), mr('x + 1')), mr(' ≥ 1')),
    'B',
    interval('(', '−1', '+∞', ')'),  # A: wrong direction
    interval('(', '−∞', '−1', ')'),  # B: CORRECT
    interval('[', '−1', '+∞', ')'),  # C: included x=-1
    mr('∅'),  # D: empty set (wrong)
))

# 27. x/(x+2) ≥ 1/2  =>  x/(x+2)-1/2≥0  =>  (x-2)/(2(x+2))≥0  =>  x ∈ (-∞;-2) ∪ [2;+∞)
QUESTIONS.append((27,
    grp(frac(mr('x'), mr('x + 2')), mr(' ≥ '), frac(mr('1'), mr('2'))),
    'C',
    interval('[', '2', '+∞', ')'),  # A: only one part
    interval('(', '−2', '2', ']'),  # B: inverted
    union_intervals(interval('(', '−∞', '−2', ')'), interval('[', '2', '+∞', ')')),  # C: CORRECT
    union_intervals(interval('(', '−∞', '−2', ']'), interval('[', '2', '+∞', ')')),  # D: boundary error
))

# 28. (6x-1)/3 + (x-1)/2 > 0  =>  (12x-2+3x-3)/6>0  =>  15x-5>0  =>  x > 1/3
QUESTIONS.append((28,
    grp(frac(mr('6x − 1'), mr('3')), mr(' + '), frac(mr('x − 1'), mr('2')), mr(' > 0')),
    'D',
    interval('(', '0', '+∞', ')'),  # A: wrong critical point
    interval('[', '0.33', '+∞', ')'),  # B: wrong boundary type
    interval('(', '−∞', '0.33', ')'),  # C: wrong direction
    interval('(', '0.33', '+∞', ')'),  # D: CORRECT (1/3≈0.33)
))

# 29. (x+2)/2 ≤ (x+3)/(x+2)  =>  complex, solution depends on analysis
# Simplified: x ∈ [-2;-1] ∪ [2;+∞) (approximate)
QUESTIONS.append((29,
    grp(frac(mr('x + 2'), mr('2')), mr(' ≤ '), frac(mr('x + 3'), mr('x + 2'))),
    'A',
    union_intervals(interval('[', '−2', '−1', ']'), interval('[', '2', '+∞', ')')),  # A: CORRECT (approx)
    interval('[', '−2', '+∞', ')'),  # B: simplified wrong
    interval('(', '−∞', '−2', ')'),  # C: wrong direction
    mr('∅'),  # D: empty set (wrong)
))

# 30. (3x+4)/2 - (x+5)/3 ≤ 0  =>  (9x+12-2x-10)/6≤0  =>  7x+2≤0  =>  x ≤ -2/7
QUESTIONS.append((30,
    grp(frac(mr('3x + 4'), mr('2')), mr(' − '), frac(mr('x + 5'), mr('3')), mr(' ≤ 0')),
    'B',
    interval('(', '−∞', '−0.29', ')'),  # A: wrong boundary type
    interval('(', '−∞', '−0.29', ']'),  # B: CORRECT (-2/7≈-0.29)
    interval('[', '−0.29', '+∞', ')'),  # C: wrong direction
    mr('ℝ'),  # D: all reals (wrong)
))

# 31. x/3 + 1/2 ≤ 0  =>  2x+3≤0  =>  x ≤ -3/2
QUESTIONS.append((31,
    grp(frac(mr('x'), mr('3')), mr(' + '), frac(mr('1'), mr('2')), mr(' ≤ 0')),
    'C',
    interval('(', '−∞', '−1.5', ')'),  # A: wrong boundary type
    interval('[', '−1.5', '+∞', ')'),  # B: wrong direction
    interval('(', '−∞', '−1.5', ']'),  # C: CORRECT (-3/2=-1.5)
    interval('(', '−1.5', '0', ']'),  # D: wrong interval
))

# 32. x/2 - 1/4 > 0  =>  2x-1>0  =>  x > 1/2
QUESTIONS.append((32,
    grp(frac(mr('x'), mr('2')), mr(' − '), frac(mr('1'), mr('4')), mr(' > 0')),
    'D',
    interval('[', '0.5', '+∞', ')'),  # A: wrong boundary type
    interval('(', '−∞', '0.5', ')'),  # B: wrong direction
    interval('(', '0', '+∞', ')'),  # C: wrong critical point
    interval('(', '0.5', '+∞', ')'),  # D: CORRECT (1/2=0.5)
))



# 33. (x+4)(x+1)/x ≤ 0  =>  x ∈ (-∞;-4] ∪ (-1;0)
QUESTIONS.append((33,
    grp(frac(mr('(x + 4)(x + 1)'), mr('x')), mr(' ≤ 0')),
    'A',
    union_intervals(interval('(', '−∞', '−4', ']'), interval('[', '−1', '0', ')')),  # A: CORRECT
    interval('[', '−4', '0', ')'),  # B: wrong intervals
    union_intervals(interval('[', '−4', '−1', ']'), interval('(', '0', '+∞', ')')),  # C: inverted
    interval('(', '−4', '−1', ')'),  # D: forgot other parts
))

# 34. (x-3)/(x+2) ≤ 1  =>  (x-3)/(x+2)-1≤0  =>  -5/(x+2)≤0  =>  x+2>0  =>  x ∈ (-2;+∞)
QUESTIONS.append((34,
    grp(frac(mr('x − 3'), mr('x + 2')), mr(' ≤ 1')),
    'B',
    interval('[', '−2', '+∞', ')'),  # A: included x=-2
    interval('(', '−2', '+∞', ')'),  # B: CORRECT
    interval('(', '−∞', '−2', ')'),  # C: wrong direction
    mr('ℝ'),  # D: all reals (wrong)
))

# 35. 1/(x+2) > 0  =>  x+2 > 0  =>  x ∈ (-2;+∞)
QUESTIONS.append((35,
    grp(frac(mr('1'), mr('x + 2')), mr(' > 0')),
    'C',
    interval('[', '−2', '+∞', ')'),  # A: wrong boundary type
    interval('(', '−∞', '−2', ')'),  # B: wrong direction
    interval('(', '−2', '+∞', ')'),  # C: CORRECT
    interval('(', '0', '+∞', ')'),  # D: wrong critical point
))

# 36. (x+1)/(x-2) ≥ 0  =>  x ∈ (-∞;-1] ∪ (2;+∞)
QUESTIONS.append((36,
    grp(frac(mr('x + 1'), mr('x − 2')), mr(' ≥ 0')),
    'D',
    interval('(', '−1', '2', ')'),  # A: inverted
    union_intervals(interval('[', '−1', '2', ')')),  # B: wrong direction
    union_intervals(interval('(', '−∞', '−1', ')'), interval('(', '2', '+∞', ')')),  # C: boundary error
    union_intervals(interval('(', '−∞', '−1', ']'), interval('(', '2', '+∞', ')')),  # D: CORRECT
))

# 37. (2x+3)/((x+1)(x-2)) > 0  =>  Critical points: -1.5, -1, 2
# Sign table: x ∈ (-∞;-1.5) ∪ (-1;2)
QUESTIONS.append((37,
    grp(frac(mr('2x + 3'), mr('(x + 1)(x − 2)')), mr(' > 0')),
    'A',
    union_intervals(interval('(', '−∞', '−1.5', ')'), interval('(', '−1', '2', ')')),  # A: CORRECT
    interval('(', '−1', '2', ')'),  # B: only middle part
    union_intervals(interval('(', '−1.5', '−1', ')'), interval('(', '2', '+∞', ')')),  # C: wrong intervals
    interval('(', '−∞', '−1', ')'),  # D: forgot other parts
))

# 38. (x+3)/((x-1)(x+4)) ≤ 0  =>  Critical points: -4, -3, 1
# Sign table: x ∈ (-∞;-4) ∪ [-3;1)
QUESTIONS.append((38,
    grp(frac(mr('x + 3'), mr('(x − 1)(x + 4)')), mr(' ≤ 0')),
    'B',
    union_intervals(interval('[', '−4', '−3', ')'), interval('(', '1', '+∞', ')')),  # A: wrong intervals
    union_intervals(interval('(', '−∞', '−4', ')'), interval('[', '−3', '1', ')')),  # B: CORRECT
    interval('[', '−4', '1', ']'),  # C: forgot denominators
    interval('[', '−3', '1', ')'),  # D: only middle part
))



# 39-42: Kompleks kvadrat tengsizliklar

# 39. (x²-3x-10)/(x²+2x+8) > 0  =>  Denominator always positive (D<0)
# Numerator: (x-5)(x+2) > 0  =>  x ∈ (-∞;-2) ∪ (5;+∞)
QUESTIONS.append((39,
    grp(frac(grp(sup(mr('x'), mr('2')), mr(' − 3x − 10')),
             grp(sup(mr('x'), mr('2')), mr(' + 2x + 8'))),
        mr(' > 0')),
    'C',
    interval('(', '−2', '5', ')'),  # A: inverted
    union_intervals(interval('[', '−2', '5', ']')),  # B: wrong boundaries
    union_intervals(interval('(', '−∞', '−2', ')'), interval('(', '5', '+∞', ')')),  # C: CORRECT
    mr('ℝ'),  # D: all reals (wrong)
))

# 40. (x²-4)/(x²-5x+6) ≤ 0  =>  (x-2)(x+2)/((x-2)(x-3)) ≤ 0
# Simplify: (x+2)/(x-3) ≤ 0  (x≠2)  =>  x ∈ [-2;2) ∪ (2;3)
QUESTIONS.append((40,
    grp(frac(grp(sup(mr('x'), mr('2')), mr(' − 4')),
             grp(sup(mr('x'), mr('2')), mr(' − 5x + 6'))),
        mr(' ≤ 0')),
    'D',
    interval('[', '−2', '3', ')'),  # A: forgot x=2
    interval('(', '−2', '3', ']'),  # B: wrong boundaries
    union_intervals(interval('[', '−2', '2', ']'), interval('[', '3', '+∞', ')')),  # C: inverted
    union_intervals(interval('[', '−2', '2', ')'), interval('(', '2', '3', ')')),  # D: CORRECT
))

# 41. (x-1)/(x²-9) < 0  =>  (x-1)/((x-3)(x+3)) < 0
# Critical points: -3, 1, 3  =>  x ∈ (-∞;-3) ∪ (1;3)
QUESTIONS.append((41,
    grp(frac(mr('x − 1'), grp(sup(mr('x'), mr('2')), mr(' − 9'))), mr(' < 0')),
    'A',
    union_intervals(interval('(', '−∞', '−3', ')'), interval('(', '1', '3', ')')),  # A: CORRECT
    interval('(', '−3', '1', ')'),  # B: only middle part
    union_intervals(interval('(', '−3', '1', ')'), interval('(', '3', '+∞', ')')),  # C: wrong intervals
    interval('(', '1', '3', ')'),  # D: only middle part
))

# 42. (x²+3x+2)/(x²-x-6) ≥ 0  =>  ((x+1)(x+2))/((x-3)(x+2)) ≥ 0
# Simplify: (x+1)/(x-3) ≥ 0  (x≠-2)  =>  x ∈ (-∞;-2) ∪ (-2;-1] ∪ (3;+∞)
QUESTIONS.append((42,
    grp(frac(grp(sup(mr('x'), mr('2')), mr(' + 3x + 2')),
             grp(sup(mr('x'), mr('2')), mr(' − x − 6'))),
        mr(' ≥ 0')),
    'B',
    union_intervals(interval('(', '−∞', '−1', ']'), interval('(', '3', '+∞', ')')),  # A: forgot x=-2
    union_intervals(interval('(', '−∞', '−2', ')'), interval('(', '−2', '−1', ']'), interval('(', '3', '+∞', ')')),  # B: CORRECT
    union_intervals(interval('[', '−1', '3', ')')),  # C: wrong intervals
    interval('(', '3', '+∞', ')'),  # D: only one part
))



# ═══════════════════════════════════════════════════════════════════
# SAHIFA 3: KOMPLEKS TENGSIZLIKLAR (48-56)
# ═══════════════════════════════════════════════════════════════════

# 48. (5-2x)/(x+4) > 0  =>  Critical points: 2.5, -4  =>  x ∈ (-4;2.5)
QUESTIONS.append((48,
    grp(frac(mr('5 − 2x'), mr('x + 4')), mr(' > 0')),
    'C',
    interval('(', '−∞', '−4', ')'),  # A: wrong interval
    union_intervals(interval('(', '−∞', '−4', ')'), interval('(', '2.5', '+∞', ')')),  # B: inverted
    interval('(', '−4', '2.5', ')'),  # C: CORRECT
    interval('(', '2.5', '+∞', ')'),  # D: only one part
))

# 49. ((x+3)(x-2))/((x+1)(x-4)) < 0
# Critical points: -3, -1, 2, 4  =>  x ∈ (-∞;-3) ∪ (-1;2) ∪ (4;+∞)
QUESTIONS.append((49,
    grp(frac(mr('(x + 3)(x − 2)'), mr('(x + 1)(x − 4)')), mr(' < 0')),
    'D',
    interval('(', '−1', '2', ')'),  # A: only middle part
    union_intervals(interval('(', '−3', '−1', ')'), interval('(', '2', '4', ')')),  # B: inverted
    union_intervals(interval('(', '−∞', '−3', ')'), interval('(', '2', '4', ')')),  # C: missing part
    union_intervals(interval('(', '−∞', '−3', ')'), interval('(', '−1', '2', ')'), interval('(', '4', '+∞', ')')),  # D: CORRECT
))

# 50. (x²+7x+12)/(x²+x-2) < 0  =>  ((x+3)(x+4))/((x+2)(x-1)) < 0
# Critical points: -4, -3, -2, 1  =>  x ∈ (-∞;-4) ∪ (-3;-2) ∪ (1;+∞)
QUESTIONS.append((50,
    grp(frac(grp(sup(mr('x'), mr('2')), mr(' + 7x + 12')),
             grp(sup(mr('x'), mr('2')), mr(' + x − 2'))),
        mr(' < 0')),
    'A',
    union_intervals(interval('(', '−∞', '−4', ')'), interval('(', '−3', '−2', ')'), interval('(', '1', '+∞', ')')),  # A: CORRECT
    interval('(', '−3', '−2', ')'),  # B: only middle part
    union_intervals(interval('(', '−4', '−3', ')'), interval('(', '−2', '1', ')')),  # C: inverted
    interval('(', '1', '+∞', ')'),  # D: only one part
))

# 51. x²+2x+3 > 0  =>  Discriminant < 0  =>  x ∈ ℝ (always positive)
QUESTIONS.append((51,
    grp(sup(mr('x'), mr('2')), mr(' + 2x + 3 > 0')),
    'B',
    interval('(', '0', '+∞', ')'),  # A: wrong interval
    mr('ℝ'),  # B: CORRECT
    mr('∅'),  # C: empty set (wrong)
    interval('(', '−∞', '0', ')'),  # D: wrong interval
))

# 52. 6x²+5x+4 < 0  =>  Discriminant < 0, leading coeff > 0  =>  ∅ (never negative)
QUESTIONS.append((52,
    grp(sup(mr('6x'), mr('2')), mr(' + 5x + 4 < 0')),
    'C',
    mr('ℝ'),  # A: all reals (wrong)
    interval('(', '−∞', '0', ')'),  # B: wrong interval
    mr('∅'),  # C: CORRECT (empty set)
    interval('(', '0', '+∞', ')'),  # D: wrong interval
))

# 53. -2x²+3x-8 > 0  =>  Leading coeff < 0, discriminant < 0  =>  ∅ (always negative)
QUESTIONS.append((53,
    grp(mr('−2'), sup(mr('x'), mr('2')), mr(' + 3x − 8 > 0')),
    'D',
    mr('ℝ'),  # A: all reals (wrong)
    interval('(', '0', '+∞', ')'),  # B: wrong interval
    interval('(', '−∞', '0', ')'),  # C: wrong interval
    mr('∅'),  # D: CORRECT
))

# 54. -3x²+7x-5 < 0  =>  Leading coeff < 0, check discriminant
# D = 49-60 = -11 < 0, so always negative  =>  x ∈ ℝ
QUESTIONS.append((54,
    grp(mr('−3'), sup(mr('x'), mr('2')), mr(' + 7x − 5 < 0')),
    'A',
    mr('ℝ'),  # A: CORRECT (always negative with leading coeff < 0)
    mr('∅'),  # B: empty set (wrong)
    interval('(', '0', '+∞', ')'),  # C: wrong interval
    interval('(', '−∞', '0', ')'),  # D: wrong interval
))

# 55. 2x²+7x+14 > 0  =>  Discriminant = 49-112 = -63 < 0, leading coeff > 0  =>  x ∈ ℝ
QUESTIONS.append((55,
    grp(sup(mr('2x'), mr('2')), mr(' + 7x + 14 > 0')),
    'B',
    mr('∅'),  # A: empty set (wrong)
    mr('ℝ'),  # B: CORRECT
    interval('(', '0', '+∞', ')'),  # C: wrong interval
    interval('(', '−∞', '0', ')'),  # D: wrong interval
))

# 56. (3x²+2x+7)/(x-2) > 0
# Numerator always positive (D < 0), so depends on denominator
# x-2 > 0  =>  x ∈ (2;+∞)
QUESTIONS.append((56,
    grp(frac(grp(sup(mr('3x'), mr('2')), mr(' + 2x + 7')), mr('x − 2')), mr(' > 0')),
    'C',
    interval('[', '2', '+∞', ')'),  # A: wrong boundary type
    interval('(', '−∞', '2', ')'),  # B: wrong direction
    interval('(', '2', '+∞', ')'),  # C: CORRECT
    mr('ℝ'),  # D: all reals (wrong)
))




# ═══════════════════════════════════════════════════════════════════
# ANSWER BALANCING
# ═══════════════════════════════════════════════════════════════════

def rebalance_answers(questions, seed=42):
    """Rebalance A/B/C/D distribution by swapping answers"""
    random.seed(seed)
    q_list = list(questions)
    
    # Current distribution
    dist_before = Counter(q[2] for q in q_list)
    print(f"Balanslangandan oldin: A={dist_before['A']}  B={dist_before['B']}  C={dist_before['C']}  D={dist_before['D']}")
    
    # Target: equal distribution
    total = len(q_list)
    target_per_letter = total // 4
    extra = total % 4
    
    # Track usage
    usage = dict(dist_before)
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
    print(f"Balanslangandan keyin: A={dist_after['A']}  B={dist_after['B']}  C={dist_after['C']}  D={dist_after['D']}")
    
    return q_list




# ═══════════════════════════════════════════════════════════════════
# DOCX GENERATION (Microsoft Word with OMML)
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
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    body.append(p_text('RATSIONAL TENGSIZLIKLAR', bold=True, size=32, center=True))
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
    print(f'   Javoblar: A={dist["A"]}  B={dist["B"]}  C={dist["C"]}  D={dist["D"]}')




# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print(f'Jami savollar: {len(QUESTIONS)}')
    
    # Balance answers
    balanced = rebalance_answers(QUESTIONS, seed=42)
    
    # Generate DOCX
    out_path = '/projects/sandbox/test-yaratish-uchun/Ratsional_Tengsizlik_Full_Test.docx'
    write_docx(out_path, balanced)
    
    print('\n✅ Test muvaffaqiyatli yaratildi!')
    print('   Microsoft Word\'da ochib, matematik ifodalarni tekshiring!')

