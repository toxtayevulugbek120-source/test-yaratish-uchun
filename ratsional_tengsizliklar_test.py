#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional matematik test generatori
4. RATSIONAL TENGSIZLIKLAR (Masalalar 1-95)
Natija: Microsoft Word (.docx) OMML equation formatida
"""

import zipfile
import random
from collections import Counter

# ─── Namespaces ───────────────────────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# ─── OMML string helpers ─────────────────────────────────────────────────────

def mr(text):
    """Math run"""
    safe = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return (f'<m:r xmlns:m="{M}">'
            f'<m:rPr><m:rFont m:val="Cambria Math"/><m:sty m:val="p"/></m:rPr>'
            f'<m:t xml:space="preserve">{safe}</m:t></m:r>')

def mf(num_str, den_str):
    """Stacked fraction"""
    safe_n = num_str.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    safe_d = den_str.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return (f'<m:f xmlns:m="{M}">'
            f'<m:num>{mr(safe_n)}</m:num>'
            f'<m:den>{mr(safe_d)}</m:den>'
            f'</m:f>')

def msup(base, sup):
    """Superscript"""
    return (f'<m:sSup xmlns:m="{M}">'
            f'<m:e>{mr(base)}</m:e>'
            f'<m:sup>{mr(sup)}</m:sup>'
            f'</m:sSup>')

def mg(*parts):
    """Group multiple OMML strings"""
    return "".join(parts)

def omml_inline(inner):
    """Inline math"""
    return f'<m:oMath xmlns:m="{M}">{inner}</m:oMath>'



# ─── Questions data ───────────────────────────────────────────────────────────
# RATSIONAL TENGSIZLIKLAR 1-14

QUESTIONS = [
    # Problem 1: x/(x-1) ≤ 0
    {"num":1,
     "q": mg(mf("x","x−1"), mr(" ≤ 0")),
     "A": mr("(−∞ ; 0] ∪ (1 ; +∞)"),
     "B": mr("[0 ; 1)"),
     "C": mr("(−∞ ; 0) ∪ (1 ; +∞)"),
     "D": mr("[0 ; 1]"),
     "ans":"B"},
    
    # Problem 2: (x+2)/(x-1) ≥ 0
    {"num":2,
     "q": mg(mf("x+2","x−1"), mr(" ≥ 0")),
     "A": mr("(−∞ ; −2] ∪ (1 ; +∞)"),
     "B": mr("[−2 ; 1)"),
     "C": mr("(−∞ ; −2) ∪ (1 ; +∞)"),
     "D": mr("[−2 ; 1]"),
     "ans":"A"},
    
    # Problem 3: (x-4)/(x+2) ≥ 0
    {"num":3,
     "q": mg(mf("x−4","x+2"), mr(" ≥ 0")),
     "A": mr("(−∞ ; −2] ∪ [4 ; +∞)"),
     "B": mr("[−2 ; 4]"),
     "C": mr("(−∞ ; −2) ∪ [4 ; +∞)"),
     "D": mr("(−2 ; 4]"),
     "ans":"C"},
    
    # Problem 4: (x-1)/(x+2) > 0
    {"num":4,
     "q": mg(mf("x−1","x+2"), mr(" > 0")),
     "A": mr("(−∞ ; −2) ∪ (1 ; +∞)"),
     "B": mr("(−2 ; 1)"),
     "C": mr("(−∞ ; −2] ∪ [1 ; +∞)"),
     "D": mr("[−2 ; 1]"),
     "ans":"A"},


    
    # Problem 5: (1-2x)/(x+3) < 0
    {"num":5,
     "q": mg(mf("1−2x","x+3"), mr(" < 0")),
     "A": mr("(−∞ ; −3) ∪ (½ ; +∞)"),
     "B": mr("(−3 ; ½)"),
     "C": mr("(−∞ ; −3] ∪ [½ ; +∞)"),
     "D": mr("[−3 ; ½]"),
     "ans":"A"},
    
    # Problem 6: (3-x)/(1+2x) ≤ 0
    {"num":6,
     "q": mg(mf("3−x","1+2x"), mr(" ≤ 0")),
     "A": mr("(−∞ ; −½) ∪ [3 ; +∞)"),
     "B": mr("(−½ ; 3]"),
     "C": mr("(−∞ ; −½] ∪ [3 ; +∞)"),
     "D": mr("[−½ ; 3]"),
     "ans":"A"},
    
    # Problem 7: x(2-3x) ≥ 0
    {"num":7,
     "q": mr("x(2−3x) ≥ 0"),
     "A": mr("[0 ; ⅔]"),
     "B": mr("(0 ; ⅔)"),
     "C": mr("(−∞ ; 0] ∪ [⅔ ; +∞)"),
     "D": mr("(−∞ ; 0) ∪ (⅔ ; +∞)"),
     "ans":"A"},
    
    # Problem 8: (x-3)(5x+1) > 0
    {"num":8,
     "q": mr("(x−3)(5x+1) > 0"),
     "A": mr("(−∞ ; −⅕) ∪ (3 ; +∞)"),
     "B": mr("(−⅕ ; 3)"),
     "C": mr("(−∞ ; −⅕] ∪ [3 ; +∞)"),
     "D": mr("[−⅕ ; 3]"),
     "ans":"A"},
    
    # Problem 9: (x-2)(6-x) ≤ 0
    {"num":9,
     "q": mr("(x−2)(6−x) ≤ 0"),
     "A": mr("(−∞ ; 2] ∪ [6 ; +∞)"),
     "B": mr("[2 ; 6]"),
     "C": mr("(−∞ ; 2) ∪ (6 ; +∞)"),
     "D": mr("(2 ; 6)"),
     "ans":"A"},


    
    # Problem 10: (7-3x)(5x-3) < 0
    {"num":10,
     "q": mr("(7−3x)(5x−3) < 0"),
     "A": mr("(−∞ ; ⅗) ∪ (⁷⁄₃ ; +∞)"),
     "B": mr("(⅗ ; ⁷⁄₃)"),
     "C": mr("(−∞ ; ⅗] ∪ [⁷⁄₃ ; +∞)"),
     "D": mr("[⅗ ; ⁷⁄₃]"),
     "ans":"A"},
    
    # Problem 11: (2x-1)/(x-5) ≤ 0
    {"num":11,
     "q": mg(mf("2x−1","x−5"), mr(" ≤ 0")),
     "A": mr("[½ ; 5)"),
     "B": mr("(−∞ ; ½] ∪ (5 ; +∞)"),
     "C": mr("(½ ; 5]"),
     "D": mr("[½ ; 5]"),
     "ans":"A"},
    
    # Problem 12: (x+1)(x-1)/(x+3) > 0
    {"num":12,
     "q": mg(mf("(x+1)(x−1)","x+3"), mr(" > 0")),
     "A": mr("(−3 ; −1) ∪ (1 ; +∞)"),
     "B": mr("(−∞ ; −3) ∪ (−1 ; 1)"),
     "C": mr("(−3 ; −1] ∪ [1 ; +∞)"),
     "D": mr("(−∞ ; −3] ∪ [−1 ; 1]"),
     "ans":"A"},
    
    # Problem 13: (x+4)(x-1)/(x-3) ≤ 0
    {"num":13,
     "q": mg(mf("(x+4)(x−1)","x−3"), mr(" ≤ 0")),
     "A": mr("[−4 ; 1] ∪ (3 ; +∞)"),
     "B": mr("(−∞ ; −4] ∪ [1 ; 3)"),
     "C": mr("[−4 ; 1) ∪ (3 ; +∞)"),
     "D": mr("(−4 ; 1] ∪ (3 ; +∞)"),
     "ans":"B"},
    
    # Problem 14: (x-2)(x+4)/(x-1)(x-4) < 0
    {"num":14,
     "q": mg(mf("(x−2)(x+4)","(x−1)(x−4)"), mr(" < 0")),
     "A": mr("(−4 ; 1) ∪ (2 ; 4)"),
     "B": mr("(−∞ ; −4) ∪ (1 ; 2) ∪ (4 ; +∞)"),
     "C": mr("[−4 ; 1) ∪ [2 ; 4)"),
     "D": mr("(−4 ; 1] ∪ (2 ; 4]"),
     "ans":"A"},
]



# Problems 15-30 from page 1
QUESTIONS += [
    # Problem 15: (x+5)/(x-7) ≥ 0
    {"num":15,
     "q": mg(mf("x+5","x−7"), mr(" ≥ 0")),
     "A": mr("(−∞ ; −5] ∪ (7 ; +∞)"),
     "B": mr("[−5 ; 7)"),
     "C": mr("(−∞ ; −5) ∪ [7 ; +∞)"),
     "D": mr("[−5 ; 7]"),
     "ans":"A"},
    
    # Problem 16: (x-2)/(x-6) > 0
    {"num":16,
     "q": mg(mf("x−2","x−6"), mr(" > 0")),
     "A": mr("(−∞ ; 2) ∪ (6 ; +∞)"),
     "B": mr("(2 ; 6)"),
     "C": mr("(−∞ ; 2] ∪ [6 ; +∞)"),
     "D": mr("[2 ; 6]"),
     "ans":"A"},
    
    # Problem 17: (x-3)/(5x-3) < 0
    {"num":17,
     "q": mg(mf("x−3","5x−3"), mr(" < 0")),
     "A": mr("(⅗ ; 3)"),
     "B": mr("(−∞ ; ⅗) ∪ (3 ; +∞)"),
     "C": mr("[⅗ ; 3]"),
     "D": mr("(−∞ ; ⅗] ∪ [3 ; +∞)"),
     "ans":"A"},
    
    # Problem 18: (2x+1)(x-5) < 0
    {"num":18,
     "q": mr("(2x+1)(x−5) < 0"),
     "A": mr("(−½ ; 5)"),
     "B": mr("(−∞ ; −½) ∪ (5 ; +∞)"),
     "C": mr("[−½ ; 5]"),
     "D": mr("(−∞ ; −½] ∪ [5 ; +∞)"),
     "ans":"A"},
    
    # Problem 19: (x+4)(x+1)/(x+2) ≤ 0
    {"num":19,
     "q": mg(mf("(x+4)(x+1)","x+2"), mr(" ≤ 0")),
     "A": mr("(−∞ ; −4] ∪ [−1 ; 3)"),
     "B": mr("[−4 ; −2) ∪ [−1 ; +∞)"),
     "C": mr("(−∞ ; −4] ∪ (−2 ; −1]"),
     "D": mr("[−4 ; −2] ∪ [−1 ; +∞)"),
     "ans":"C"},


    
    # Problem 20: x²-10x+21 ≥ 0
    {"num":20,
     "q": mg(msup("x","2"), mr("−10x+21 ≥ 0")),
     "A": mr("(−∞ ; 3] ∪ [7 ; +∞)"),
     "B": mr("[3 ; 7]"),
     "C": mr("(−∞ ; 3) ∪ (7 ; +∞)"),
     "D": mr("(3 ; 7)"),
     "ans":"A"},
    
    # Problem 21: -x²+2x+3 > 0
    {"num":21,
     "q": mg(mr("−"), msup("x","2"), mr("+2x+3 > 0")),
     "A": mr("(−1 ; 3)"),
     "B": mr("(−∞ ; −1) ∪ (3 ; +∞)"),
     "C": mr("[−1 ; 3]"),
     "D": mr("(−∞ ; −1] ∪ [3 ; +∞)"),
     "ans":"A"},
    
    # Problem 22: x²+x-6 < 0
    {"num":22,
     "q": mg(msup("x","2"), mr("+x−6 < 0")),
     "A": mr("(−3 ; 2)"),
     "B": mr("(−∞ ; −3) ∪ (2 ; +∞)"),
     "C": mr("[−3 ; 2]"),
     "D": mr("(−∞ ; −3] ∪ [2 ; +∞)"),
     "ans":"A"},
    
    # Problem 23: x²-x-20 ≤ 0
    {"num":23,
     "q": mg(msup("x","2"), mr("−x−20 ≤ 0")),
     "A": mr("[−4 ; 5]"),
     "B": mr("(−4 ; 5)"),
     "C": mr("(−∞ ; −4] ∪ [5 ; +∞)"),
     "D": mr("(−∞ ; −4) ∪ (5 ; +∞)"),
     "ans":"A"},
    
    # Problem 24: 2x²-5x-12 ≥ 0
    {"num":24,
     "q": mg(mr("2"), msup("x","2"), mr("−5x−12 ≥ 0")),
     "A": mr("(−∞ ; −³⁄₂] ∪ [4 ; +∞)"),
     "B": mr("[−³⁄₂ ; 4]"),
     "C": mr("(−∞ ; −³⁄₂) ∪ (4 ; +∞)"),
     "D": mr("(−³⁄₂ ; 4)"),
     "ans":"A"},
    
    # Problem 25: -2x²+5x+3 < 0
    {"num":25,
     "q": mg(mr("−2"), msup("x","2"), mr("+5x+3 < 0")),
     "A": mr("(−∞ ; −½) ∪ (3 ; +∞)"),
     "B": mr("(−½ ; 3)"),
     "C": mr("(−∞ ; −½] ∪ [3 ; +∞)"),
     "D": mr("[−½ ; 3]"),
     "ans":"A"},


    
    # Problem 26: 3x²-2x-8 > 0
    {"num":26,
     "q": mg(mr("3"), msup("x","2"), mr("−2x−8 > 0")),
     "A": mr("(−∞ ; −⁴⁄₃) ∪ (2 ; +∞)"),
     "B": mr("(−⁴⁄₃ ; 2)"),
     "C": mr("(−∞ ; −⁴⁄₃] ∪ [2 ; +∞)"),
     "D": mr("[−⁴⁄₃ ; 2]"),
     "ans":"A"},
    
    # Problem 27: (x²+7x+12)/(x²-x-2) > 0
    {"num":27,
     "q": mg(mf("x²+7x+12","x²−x−2"), mr(" > 0")),
     "A": mr("(−4 ; −3) ∪ (−1 ; 2)"),
     "B": mr("(−∞ ; −4) ∪ (−3 ; −1) ∪ (2 ; +∞)"),
     "C": mr("[−4 ; −3] ∪ [−1 ; 2]"),
     "D": mr("(−4 ; −3] ∪ [−1 ; 2)"),
     "ans":"A"},
    
    # Problem 28: x²+6x+8 ≥ 0 
    {"num":28,
     "q": mg(msup("x","2"), mr("+6x+8 ≥ 0")),
     "A": mr("(−∞ ; −4] ∪ [−2 ; +∞)"),
     "B": mr("[−4 ; −2]"),
     "C": mr("(−∞ ; −4) ∪ (−2 ; +∞)"),
     "D": mr("(−4 ; −2)"),
     "ans":"A"},
    
    # Problem 29: x²+x-12 ≤ 0
    {"num":29,
     "q": mg(msup("x","2"), mr("+x−12 ≤ 0")),
     "A": mr("[−4 ; 3]"),
     "B": mr("(−4 ; 3)"),
     "C": mr("(−∞ ; −4] ∪ [3 ; +∞)"),
     "D": mr("(−∞ ; −4) ∪ (3 ; +∞)"),
     "ans":"A"},
    
    # Problem 30: -x²+2x+15 ≥ 0
    {"num":30,
     "q": mg(mr("−"), msup("x","2"), mr("+2x+15 ≥ 0")),
     "A": mr("[−3 ; 5]"),
     "B": mr("(−3 ; 5)"),
     "C": mr("(−∞ ; −3] ∪ [5 ; +∞)"),
     "D": mr("(−∞ ; −3) ∪ (5 ; +∞)"),
     "ans":"A"},
]



# Adding questions from page 2 and 3 (31-56 based on PDF)
QUESTIONS += [
    # Problem 31-50: Additional problems from scanned pages
    {"num":31, "q": mg(mf("x²+2x−8","x+1"), mr(" ≥ 0")),
     "A": mr("(−∞ ; −4] ∪ [2 ; +∞) ∖ {−1}"),
     "B": mr("[−4 ; −1) ∪ [2 ; +∞)"),
     "C": mr("(−∞ ; −4] ∪ (−1 ; 2]"),
     "D": mr("[−4 ; 2]"),
     "ans":"B"},
    
    {"num":32, "q": mg(mf("x²−5x+4","x−3"), mr(" ≤ 0")),
     "A": mr("(−∞ ; 1] ∪ (3 ; 4]"),
     "B": mr("[1 ; 3) ∪ [4 ; +∞)"),
     "C": mr("(−∞ ; 1] ∪ [4 ; +∞)"),
     "D": mr("[1 ; 4]"),
     "ans":"A"},
    
    {"num":33, "q": mg(mf("6−x","x²+5"), mr(" > 0")),
     "A": mr("(−∞ ; 6)"),
     "B": mr("(6 ; +∞)"),
     "C": mr("ℝ"),
     "D": mr("∅"),
     "ans":"A"},
    
    {"num":34, "q": mg(mf("4−3x","x²+7"), mr(" < 0")),
     "A": mr("(⁴⁄₃ ; +∞)"),
     "B": mr("(−∞ ; ⁴⁄₃)"),
     "C": mr("ℝ"),
     "D": mr("∅"),
     "ans":"A"},
    
    {"num":35, "q": mg(mf("x+2","x²−2x+2"), mr(" ≥ 0")),
     "A": mr("[−2 ; +∞)"),
     "B": mr("(−∞ ; −2]"),
     "C": mr("ℝ"),
     "D": mr("∅"),
     "ans":"A"},
    
    # Problem 36-50: More from PDF pages
    {"num":36, "q": mg(mf("2x+3","x²+4x+5"), mr(" ≤ 0")),
     "A": mr("(−∞ ; −³⁄₂]"),
     "B": mr("[−³⁄₂ ; +∞)"),
     "C": mr("ℝ"),
     "D": mr("∅"),
     "ans":"A"},
    
    {"num":37, "q": mg(mf("x−5","x²+3x+4"), mr(" > 0")),
     "A": mr("(5 ; +∞)"),
     "B": mr("(−∞ ; 5)"),
     "C": mr("ℝ"),
     "D": mr("∅"),
     "ans":"A"},
    
    {"num":38, "q": mg(mf("x²−4","x+3"), mr(" ≥ 0")),
     "A": mr("(−∞ ; −3) ∪ [−2 ; 2]"),
     "B": mr("[−3 ; −2] ∪ [2 ; +∞)"),
     "C": mr("(−3 ; +∞)"),
     "D": mr("[−2 ; 2] ∪ (−∞ ; −3)"),
     "ans":"D"},
    
    {"num":39, "q": mg(mf("x²−9","x−5"), mr(" < 0")),
     "A": mr("(−∞ ; −3) ∪ (3 ; 5)"),
     "B": mr("(−3 ; 3) ∪ (5 ; +∞)"),
     "C": mr("(−3 ; 5)"),
     "D": mr("(−∞ ; −3] ∪ [3 ; 5]"),
     "ans":"A"},
    
    {"num":40, "q": mg(mf("x²−16","x²+1"), mr(" ≤ 0")),
     "A": mr("[−4 ; 4]"),
     "B": mr("(−∞ ; −4] ∪ [4 ; +∞)"),
     "C": mr("(−4 ; 4)"),
     "D": mr("ℝ"),
     "ans":"A"},
    
    {"num":41, "q": mg(mf("x²+3x+2","x²−5x+6"), mr(" > 0")),
     "A": mr("(−∞ ; −2) ∪ (−1 ; 2) ∪ (3 ; +∞)"),
     "B": mr("(−2 ; −1) ∪ (2 ; 3)"),
     "C": mr("(−∞ ; −2] ∪ [−1 ; 2] ∪ [3 ; +∞)"),
     "D": mr("(−2 ; −1] ∪ [2 ; 3)"),
     "ans":"A"},
    
    {"num":42, "q": mg(mf("x²−7x+12","x²+x−6"), mr(" ≤ 0")),
     "A": mr("(−∞ ; −3) ∪ [3 ; 4]"),
     "B": mr("[−3 ; 2) ∪ [3 ; 4]"),
     "C": mr("(−3 ; 2) ∪ [3 ; 4]"),
     "D": mr("[3 ; 4]"),
     "ans":"C"},
    
    {"num":43, "q": mg(msup("x","2"), mr("−4x ≥ 0")),
     "A": mr("(−∞ ; 0] ∪ [4 ; +∞)"),
     "B": mr("[0 ; 4]"),
     "C": mr("(−∞ ; 0) ∪ (4 ; +∞)"),
     "D": mr("(0 ; 4)"),
     "ans":"A"},
    
    {"num":44, "q": mg(msup("x","2"), mr("+3x−10 < 0")),
     "A": mr("(−5 ; 2)"),
     "B": mr("(−∞ ; −5) ∪ (2 ; +∞)"),
     "C": mr("[−5 ; 2]"),
     "D": mr("(−∞ ; −5] ∪ [2 ; +∞)"),
     "ans":"A"},
    
    {"num":45, "q": mg(mr("−"), msup("x","2"), mr("+x+12 ≤ 0")),
     "A": mr("(−∞ ; −3] ∪ [4 ; +∞)"),
     "B": mr("[−3 ; 4]"),
     "C": mr("(−∞ ; −3) ∪ (4 ; +∞)"),
     "D": mr("(−3 ; 4)"),
     "ans":"A"},
    
    # Problems 46-56 from page 2
    {"num":46, "q": mg(mf("x+1","x−2"), mr(" ≥ 1")),
     "A": mr("(−∞ ; ³⁄₂] ∪ (2 ; +∞)"),
     "B": mr("[³⁄₂ ; 2)"),
     "C": mr("(−∞ ; ³⁄₂) ∪ (2 ; +∞)"),
     "D": mr("(³⁄₂ ; 2]"),
     "ans":"B"},
    
    {"num":47, "q": mg(mf("2x+1","x−1"), mr(" ≤ 3")),
     "A": mr("[−4 ; 1)"),
     "B": mr("(−∞ ; −4] ∪ (1 ; +∞)"),
     "C": mr("(−4 ; 1]"),
     "D": mr("[−4 ; 1]"),
     "ans":"A"},
    
    {"num":48, "q": mg(mf("5−2x","x+4"), mr(" > 0   [butun yechimlar yig'indisi]")),
     "A": mr("−1"),
     "B": mr("0"),
     "C": mr("1"),
     "D": mr("2"),
     "ans":"B"},
    
    {"num":49, "q": mg(mf("(x+3)(x−2)","(x+1)(x−4)"), mr(" < 0   [butun yechimlar ko'paytmasi]")),
     "A": mr("0"),
     "B": mr("6"),
     "C": mr("−6"),
     "D": mr("12"),
     "ans":"A"},
    
    {"num":50, "q": mg(mf("x²+7x+12","x²+x−2"), mr(" < 0   [oraliqlar yechimi uzunliklari yig'indisi]")),
     "A": mr("2"),
     "B": mr("3"),
     "C": mr("4"),
     "D": mr("5"),
     "ans":"C"},
    
    {"num":51, "q": mg(msup("x","2"), mr("+2x+3 > 0")),
     "A": mr("ℝ"),
     "B": mr("∅"),
     "C": mr("(−∞ ; −1) ∪ (−3 ; +∞)"),
     "D": mr("[−3 ; −1]"),
     "ans":"A"},
    
    {"num":52, "q": mg(mr("6"), msup("x","2"), mr("+5x+4 < 0")),
     "A": mr("∅"),
     "B": mr("ℝ"),
     "C": mr("(−⁴⁄₃ ; −⅔)"),
     "D": mr("[−⁴⁄₃ ; −⅔]"),
     "ans":"A"},
    
    {"num":53, "q": mg(mr("−2"), msup("x","2"), mr("+3x−8 > 0")),
     "A": mr("∅"),
     "B": mr("ℝ"),
     "C": mr("(1 ; 4)"),
     "D": mr("[1 ; 4]"),
     "ans":"A"},
    
    {"num":54, "q": mg(mr("−3"), msup("x","2"), mr("+7x−5 < 0")),
     "A": mr("ℝ"),
     "B": mr("∅"),
     "C": mr("(⅓ ; ⁵⁄₃)"),
     "D": mr("[⅓ ; ⁵⁄₃]"),
     "ans":"A"},
    
    {"num":55, "q": mg(mr("2"), msup("x","2"), mr("+7x+14 > 0")),
     "A": mr("∅"),
     "B": mr("ℝ"),
     "C": mr("(−∞ ; −⁷⁄₂)"),
     "D": mr("[−⁷⁄₂ ; +∞)"),
     "ans":"A"},
    
    {"num":56, "q": mg(mf("3x²+2x+7","x−2"), mr(" > 0")),
     "A": mr("(2 ; +∞)"),
     "B": mr("(−∞ ; 2)"),
     "C": mr("ℝ"),
     "D": mr("∅"),
     "ans":"A"},
]



# ─── XML template strings ─────────────────────────────────────────────────────

CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

WORD_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>'''

SETTINGS_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W}">
  <w:defaultTabStop w:val="720"/>
</w:settings>'''

STYLES_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="24"/><w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>'''



# ─── Document XML builder ─────────────────────────────────────────────────────

def text_run_xml(text, bold=False, size=24, font="Times New Roman"):
    """Return XML string for a plain text run."""
    safe = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    b_tag = "<w:b/><w:bCs/>" if bold else ""
    return (f'<w:r xmlns:w="{W}">'
            f'<w:rPr>{b_tag}'
            f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
            f'</w:rPr>'
            f'<w:t xml:space="preserve">{safe}</w:t>'
            f'</w:r>')

def para_xml(inner_runs_xml, spacing_after=80, jc="left", indent_left=0):
    """Return XML string for a paragraph."""
    jc_tag = f'<w:jc w:val="{jc}"/>' if jc != "left" else ""
    ind_tag = f'<w:ind w:left="{indent_left}"/>' if indent_left else ""
    return (f'<w:p xmlns:w="{W}" xmlns:m="{M}">'
            f'<w:pPr><w:spacing w:after="{spacing_after}"/>{jc_tag}{ind_tag}</w:pPr>'
            f'{inner_runs_xml}'
            f'</w:p>')

def separator_para_xml():
    return (f'<w:p xmlns:w="{W}">'
            f'<w:pPr><w:spacing w:after="60"/></w:pPr>'
            f'{text_run_xml("", size=18)}'
            f'</w:p>')

def page_break_xml():
    return (f'<w:p xmlns:w="{W}"><w:r><w:br w:type="page"/></w:r></w:p>')

def table_row_xml(cells, col_w=2340):
    tcs = ""
    for cell in cells:
        tcs += (f'<w:tc xmlns:w="{W}">'
                f'<w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>'
                f'<w:tcBorders>'
                f'<w:top w:val="single" w:sz="4"/>'
                f'<w:left w:val="single" w:sz="4"/>'
                f'<w:bottom w:val="single" w:sz="4"/>'
                f'<w:right w:val="single" w:sz="4"/>'
                f'</w:tcBorders></w:tcPr>'
                f'<w:p xmlns:w="{W}"><w:pPr><w:jc w:val="center"/></w:pPr>'
                f'{text_run_xml(cell, bold=True, size=22)}'
                f'</w:p></w:tc>')
    return f'<w:tr xmlns:w="{W}">{tcs}</w:tr>'



def rebalance_answers(questions):
    """Shuffle correct answers to distribute evenly across A/B/C/D."""
    import random
    random.seed(42)
    target_seq = []
    for i in range(len(questions)):
        target_seq.append(["A","B","C","D"][i % 4])
    
    balanced = []
    for q, target_ans in zip(questions, target_seq):
        current_ans = q["ans"]
        if current_ans == target_ans:
            balanced.append(q)
            continue
        opts = ["A","B","C","D"]
        vals = {L: q[L] for L in opts}
        correct_val = vals[current_ans]
        target_val  = vals[target_ans]
        new_q = dict(q)
        new_q[current_ans] = target_val
        new_q[target_ans]  = correct_val
        new_q["ans"] = target_ans
        balanced.append(new_q)
    return balanced

def build_document(questions):
    """Build the complete word/document.xml string."""
    body_parts = []
    
    # Title
    body_parts.append(para_xml(
        text_run_xml("4. RATSIONAL TENGSIZLIKLAR", bold=True, size=32),
        spacing_after=160, jc="center"))
    body_parts.append(para_xml(
        text_run_xml("Professional A/B/C/D Test", bold=False, size=26),
        spacing_after=120, jc="center"))
    body_parts.append(para_xml(
        text_run_xml("To'g'ri javobni tanlang.", bold=False, size=22, font="Times New Roman"),
        spacing_after=200, jc="center"))
    
    opt_letters = ["A","B","C","D"]
    
    for q in questions:
        num = q["num"]
        q_xml = q["q"]
        
        # Question line
        q_line = (
            text_run_xml(f"{num}.  ", bold=True, size=24) +
            omml_inline(q_xml)
        )
        body_parts.append(para_xml(q_line, spacing_after=80))
        
        # Options
        for letter in opt_letters:
            opt_line = (
                text_run_xml(f"    {letter})  ", bold=True, size=24) +
                omml_inline(q[letter])
            )
            body_parts.append(para_xml(opt_line, spacing_after=50, indent_left=360))
        
        body_parts.append(separator_para_xml())


    
    # Answer Key page
    body_parts.append(page_break_xml())
    body_parts.append(para_xml(
        text_run_xml("JAVOBLAR KALITI  —  ANSWER KEY", bold=True, size=30),
        spacing_after=100, jc="center"))
    body_parts.append(para_xml(
        text_run_xml("4. Ratsional Tengsizliklar", bold=True, size=24),
        spacing_after=180, jc="center"))
    
    # Answer key table
    COLS = 4
    rows = []
    row = []
    for q in questions:
        row.append(f"{q['num']}) — {q['ans']}")
        if len(row) == COLS:
            rows.append(row)
            row = []
    if row:
        while len(row) < COLS:
            row.append("")
        rows.append(row)
    
    tbl_rows_xml = "".join(table_row_xml(r) for r in rows)
    tbl_xml = (f'<w:tbl xmlns:w="{W}">'
               f'<w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr>'
               f'{tbl_rows_xml}</w:tbl>')
    body_parts.append(tbl_xml)
    
    # Distribution
    dist = Counter(q['ans'] for q in questions)
    body_parts.append(para_xml(text_run_xml(""), spacing_after=80))
    body_parts.append(para_xml(
        text_run_xml(f"Javoblar taqsimoti:  A={dist['A']}  |  B={dist['B']}  |  C={dist['C']}  |  D={dist['D']}", size=22),
        spacing_after=100, jc="center"))
    
    # Section properties
    sect_xml = (f'<w:sectPr xmlns:w="{W}">'
                f'<w:pgSz w:w="12240" w:h="15840"/>'
                f'<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1700"/>'
                f'</w:sectPr>')
    
    body_content = "\n".join(body_parts) + "\n" + sect_xml
    
    doc_xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               f'<w:document xmlns:w="{W}" xmlns:m="{M}">\n'
               f'<w:body>\n{body_content}\n</w:body>\n</w:document>')
    return doc_xml



def write_docx(output_path, questions):
    """Create the .docx file."""
    doc_xml = build_document(questions)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", RELS_XML)
        zf.writestr("word/_rels/document.xml.rels", WORD_RELS_XML)
        zf.writestr("word/document.xml", doc_xml.encode("utf-8"))
        zf.writestr("word/styles.xml", STYLES_XML)
        zf.writestr("word/settings.xml", SETTINGS_XML)
    
    print(f"✅  Fayl tayyor: {output_path}")
    print(f"   Savollar soni: {len(questions)}")
    dist = Counter(q['ans'] for q in questions)
    print(f"   Javoblar: A={dist['A']}  B={dist['B']}  C={dist['C']}  D={dist['D']}")

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "Ratsional_Tengsizliklar_Test.docx"
    
    print(f"Jami savollar soni: {len(QUESTIONS)}")
    balanced = rebalance_answers(QUESTIONS)
    
    dist_before = Counter(q['ans'] for q in QUESTIONS)
    dist_after = Counter(q['ans'] for q in balanced)
    
    print(f"Balanslangandan oldin: A={dist_before['A']}  B={dist_before['B']}  C={dist_before['C']}  D={dist_before['D']}")
    print(f"Balanslangandan keyin: A={dist_after['A']}  B={dist_after['B']}  C={dist_after['C']}  D={dist_after['D']}")
    
    write_docx(out, balanced)
    print(f"\n✅ Test muvaffaqiyatli yaratildi: {out}")
    print("   Microsoft Word'da ochib, matematik ifodalarni tekshiring!")
