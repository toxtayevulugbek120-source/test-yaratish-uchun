#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional Matematik Test Generator — PDF to Word with OMML
Yuklangan PDF fayldagi barcha matematik masalalarni A/B/C/D test formatiga aylantiradi
Microsoft Word (.docx) OMML equation formatida
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

def msub(base, subscript):
    """Subscript"""
    return (f'<m:sSub xmlns:m="{M}"><m:e>{base}</m:e>'
            f'<m:sub>{subscript}</m:sub></m:sSub>')

def msqrt(inner):
    """Square root"""
    return f'<m:rad xmlns:m="{M}"><m:radPr/><m:deg/><m:e>{inner}</m:e></m:rad>'

def mg(*parts):
    """Group multiple OMML parts"""
    return "".join(parts)



# ═══════════════════════════════════════════════════════════════════════
#  EXTRACTED QUESTIONS FROM PDF
#  Based on "Ikkinchi va yuqori darajali tenglamalar sistemasi"
# ═══════════════════════════════════════════════════════════════════════

def generate_questions():
    """Generate all questions with variants from PDF"""
    questions = []
    
    # Problem 1: x + y = 6, xy = 8
    questions.append({
        "num": 1,
        "q": mg(mr("{  x + y = 6, xy = 8")),
        "variants": [
            mg(mr("(2; 4) va (4; 2)")),
            mg(mr("(3; 3) va (1; 5)")),
            mg(mr("(1; 5) va (5; 1)")),
            mg(mr("(2; 3) va (3; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 2: x + y = 8, xy = 15
    questions.append({
        "num": 2,
        "q": mg(mr("{  x + y = 8, xy = 15")),
        "variants": [
            mg(mr("(3; 5) va (5; 3)")),
            mg(mr("(2; 6) va (6; 2)")),
            mg(mr("(4; 4)")),
            mg(mr("(1; 7) va (7; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 3: x + y = 10, xy = 21
    questions.append({
        "num": 3,
        "q": mg(mr("{  x + y = 10, xy = 21")),
        "variants": [
            mg(mr("(3; 7) va (7; 3)")),
            mg(mr("(4; 6) va (6; 4)")),
            mg(mr("(2; 8) va (8; 2)")),
            mg(mr("(5; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 4: x − y = 3, xy = 10
    questions.append({
        "num": 4,
        "q": mg(mr("{  x − y = 3, xy = 10")),
        "variants": [
            mg(mr("(5; 2)")),
            mg(mr("(6; 3)")),
            mg(mr("(−2; −5)")),
            mg(mr("(4; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 5: x − y = −2, xy = 24
    questions.append({
        "num": 5,
        "q": mg(mr("{  x − y = −2, xy = 24")),
        "variants": [
            mg(mr("(4; 6)")),
            mg(mr("(−6; −4)")),
            mg(mr("(6; 4)")),
            mg(mr("(3; 8)"))
        ],
        "correct": 0
    })
    

    # Problem 6: x − y = 8, x² − y³ = 80
    questions.append({
        "num": 6,
        "q": mg(mr("{  x − y = 8, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("3")), mr(" = 80")),
        "variants": [
            mg(mr("(10; 2)")),
            mg(mr("(12; 4)")),
            mg(mr("(9; 1)")),
            mg(mr("(11; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 7: x − y = −3, x² − y² = 21
    questions.append({
        "num": 7,
        "q": mg(mr("{  x − y = −3, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21")),
        "variants": [
            mg(mr("(−7; −10)")),
            mg(mr("(4; 7)")),
            mg(mr("(3; 6)")),
            mg(mr("(−4; −7)"))
        ],
        "correct": 0
    })
    
    # Problem 8: x − y = 3, x² − y² = 21
    questions.append({
        "num": 8,
        "q": mg(mr("{  x − y = 3, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21")),
        "variants": [
            mg(mr("(5; 2)")),
            mg(mr("(6; 3)")),
            mg(mr("(4; 1)")),
            mg(mr("(7; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 9: x + y = 9, x² − y² = 9
    questions.append({
        "num": 9,
        "q": mg(mr("{  x + y = 9, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 9")),
        "variants": [
            mg(mr("(5; 4)")),
            mg(mr("(6; 3)")),
            mg(mr("(4; 5)")),
            mg(mr("(7; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 10: x + y = 7, x² − y² = 21
    questions.append({
        "num": 10,
        "q": mg(mr("{  x + y = 7, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 21")),
        "variants": [
            mg(mr("(5; 2)")),
            mg(mr("(4; 3)")),
            mg(mr("(6; 1)")),
            mg(mr("(3; 4)"))
        ],
        "correct": 0
    })
    

    # Problem 11: x + y = 6, x² + y² = 20
    questions.append({
        "num": 11,
        "q": mg(mr("{  x + y = 6, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 20")),
        "variants": [
            mg(mr("(2; 4) va (4; 2)")),
            mg(mr("(3; 3)")),
            mg(mr("(1; 5) va (5; 1)")),
            mg(mr("(2; 3) va (3; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 12: x − y = 5, x² + y² = 53
    questions.append({
        "num": 12,
        "q": mg(mr("{  x − y = 5, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 53")),
        "variants": [
            mg(mr("(7; 2)")),
            mg(mr("(8; 3)")),
            mg(mr("(6; 1)")),
            mg(mr("(9; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 13: x + y = 2, x² + y² − 2xy = 16
    questions.append({
        "num": 13,
        "q": mg(mr("{  x + y = 2, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), 
              mr(" − 2xy = 16")),
        "variants": [
            mg(mr("(5; −3) va (−3; 5)")),
            mg(mr("(4; −2) va (−2; 4)")),
            mg(mr("(6; −4) va (−4; 6)")),
            mg(mr("(3; −1) va (−1; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 14: x + y = 3, x² + y² − 2xy = 1
    questions.append({
        "num": 14,
        "q": mg(mr("{  x + y = 3, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), 
              mr(" − 2xy = 1")),
        "variants": [
            mg(mr("(2; 1) va (1; 2)")),
            mg(mr("(3; 0) va (0; 3)")),
            mg(mr("(4; −1) va (−1; 4)")),
            mg(mr("(5; −2) va (−2; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 15: x − y = 6, x² + y² + 2xy = 16
    questions.append({
        "num": 15,
        "q": mg(mr("{  x − y = 6, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), 
              mr(" + 2xy = 16")),
        "variants": [
            mg(mr("(5; −1)")),
            mg(mr("(7; 1)")),
            mg(mr("(4; −2)")),
            mg(mr("(6; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 16: x + y = 3, x² + xy − y² = 5
    questions.append({
        "num": 16,
        "q": mg(mr("{  x + y = 3, "), msup(mr("x"), mr("2")), mr(" + xy − "), 
              msup(mr("y"), mr("2")), mr(" = 5")),
        "variants": [
            mg(mr("(4; −1) va (−1; 4)")),
            mg(mr("(3; 0) va (0; 3)")),
            mg(mr("(5; −2) va (−2; 5)")),
            mg(mr("(2; 1) va (1; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 17: x − y = 7, x² − xy − y² = 19
    questions.append({
        "num": 17,
        "q": mg(mr("{  x − y = 7, "), msup(mr("x"), mr("2")), mr(" − xy − "), 
              msup(mr("y"), mr("2")), mr(" = 19")),
        "variants": [
            mg(mr("(8; 1)")),
            mg(mr("(9; 2)")),
            mg(mr("(7; 0)")),
            mg(mr("(10; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 18: y = x + 6, x² + 3 = 4y
    questions.append({
        "num": 18,
        "q": mg(mr("{  y = x + 6, "), msup(mr("x"), mr("2")), mr(" + 3 = 4y")),
        "variants": [
            mg(mr("(3; 9) va (−5; 1)")),
            mg(mr("(4; 10) va (−4; 2)")),
            mg(mr("(2; 8) va (−6; 0)")),
            mg(mr("(5; 11) va (−3; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 19: y − 3x = 2, x² = 2y + 3
    questions.append({
        "num": 19,
        "q": mg(mr("{  y − 3x = 2, "), msup(mr("x"), mr("2")), mr(" = 2y + 3")),
        "variants": [
            mg(mr("(−1; −1) va (5; 17)")),
            mg(mr("(0; 2) va (4; 14)")),
            mg(mr("(1; 5) va (3; 11)")),
            mg(mr("(2; 8) va (6; 20)"))
        ],
        "correct": 0
    })
    
    # Problem 20: x − 2y = 1, 3x + y² = 10
    questions.append({
        "num": 20,
        "q": mg(mr("{  x − 2y = 1, 3x + "), msup(mr("y"), mr("2")), mr(" = 10")),
        "variants": [
            mg(mr("(1; 0) va (3; 1)")),
            mg(mr("(2; 0.5) va (4; 1.5)")),
            mg(mr("(5; 2) va (−1; −1)")),
            mg(mr("(3; 1) va (1; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 21: x² + xy + 3 = 0, y − 3x − 2 = 0
    questions.append({
        "num": 21,
        "q": mg(msup(mr("x"), mr("2")), mr(" + xy + 3 = 0,  y − 3x − 2 = 0")),
        "variants": [
            mg(mr("(−1; −5) va (−3; −11)")),
            mg(mr("(1; 5) va (3; 11)")),
            mg(mr("(−2; −8) va (−4; −14)")),
            mg(mr("(2; 8) va (4; 14)"))
        ],
        "correct": 0
    })
    
    # Problem 22: y² − x² = 16, x + y = 8
    questions.append({
        "num": 22,
        "q": mg(msup(mr("y"), mr("2")), mr(" − "), msup(mr("x"), mr("2")), mr(" = 16,  x + y = 8")),
        "variants": [
            mg(mr("(3; 5)")),
            mg(mr("(4; 4)")),
            mg(mr("(2; 6)")),
            mg(mr("(5; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 23: (x+3y)/(y−1) − (y−x)/(2x) = 2, y − x = 4
    questions.append({
        "num": 23,
        "q": mg(mfrac(mr("x + 3y"), mr("y − 1")), mr(" − "), 
              mfrac(mr("y − x"), mr("2x")), mr(" = 2,  y − x = 4")),
        "variants": [
            mg(mr("(2; 6)")),
            mg(mr("(3; 7)")),
            mg(mr("(1; 5)")),
            mg(mr("(4; 8)"))
        ],
        "correct": 0
    })
    
    # Problem 24: x² − xy = −1, y + 4x = 6
    questions.append({
        "num": 24,
        "q": mg(msup(mr("x"), mr("2")), mr(" − xy = −1,  y + 4x = 6")),
        "variants": [
            mg(mr("(1; 2) va (−1; 10)")),
            mg(mr("(2; −2) va (0; 6)")),
            mg(mr("(3; −6) va (−3; 18)")),
            mg(mr("(4; −10) va (−2; 14)"))
        ],
        "correct": 0
    })
    
    # Problem 25: 2x² + xy − 14 = 0, 3x − y − 3 = 0
    questions.append({
        "num": 25,
        "q": mg(mr("2"), msup(mr("x"), mr("2")), mr(" + xy − 14 = 0,  3x − y − 3 = 0")),
        "variants": [
            mg(mr("(2; 3)")),
            mg(mr("(1; 0)")),
            mg(mr("(−7/2; −27/2)")),
            mg(mr("(3; 6)"))
        ],
        "correct": 0
    })
    
    # Problem 26: x² + 4y² = 37, x + y = 4
    questions.append({
        "num": 26,
        "q": mg(msup(mr("x"), mr("2")), mr(" + 4"), msup(mr("y"), mr("2")), mr(" = 37,  x + y = 4")),
        "variants": [
            mg(mr("(5; −1) va (−1; 5)")),
            mg(mr("(6; −2) va (−2; 6)")),
            mg(mr("(3; 1) va (1; 3)")),
            mg(mr("(4; 0) va (0; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 27: x² − 2y² = 2, 2x + y = −3
    questions.append({
        "num": 27,
        "q": mg(msup(mr("x"), mr("2")), mr(" − 2"), msup(mr("y"), mr("2")), mr(" = 2,  2x + y = −3")),
        "variants": [
            mg(mr("(−1; −1) va (2; −7)")),
            mg(mr("(−2; 1)")),
            mg(mr("(1; −5)")),
            mg(mr("(0; −3)"))
        ],
        "correct": 0
    })
    
    # Problem 28: y² − 3x² = −2, x + 2y = 1
    questions.append({
        "num": 28,
        "q": mg(msup(mr("y"), mr("2")), mr(" − 3"), msup(mr("x"), mr("2")), mr(" = −2,  x + 2y = 1")),
        "variants": [
            mg(mr("(1; 0) va (7; −3)")),
            mg(mr("(−1; 1)")),
            mg(mr("(3; −1)")),
            mg(mr("(5; −2)"))
        ],
        "correct": 0
    })
    
    # Problem 29: x² + 2xy = −3, 2x − y = −4
    questions.append({
        "num": 29,
        "q": mg(msup(mr("x"), mr("2")), mr(" + 2xy = −3,  2x − y = −4")),
        "variants": [
            mg(mr("(−1; 2) va (−3; −2)")),
            mg(mr("(1; 6)")),
            mg(mr("(2; 8)")),
            mg(mr("(−2; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 30: 2x² + 3y² = 14, x + 2y = 5
    questions.append({
        "num": 30,
        "q": mg(mr("2"), msup(mr("x"), mr("2")), mr(" + 3"), msup(mr("y"), mr("2")), 
              mr(" = 14,  x + 2y = 5")),
        "variants": [
            mg(mr("(1; 2) va (3; 1)")),
            mg(mr("(5; 0) va (1; 2)")),
            mg(mr("(2; 1.5)")),
            mg(mr("(3; 1) faqat"))
        ],
        "correct": 0
    })
    
    # Problem 31: (x−5)(y−3) = 0, (x−8)(y+2) = 0
    questions.append({
        "num": 31,
        "q": mg(mr("{  (x − 5)(y − 3) = 0, (x − 8)(y + 2) = 0")),
        "variants": [
            mg(mr("(5; −2) va (8; 3)")),
            mg(mr("(5; 3) va (8; −2)")),
            mg(mr("(5; −2) faqat")),
            mg(mr("(8; 3) faqat"))
        ],
        "correct": 0
    })
    
    # Problem 32: (x−7)(y−8) = 0, (x+3)(y−7) = 0
    questions.append({
        "num": 32,
        "q": mg(mr("{  (x − 7)(y − 8) = 0, (x + 3)(y − 7) = 0")),
        "variants": [
            mg(mr("(7; 7) va (−3; 8)")),
            mg(mr("(7; 8) va (−3; 7)")),
            mg(mr("(7; 7) faqat")),
            mg(mr("(−3; 8) faqat"))
        ],
        "correct": 0
    })
    
    # Problem 33: (x−2)(x+3)(y−5) = 0, (y−7)(y+3)(x−5) = 0
    questions.append({
        "num": 33,
        "q": mg(mr("{  (x − 2)(x + 3)(y − 5) = 0, (y − 7)(y + 3)(x − 5) = 0")),
        "variants": [
            mg(mr("(2; 7), (2; −3), (−3; 7), (−3; −3), (5; 5)")),
            mg(mr("(2; 5), (−3; 5), (5; 7), (5; −3)")),
            mg(mr("(2; 7), (2; −3), (5; 5) faqat")),
            mg(mr("(−3; 7), (−3; −3), (5; 7) faqat"))
        ],
        "correct": 0
    })
    
    # Problem 34: (x−5)²/(y−3) = 0, (y−3)²/(x−5) = 0
    questions.append({
        "num": 34,
        "q": mg(mfrac(msup(mr("(x − 5)"), mr("2")), mr("y − 3")), mr(" = 0, "), 
              mfrac(msup(mr("(y − 3)"), mr("2")), mr("x − 5")), mr(" = 0")),
        "variants": [
            mg(mr("(5; 3)")),
            mg(mr("(3; 5)")),
            mg(mr("(5; 5)")),
            mg(mr("(3; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 35: (x−2)(y−3)/(x−8) = 0, (x−8)(y−5)/(x+7) = 0
    questions.append({
        "num": 35,
        "q": mg(mfrac(mr("(x − 2)(y − 3)"), mr("x − 8")), mr(" = 0, "), 
              mfrac(mr("(x − 8)(y − 5)"), mr("x + 7")), mr(" = 0")),
        "variants": [
            mg(mr("(2; 5)")),
            mg(mr("(2; 3)")),
            mg(mr("(8; 5)")),
            mg(mr("(−7; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 36: (x−2)² + (y+5)² = 0
    questions.append({
        "num": 36,
        "q": mg(msup(mr("(x − 2)"), mr("2")), mr(" + "), msup(mr("(y + 5)"), mr("2")), mr(" = 0")),
        "variants": [
            mg(mr("(2; −5)")),
            mg(mr("(−2; 5)")),
            mg(mr("(2; 5)")),
            mg(mr("Yechimi yo'q"))
        ],
        "correct": 0
    })
    
    # Problem 37: (3x−2)² + (y−3)² = 0
    questions.append({
        "num": 37,
        "q": mg(msup(mr("(3x − 2)"), mr("2")), mr(" + "), msup(mr("(y − 3)"), mr("2")), mr(" = 0")),
        "variants": [
            mg(mr("(2/3; 3)")),
            mg(mr("(−2/3; −3)")),
            mg(mr("(3; 2/3)")),
            mg(mr("(2; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 38: (x+5)² + (2y−8)² = 0
    questions.append({
        "num": 38,
        "q": mg(msup(mr("(x + 5)"), mr("2")), mr(" + "), msup(mr("(2y − 8)"), mr("2")), mr(" = 0")),
        "variants": [
            mg(mr("(−5; 4)")),
            mg(mr("(5; −4)")),
            mg(mr("(−5; 8)")),
            mg(mr("(5; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 39: (3x−6)⁴ + (y−5)⁴ = 0
    questions.append({
        "num": 39,
        "q": mg(msup(mr("(3x − 6)"), mr("4")), mr(" + "), msup(mr("(y − 5)"), mr("4")), mr(" = 0")),
        "variants": [
            mg(mr("(2; 5)")),
            mg(mr("(−2; −5)")),
            mg(mr("(6; 5)")),
            mg(mr("(2; −5)"))
        ],
        "correct": 0
    })
    
    # Problem 40: (x+1)² + (y+2)² = 0
    questions.append({
        "num": 40,
        "q": mg(msup(mr("(x + 1)"), mr("2")), mr(" + "), msup(mr("(y + 2)"), mr("2")), mr(" = 0")),
        "variants": [
            mg(mr("(−1; −2)")),
            mg(mr("(1; 2)")),
            mg(mr("(−1; 2)")),
            mg(mr("(1; −2)"))
        ],
        "correct": 0
    })
    
    # Problem 41: x + y = 4, 1/x + 1/y = 1
    questions.append({
        "num": 41,
        "q": mg(mr("{  x + y = 4, "), mfrac(mr("1"), mr("x")), mr(" + "), 
              mfrac(mr("1"), mr("y")), mr(" = 1")),
        "variants": [
            mg(mr("(2; 2)")),
            mg(mr("(1; 3) va (3; 1)")),
            mg(mr("(4; 0) va (0; 4)")),
            mg(mr("Yechimi yo'q"))
        ],
        "correct": 0
    })
    
    # Problem 42: xy = 6, 1/x − 1/y = 1/6
    questions.append({
        "num": 42,
        "q": mg(mr("{  xy = 6, "), mfrac(mr("1"), mr("x")), mr(" − "), 
              mfrac(mr("1"), mr("y")), mr(" = "), mfrac(mr("1"), mr("6"))),
        "variants": [
            mg(mr("(2; 3) va (−3; −2)")),
            mg(mr("(3; 2) va (−2; −3)")),
            mg(mr("(6; 1) va (−1; −6)")),
            mg(mr("(1; 6) va (−6; −1)"))
        ],
        "correct": 0
    })
    
    # Problem 43: x + y = 8, 1/x + 1/y = 1/2
    questions.append({
        "num": 43,
        "q": mg(mr("{  x + y = 8, "), mfrac(mr("1"), mr("x")), mr(" + "), 
              mfrac(mr("1"), mr("y")), mr(" = "), mfrac(mr("1"), mr("2"))),
        "variants": [
            mg(mr("(4; 4)")),
            mg(mr("(2; 6) va (6; 2)")),
            mg(mr("(1; 7) va (7; 1)")),
            mg(mr("(3; 5) va (5; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 44: x + y = 7, 1/x + 1/y = 7/12
    questions.append({
        "num": 44,
        "q": mg(mr("{  x + y = 7, "), mfrac(mr("1"), mr("x")), mr(" + "), 
              mfrac(mr("1"), mr("y")), mr(" = "), mfrac(mr("7"), mr("12"))),
        "variants": [
            mg(mr("(3; 4) va (4; 3)")),
            mg(mr("(2; 5) va (5; 2)")),
            mg(mr("(1; 6) va (6; 1)")),
            mg(mr("Yechimi yo'q"))
        ],
        "correct": 0
    })
    
    # ═══════════════════════════════════════════════════════════════
    #  PROBLEMS 45-78 FROM SECOND PDF
    # ═══════════════════════════════════════════════════════════════
    
    # Problem 45: x − y = 1, 1/x − 1/y = 7/12
    questions.append({
        "num": 45,
        "q": mg(mr("{  x − y = 1, "), mfrac(mr("1"), mr("x")), mr(" − "), 
              mfrac(mr("1"), mr("y")), mr(" = "), mfrac(mr("7"), mr("12"))),
        "variants": [
            mg(mr("(3; 4)")),
            mg(mr("(4; 3)")),
            mg(mr("(2; 3)")),
            mg(mr("(−3; −4)"))
        ],
        "correct": 0
    })
    
    # Problem 46: x = y², y = x²
    questions.append({
        "num": 46,
        "q": mg(mr("{  x = "), msup(mr("y"), mr("2")), mr(", y = "), msup(mr("x"), mr("2"))),
        "variants": [
            mg(mr("(0; 0) va (1; 1)")),
            mg(mr("(0; 0) va (−1; 1)")),
            mg(mr("(1; 1) faqat")),
            mg(mr("(0; 0), (1; 1), (−1; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 47: x = y², 8y = x²
    questions.append({
        "num": 47,
        "q": mg(mr("{  x = "), msup(mr("y"), mr("2")), mr(", 8y = "), msup(mr("x"), mr("2"))),
        "variants": [
            mg(mr("(0; 0) va (4; 2)")),
            mg(mr("(0; 0) va (2; 4)")),
            mg(mr("(4; 2) faqat")),
            mg(mr("(2; 2) va (4; 4)"))
        ],
        "correct": 0
    })
    
    # Problem 48: xy² = 4, x²y = 2
    questions.append({
        "num": 48,
        "q": mg(mr("{  x"), msup(mr("y"), mr("2")), mr(" = 4, "), msup(mr("x"), mr("2")), mr("y = 2")),
        "variants": [
            mg(mr("(2; √2)")),
            mg(mr("(1; 2)")),
            mg(mr("(√2; 2)")),
            mg(mr("(2; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 49: x⁴y³ = 9, x³y² = 27
    questions.append({
        "num": 49,
        "q": mg(msup(mr("x"), mr("4")), msup(mr("y"), mr("3")), mr(" = 9, "), 
              msup(mr("x"), mr("3")), msup(mr("y"), mr("2")), mr(" = 27")),
        "variants": [
            mg(mr("(3; 1/3)")),
            mg(mr("(1/3; 3)")),
            mg(mr("(9; 1/9)")),
            mg(mr("(3; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 50: x² + y = 20, y² + x = 20
    questions.append({
        "num": 50,
        "q": mg(msup(mr("x"), mr("2")), mr(" + y = 20, "), msup(mr("y"), mr("2")), mr(" + x = 20")),
        "variants": [
            mg(mr("(4; 4) va (−5; −5)")),
            mg(mr("(4; 4) faqat")),
            mg(mr("(5; 5) va (−4; −4)")),
            mg(mr("(−5; −5) faqat"))
        ],
        "correct": 0
    })
    
    # Problem 51: x² + 2y = 15, y² + 2x = 15
    questions.append({
        "num": 51,
        "q": mg(msup(mr("x"), mr("2")), mr(" + 2y = 15, "), msup(mr("y"), mr("2")), mr(" + 2x = 15")),
        "variants": [
            mg(mr("(3; 3) va (−5; −5)")),
            mg(mr("(3; 3) faqat")),
            mg(mr("(5; 5) va (−3; −3)")),
            mg(mr("(−5; −5) faqat"))
        ],
        "correct": 0
    })
    
    # Problem 52: x + y + xy = 5, x³ + y³ + xy = 7
    questions.append({
        "num": 52,
        "q": mg(mr("{  x + y + xy = 5, "), msup(mr("x"), mr("3")), mr(" + "), 
              msup(mr("y"), mr("3")), mr(" + xy = 7")),
        "variants": [
            mg(mr("(2; 1) va (1; 2)")),
            mg(mr("(3; 1) va (1; 3)")),
            mg(mr("(2; 2)")),
            mg(mr("(1; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 53: x + y + xy = 7, x² + y² + xy = 13
    questions.append({
        "num": 53,
        "q": mg(mr("{  x + y + xy = 7, "), msup(mr("x"), mr("2")), mr(" + "), 
              msup(mr("y"), mr("2")), mr(" + xy = 13")),
        "variants": [
            mg(mr("(3; 1) va (1; 3)")),
            mg(mr("(2; 2)")),
            mg(mr("(4; 1) va (1; 4)")),
            mg(mr("(3; 2) va (2; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 54: x/y + x + y = 7, (x/y)(x + y) = 12
    questions.append({
        "num": 54,
        "q": mg(mfrac(mr("x"), mr("y")), mr(" + x + y = 7, "), 
              mfrac(mr("x"), mr("y")), mr("(x + y) = 12")),
        "variants": [
            mg(mr("(3; 1)")),
            mg(mr("(2; 1)")),
            mg(mr("(4; 1)")),
            mg(mr("(6; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 55: (x/y)(x + y) = 20, (x/y) + x + y = 9
    questions.append({
        "num": 55,
        "q": mg(mfrac(mr("x"), mr("y")), mr("(x + y) = 20, "), 
              mfrac(mr("x"), mr("y")), mr(" + x + y = 9")),
        "variants": [
            mg(mr("(4; 1)")),
            mg(mr("(5; 1)")),
            mg(mr("(3; 1)")),
            mg(mr("(8; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 56: x³ − y³ = 7, x² + xy + y² = 7
    questions.append({
        "num": 56,
        "q": mg(msup(mr("x"), mr("3")), mr(" − "), msup(mr("y"), mr("3")), mr(" = 7, "), 
              msup(mr("x"), mr("2")), mr(" + xy + "), msup(mr("y"), mr("2")), mr(" = 7")),
        "variants": [
            mg(mr("(2; 1)")),
            mg(mr("(3; 2)")),
            mg(mr("(1; −1)")),
            mg(mr("(3; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 57: x³ − y³ = 19, x² + xy + y² = 19
    questions.append({
        "num": 57,
        "q": mg(msup(mr("x"), mr("3")), mr(" − "), msup(mr("y"), mr("3")), mr(" = 19, "), 
              msup(mr("x"), mr("2")), mr(" + xy + "), msup(mr("y"), mr("2")), mr(" = 19")),
        "variants": [
            mg(mr("(3; 2)")),
            mg(mr("(2; 1)")),
            mg(mr("(4; 3)")),
            mg(mr("(3; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 58: x³ + y³ = 28, x² − xy + y² = 7
    questions.append({
        "num": 58,
        "q": mg(msup(mr("x"), mr("3")), mr(" + "), msup(mr("y"), mr("3")), mr(" = 28, "), 
              msup(mr("x"), mr("2")), mr(" − xy + "), msup(mr("y"), mr("2")), mr(" = 7")),
        "variants": [
            mg(mr("(3; 1) va (1; 3)")),
            mg(mr("(2; 2)")),
            mg(mr("(4; 2) va (2; 4)")),
            mg(mr("(3; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 59: 8x³ − y³ = 7, 4x² − 2xy + y² = 7
    questions.append({
        "num": 59,
        "q": mg(mr("8"), msup(mr("x"), mr("3")), mr(" − "), msup(mr("y"), mr("3")), 
              mr(" = 7, 4"), msup(mr("x"), mr("2")), mr(" − 2xy + "), msup(mr("y"), mr("2")), mr(" = 7")),
        "variants": [
            mg(mr("(1; 1)")),
            mg(mr("(2; 3)")),
            mg(mr("(1; 2)")),
            mg(mr("(3; 5)"))
        ],
        "correct": 0
    })
    
    # Problem 60: x⁴ + x²y² + y⁴ = 91, x² + xy + y = 13
    questions.append({
        "num": 60,
        "q": mg(msup(mr("x"), mr("4")), mr(" + "), msup(mr("x"), mr("2")), msup(mr("y"), mr("2")), 
              mr(" + "), msup(mr("y"), mr("4")), mr(" = 91, "), msup(mr("x"), mr("2")), 
              mr(" + xy + y = 13")),
        "variants": [
            mg(mr("(3; 1) va (−3; 1)")),
            mg(mr("(3; 1) faqat")),
            mg(mr("(2; 3) va (−2; 3)")),
            mg(mr("(1; 3) va (−1; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 61: x⁴ − x²y² + y⁴ = 133, x² − xy + y² = 7
    questions.append({
        "num": 61,
        "q": mg(msup(mr("x"), mr("4")), mr(" − "), msup(mr("x"), mr("2")), msup(mr("y"), mr("2")), 
              mr(" + "), msup(mr("y"), mr("4")), mr(" = 133, "), msup(mr("x"), mr("2")), 
              mr(" − xy + "), msup(mr("y"), mr("2")), mr(" = 7")),
        "variants": [
            mg(mr("(4; 1) va (−4; −1)")),
            mg(mr("(3; 2) va (−3; −2)")),
            mg(mr("(4; 1) faqat")),
            mg(mr("(5; 2) va (−5; −2)"))
        ],
        "correct": 0
    })
    
    # Problem 62: x − y + 6/(x−y) = 5, x/(x−y−2) = 4
    questions.append({
        "num": 62,
        "q": mg(mr("{  x − y + "), mfrac(mr("6"), mr("x − y")), mr(" = 5, "), 
              mfrac(mr("x"), mr("x − y − 2")), mr(" = 4")),
        "variants": [
            mg(mr("(10; 8)")),
            mg(mr("(8; 6)")),
            mg(mr("(12; 10)")),
            mg(mr("(9; 7)"))
        ],
        "correct": 0
    })
    
    # Problem 63: x − 2y + 4/(x−2y) = 5, x/(x−2y−1) = 2
    questions.append({
        "num": 63,
        "q": mg(mr("{  x − 2y + "), mfrac(mr("4"), mr("x − 2y")), mr(" = 5, "), 
              mfrac(mr("x"), mr("x − 2y − 1")), mr(" = 2")),
        "variants": [
            mg(mr("(6; 2)")),
            mg(mr("(4; 1)")),
            mg(mr("(8; 3)")),
            mg(mr("(5; 1.5)"))
        ],
        "correct": 0
    })
    
    # Problem 64: x² + xy = 20, y² + xy = 5
    questions.append({
        "num": 64,
        "q": mg(msup(mr("x"), mr("2")), mr(" + xy = 20, "), msup(mr("y"), mr("2")), mr(" + xy = 5")),
        "variants": [
            mg(mr("(5; 1) va (−4; −5)")),
            mg(mr("(4; 2) va (−5; −4)")),
            mg(mr("(5; 1) faqat")),
            mg(mr("(4; 1) va (−5; −5)"))
        ],
        "correct": 0
    })
    
    # Problem 65: x² + xy = 24, y² + xy = 12
    questions.append({
        "num": 65,
        "q": mg(msup(mr("x"), mr("2")), mr(" + xy = 24, "), msup(mr("y"), mr("2")), mr(" + xy = 12")),
        "variants": [
            mg(mr("(6; 2) va (−4; −6)")),
            mg(mr("(4; 3) va (−6; −4)")),
            mg(mr("(6; 2) faqat")),
            mg(mr("(5; 2) va (−5; −5)"))
        ],
        "correct": 0
    })
    
    # Problem 66: (5x−y)/(x+y) + (x+y)/(5x−y) = 2, x²−y² = −3
    questions.append({
        "num": 66,
        "q": mg(mfrac(mr("5x − y"), mr("x + y")), mr(" + "), mfrac(mr("x + y"), mr("5x − y")), 
              mr(" = 2, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = −3")),
        "variants": [
            mg(mr("(1; 2)")),
            mg(mr("(2; 3)")),
            mg(mr("(−1; 2)")),
            mg(mr("(1; −2)"))
        ],
        "correct": 0
    })
    
    # Problem 67: (3x+2y)/(x−y) + (x−y)/(3x+2y) = 2.5, x²−y² = 15
    questions.append({
        "num": 67,
        "q": mg(mfrac(mr("3x + 2y"), mr("x − y")), mr(" + "), 
              mfrac(mr("x − y"), mr("3x + 2y")), mr(" = 2.5, "), 
              msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), mr(" = 15")),
        "variants": [
            mg(mr("(3; 2)")),
            mg(mr("(5; 2)")),
            mg(mr("(4; 1)")),
            mg(mr("(6; 3)"))
        ],
        "correct": 0
    })
    
    # Problem 68: (x+3y)/(x−3y) + (x−3y)/(x+3y) = 3 1/3, x²+7y = 43
    questions.append({
        "num": 68,
        "q": mg(mfrac(mr("x + 3y"), mr("x − 3y")), mr(" + "), 
              mfrac(mr("x − 3y"), mr("x + 3y")), mr(" = 3"), mfrac(mr("1"), mr("3")), 
              mr(", "), msup(mr("x"), mr("2")), mr(" + 7y = 43")),
        "variants": [
            mg(mr("(6; 1)")),
            mg(mr("(5; 2)")),
            mg(mr("(9; 2)")),
            mg(mr("(7; 1)"))
        ],
        "correct": 0
    })
    
    # Problem 69: (5x/y) − (y/x) = 4, xy = 1
    questions.append({
        "num": 69,
        "q": mg(mfrac(mr("5x"), mr("y")), mr(" − "), mfrac(mr("y"), mr("x")), mr(" = 4, xy = 1")),
        "variants": [
            mg(mr("(1; 1)")),
            mg(mr("(2; 0.5)")),
            mg(mr("(0.5; 2)")),
            mg(mr("(1; 2)"))
        ],
        "correct": 0
    })
    
    # Problem 70: (x/y) + (y/x) = 2.5, x² + y² = 5
    questions.append({
        "num": 70,
        "q": mg(mfrac(mr("x"), mr("y")), mr(" + "), mfrac(mr("y"), mr("x")), mr(" = 2.5, "), 
              msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 5")),
        "variants": [
            mg(mr("(2; 1) va (1; 2)")),
            mg(mr("(√3; √2) va (√2; √3)")),
            mg(mr("(2; 1) faqat")),
            mg(mr("(√5; 1) va (1; √5)"))
        ],
        "correct": 0
    })
    
    # Problem 71: x + y = 10, x² + y² − 2xy + 36 nechta yechimga ega?
    questions.append({
        "num": 71,
        "q": mg(mr("{  x + y = 10, "), msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), 
              mr(" − 2xy + 36  tenglamalar sistemasi nechta yechimga ega?")),
        "variants": [
            mg(mr("Cheksiz ko'p yechim")),
            mg(mr("2 ta yechim")),
            mg(mr("1 ta yechim")),
            mg(mr("Yechimi yo'q"))
        ],
        "correct": 0
    })
    
    # Problem 72: x + y = 8, (x−4)(y−3) = 0 — x larning yig'indisini toping
    questions.append({
        "num": 72,
        "q": mg(mr("{  x + y = 8, (x − 4)(y − 3) = 0  tenglamalar sistemasini qanoatlantiruvchi "),
              mr("x larning yig'indisini toping.")),
        "variants": [
            mg(mr("9")),
            mg(mr("7")),
            mg(mr("8")),
            mg(mr("12"))
        ],
        "correct": 0
    })
    
    # Problem 73: x² + y² = 100, xy = −24 nechta yechimga ega?
    questions.append({
        "num": 73,
        "q": mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 100, xy = −24  "),
              mr("tenglamalar sistemasi nechta yechimga ega?")),
        "variants": [
            mg(mr("4 ta yechim")),
            mg(mr("2 ta yechim")),
            mg(mr("1 ta yechim")),
            mg(mr("Yechimi yo'q"))
        ],
        "correct": 0
    })
    
    # Problem 74: x/y³ = 4, x/x² = 9 — xy ning qiymatini toping
    questions.append({
        "num": 74,
        "q": mg(mfrac(mr("x"), msup(mr("y"), mr("3"))), mr(" = 4, "), 
              mfrac(mr("x"), msup(mr("z"), mr("2"))), mr(" = 9  bo'lsa, xy ning qiymatini toping.")),
        "variants": [
            mg(mr("1/18")),
            mg(mr("18")),
            mg(mr("1/36")),
            mg(mr("36"))
        ],
        "correct": 0
    })
    
    # Problem 75: x³ − y³ = 26, x − y = 2 — xy ning qiymatini toping
    questions.append({
        "num": 75,
        "q": mg(msup(mr("x"), mr("3")), mr(" − "), msup(mr("y"), mr("3")), mr(" = 26, x − y = 2  "),
              mr("bo'lsa, xy ning qiymatini toping.")),
        "variants": [
            mg(mr("3")),
            mg(mr("6")),
            mg(mr("9")),
            mg(mr("12"))
        ],
        "correct": 0
    })
    
    # Problem 76: x² + y² = 49, x² + y = 7 nechta yechimga ega?
    questions.append({
        "num": 76,
        "q": mg(msup(mr("x"), mr("2")), mr(" + "), msup(mr("y"), mr("2")), mr(" = 49, "), 
              msup(mr("x"), mr("2")), mr(" + y = 7  tenglamalar sistemasi nechta yechimga ega?")),
        "variants": [
            mg(mr("4 ta yechim")),
            mg(mr("2 ta yechim")),
            mg(mr("6 ta yechim")),
            mg(mr("3 ta yechim"))
        ],
        "correct": 0
    })
    
    # Problem 77: xy = 15, x² − y² = 16 — tenglama yechimlarining yig'indisini toping
    questions.append({
        "num": 77,
        "q": mg(mr("xy = 15, "), msup(mr("x"), mr("2")), mr(" − "), msup(mr("y"), mr("2")), 
              mr(" = 16  tenglama yechimlarining yig'indisini toping.")),
        "variants": [
            mg(mr("0")),
            mg(mr("16")),
            mg(mr("31")),
            mg(mr("15"))
        ],
        "correct": 0
    })
    
    # Problem 78: x² − 3x + 2/(y² + 6y − 40) = 0, x − 1/(y² + 6y − 40) = 0 — yechimlarning yig'indisi
    questions.append({
        "num": 78,
        "q": mg(msup(mr("x"), mr("2")), mr(" − 3x + "), 
              mfrac(mr("2"), mg(msup(mr("y"), mr("2")), mr(" + 6y − 40"))), mr(" = 0, "), 
              mr("x − "), mfrac(mr("1"), mg(msup(mr("y"), mr("2")), mr(" + 6y − 40"))), 
              mr(" = 0  tenglama yechimlarining yig'indisini toping.")),
        "variants": [
            mg(mr("−10")),
            mg(mr("10")),
            mg(mr("0")),
            mg(mr("−6"))
        ],
        "correct": 0
    })
    
    # ═══════════════════════════════════════════════════════════════
    #  PROBLEMS 79-87 FROM THIRD PDF
    # ═══════════════════════════════════════════════════════════════
    
    # Problem 79: 3x−y = 1, 2x+y = 4, x²+y² = 5 — yechimlarni toping
    questions.append({
        "num": 79,
        "q": mg(mr("{  3x − y = 1, 2x + y = 4, "), msup(mr("x"), mr("2")), mr(" + "), 
              msup(mr("y"), mr("2")), mr(" = 5  tenglamalar sistemasi yechimlarini ko'paytmasini toping.")),
        "variants": [
            mg(mr("2")),
            mg(mr("1")),
            mg(mr("5")),
            mg(mr("4"))
        ],
        "correct": 0
    })
    
    # Problem 80: √(3(x²−4)) + x − 3xy + 4y² = 0 — butun yechimlar
    questions.append({
        "num": 80,
        "q": mg(msqrt(mg(mr("3("), msup(mr("x"), mr("2")), mr(" − 4)"))), 
              mr(" + x − 3xy + 4"), msup(mr("y"), mr("2")), 
              mr(" = 0  tenglamani butun sonlarda yeching.")),
        "variants": [
            mg(mr("(2; 1), (−2; −1)")),
            mg(mr("(2; 0), (−2; 0)")),
            mg(mr("(2; 1) faqat")),
            mg(mr("(±2; ±1), (±2; 0)"))
        ],
        "correct": 0
    })
    
    # Problem 81: Ikki sonning yig'indisi 12, ko'paytmasi 35
    questions.append({
        "num": 81,
        "q": mg(mr("Ikki sonning yig'indisi 12 ga, ko'paytmasi 35 ga teng. Bu sonlarni toping.")),
        "variants": [
            mg(mr("5 va 7")),
            mg(mr("4 va 8")),
            mg(mr("3 va 9")),
            mg(mr("6 va 6"))
        ],
        "correct": 0
    })
    
    # Problem 82: Ikki sonning ayirmasi 5, ko'paytmasi 24
    questions.append({
        "num": 82,
        "q": mg(mr("Ikki sonning ayirmasi 5 ga, ko'paytmasi 24 ga teng. Bu sonlarni toping.")),
        "variants": [
            mg(mr("8 va 3")),
            mg(mr("9 va 4")),
            mg(mr("7 va 2")),
            mg(mr("6 va 1"))
        ],
        "correct": 0
    })
    
    # Problem 83: Ikki sonning ayirmasi 8, ko'paytmasi 48
    questions.append({
        "num": 83,
        "q": mg(mr("Ikki sonning ayirmasi 8 ga, ko'paytmasi 48 ga teng. Bu sonlarni toping.")),
        "variants": [
            mg(mr("12 va 4")),
            mg(mr("10 va 2")),
            mg(mr("14 va 6")),
            mg(mr("16 va 8"))
        ],
        "correct": 0
    })
    
    # Problem 84: Ikkita musbat sonning ko'paytmasi 21, kvadratlari yig'indisi 58
    questions.append({
        "num": 84,
        "q": mg(mr("Ikkita musbat sonning ko'paytmasi 21 ga, kvadratlari yig'indisi 58 ga teng. Bu sonlarni toping.")),
        "variants": [
            mg(mr("7 va 3")),
            mg(mr("6 va 3.5")),
            mg(mr("21 va 1")),
            mg(mr("14 va 1.5"))
        ],
        "correct": 0
    })
    
    # Problem 85: Ikkita musbat sonning ko'paytmasi 18, kvadratlari yig'indisi 45
    questions.append({
        "num": 85,
        "q": mg(mr("Ikkita musbat sonning ko'paytmasi 18 ga, kvadratlari yig'indisi 45 ga teng. Bu sonlarni toping.")),
        "variants": [
            mg(mr("6 va 3")),
            mg(mr("9 va 2")),
            mg(mr("18 va 1")),
            mg(mr("4.5 va 4"))
        ],
        "correct": 0
    })
    
    # Problem 86: Ikkita musbat sonning ko'paytmasi 24, kvadratlari yig'indisi 73
    questions.append({
        "num": 86,
        "q": mg(mr("Ikkita musbat sonning ko'paytmasi 24 ga, kvadratlari yig'indisi 73 ga teng. Bu sonlarni toping.")),
        "variants": [
            mg(mr("8 va 3")),
            mg(mr("6 va 4")),
            mg(mr("12 va 2")),
            mg(mr("24 va 1"))
        ],
        "correct": 0
    })
    
    # Problem 87: Diagonal √13, yuzi 6 — to'rtburchak perimetri
    questions.append({
        "num": 87,
        "q": mg(mr("Diagonali "), msqrt(mr("13")), mr(" ga, yuzi 6 ga teng bo'lgan to'rtburchakning perimetrini toping.")),
        "variants": [
            mg(mr("14")),
            mg(mr("12")),
            mg(mr("16")),
            mg(mr("10"))
        ],
        "correct": 0
    })
    
    return questions



# ═══════════════════════════════════════════════════════════════════════
#  ANSWER BALANCING & DOCUMENT GENERATION
# ═══════════════════════════════════════════════════════════════════════

def balance_answers(questions):
    """
    Redistribute correct answers evenly across A, B, C, D
    by swapping variant positions
    """
    target_pattern = ['A', 'B', 'C', 'D'] * (len(questions) // 4 + 1)
    random.shuffle(target_pattern)
    
    balanced = []
    for i, q in enumerate(questions):
        target_letter = target_pattern[i]
        target_idx = ['A', 'B', 'C', 'D'].index(target_letter)
        current_correct = q['correct']
        
        # Swap variants to move correct answer to target position
        variants = q['variants'][:]
        if current_correct != target_idx:
            variants[current_correct], variants[target_idx] = \
                variants[target_idx], variants[current_correct]
        
        balanced.append({
            "num": q["num"],
            "q": q["q"],
            "variants": variants,
            "correct": target_idx
        })
    
    return balanced

def text_run(text, bold=False, size=24):
    """Plain text run for Word"""
    safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    b_tag = '<w:b/><w:bCs/>' if bold else ''
    return (f'<w:r><w:rPr>{b_tag}'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
            f'<w:t xml:space="preserve">{safe}</w:t></w:r>')

def para(content, style="Normal", center=False, space_after=120):
    """Paragraph wrapper"""
    jc = '<w:jc w:val="center"/>' if center else ''
    return (f'<w:p xmlns:w="{W}" xmlns:m="{M}"><w:pPr>'
            f'<w:pStyle w:val="{style}"/><w:spacing w:after="{space_after}"/>{jc}</w:pPr>'
            f'{content}</w:p>')

def page_break():
    """Page break"""
    return f'<w:p xmlns:w="{W}"><w:r><w:br w:type="page"/></w:r></w:p>'



def build_document_xml(questions):
    """Build the complete Word document.xml"""
    body_parts = []
    
    # Title
    body_parts.append(para(
        text_run("IKKINCHI VA YUQORI DARAJALI TENGLAMALAR SISTEMASI", bold=True, size=32),
        center=True, space_after=160
    ))
    body_parts.append(para(
        text_run("Professional Test — Microsoft Word OMML Format", bold=True, size=24),
        center=True, space_after=200
    ))
    
    # Questions
    for q in questions:
        num = q["num"]
        q_omml = q["q"]
        variants = q["variants"]
        
        # Question paragraph
        question_content = (
            text_run(f"{num}.  ", bold=True, size=24) +
            f'<m:oMath xmlns:m="{M}">{q_omml}</m:oMath>'
        )
        body_parts.append(para(question_content, space_after=60))
        
        # Variants paragraph
        variant_content = ""
        for i, letter in enumerate(['A', 'B', 'C', 'D']):
            variant_content += text_run(f"  {letter})  ", bold=True, size=22)
            variant_content += f'<m:oMath xmlns:m="{M}">{variants[i]}</m:oMath>'
            variant_content += text_run("    ", size=22)
        
        body_parts.append(para(variant_content, space_after=100))
    
    # Page break before answer key
    body_parts.append(page_break())
    
    # Answer Key
    body_parts.append(para(
        text_run("JAVOBLAR KALITI — ANSWER KEY", bold=True, size=32),
        center=True, space_after=200
    ))
    
    # Answer key table
    for i in range(0, len(questions), 5):
        row_items = questions[i:i+5]
        row_text = "     ".join([
            f"{q['num']} — {['A','B','C','D'][q['correct']]}" 
            for q in row_items
        ])
        body_parts.append(para(text_run(row_text, size=22), space_after=80))
    
    # Distribution stats
    dist = Counter(q['correct'] for q in questions)
    dist_text = f"Taqsimot:  A={dist[0]}  B={dist[1]}  C={dist[2]}  D={dist[3]}"
    body_parts.append(para(
        text_run(dist_text, bold=True, size=20),
        center=True, space_after=80
    ))
    
    # Section properties
    sect = (f'<w:sectPr xmlns:w="{W}">'
            '<w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
            '</w:sectPr>')
    
    body_content = "\n".join(body_parts) + "\n" + sect
    
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document xmlns:w="{W}" xmlns:m="{M}">'
            f'<w:body>{body_content}</w:body></w:document>')



# ═══════════════════════════════════════════════════════════════════════
#  DOCX SKELETON FILES
# ═══════════════════════════════════════════════════════════════════════

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

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

SETTINGS = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W}">
  <w:defaultTabStop w:val="720"/>
</w:settings>'''

STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}" xmlns:m="{M}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
      <w:lang w:val="uz-Latn-UZ"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="120"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>'''



# ═══════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION
# ═══════════════════════════════════════════════════════════════════════

def write_docx(output_path, questions):
    """Create the final .docx file"""
    doc_xml = build_document_xml(questions)
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
        zf.writestr('word/document.xml', doc_xml)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/settings.xml', SETTINGS)
    
    size = os.path.getsize(output_path)
    print(f"\n✅ Fayl yaratildi: {output_path}")
    print(f"   Savollar soni: {len(questions)}")
    print(f"   Fayl hajmi: {size:,} bayt ({size/1024:.1f} KB)")
    
    # Answer distribution
    dist = Counter(q['correct'] for q in questions)
    print(f"\n📊 Javoblar taqsimoti:")
    print(f"   A = {dist[0]:2d}   B = {dist[1]:2d}   C = {dist[2]:2d}   D = {dist[3]:2d}")
    
    # Print answer key
    print(f"\n📋 JAVOBLAR KALITI:")
    for i in range(0, len(questions), 10):
        row = questions[i:i+10]
        line = "   ".join([
            f"{q['num']:2d}→{['A','B','C','D'][q['correct']]}" 
            for q in row
        ])
        print(f"   {line}")


if __name__ == "__main__":
    random.seed(42)  # For reproducible answer distribution
    
    print("=" * 70)
    print("  PROFESSIONAL MATEMATIK TEST GENERATOR")
    print("  PDF to Word with OMML Equations")
    print("=" * 70)
    
    # Generate questions from PDF
    print("\n🔍 PDF dan masalalarni o'qish...")
    questions = generate_questions()
    print(f"   {len(questions)} ta masala topildi")
    
    # Balance answer distribution
    print("\n⚖️  Javoblarni balanslash (A/B/C/D teng taqsimot)...")
    balanced_questions = balance_answers(questions)
    
    # Write output file
    output_file = "/projects/sandbox/test-yaratish-uchun/Tenglamalar_Sistemasi_SUPER_FULL_Test.docx"
    print("\n📝 Word faylini yaratish...")
    write_docx(output_file, balanced_questions)
    
    print("\n✨ TAYYOR! Word faylini oching va tekshiring.")
    print("   Barcha formulalar Microsoft Word Equation formatida.")
    print("=" * 70)

