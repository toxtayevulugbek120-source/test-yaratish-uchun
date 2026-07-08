# Ratsional Tengsizliklar Test Generator

Professional matematik test generatori — Ratsional tengsizliklar uchun Microsoft Word formatida.

## 📋 Tavsif

Ushbu dastur matematik tengsizliklardan professional A/B/C/D variantli testlar yaratadi. Barcha matematik ifodalar **Microsoft Word Equation (OMML)** formatida yozilgan.

## ✨ Xususiyatlar

- ✅ **88 ta professional savol** — Ratsional tengsizliklar bo'yicha (2 PDF dan)
- ✅ **OMML formatdagi matematik ifodalar** — Word'da to'liq tahrir qilinadigan
- ✅ **Muvozanatli javoblar taqsimoti** — A, B, C, D har biri 22 martadan
- ✅ **Interval yozuvi** — Professional matematika kitoblari uslubida
- ✅ **Javoblar kaliti** — Jadval ko'rinishida, statistika bilan
- ✅ **Tayyor nashrga** — Print-ready formatda

## 📁 Fayllar

- `ratsional_tengsizliklar_test.py` — Test generatori (Python dasturi)
- `Ratsional_Tengsizliklar_FULL_Test.docx` — **TO'LIQ TEST (88 savol)** ⭐
- `Ratsional_Tengsizliklar_Test_56q.docx` — Birinchi versiya (56 savol)
- `Ratsional tengsizlik 1.pdf` — Asl manba 1 (PDF format)
- `Ratsional tengsizliklar 2.pdf` — Asl manba 2 (PDF format)

## 🚀 Ishlatish

### Tayyorlangan testni ochish

```bash
# TAVSIYA: To'liq versiya (88 savol)
Ratsional_Tengsizliklar_FULL_Test.docx

# Yoki birinchi versiya (56 savol)
Ratsional_Tengsizliklar_Test_56q.docx
```

### Yangi test yaratish

```bash
# Python 3 kerak
python3 ratsional_tengsizliklar_test.py
```

Natija: `Ratsional_Tengsizliklar_Test.docx` fayli yaratiladi.

## 📊 Test Tuzilishi (88 ta savol)

### Savollar taqsimoti

1. **1-14**: Oddiy ratsional tengsizliklar
2. **15-30**: Kvadratik va ko'p hadli tengsizliklar  
3. **31-45**: Murakkab kasrli tengsizliklar
4. **46-56**: Maxsus holatlar (diskriminant bilan)
5. **57-75**: Nisbatlar va aralash kasrlar
6. **76-85**: Murakkab kasrli sistemalar
7. **86-88**: Iqtidorli o'quvchilar uchun (yuqori daraja)

### Javoblar formati

Barcha javoblar quyidagi ko'rinishlarda:

- **Intervallar**: `(−∞ ; 3]`, `[2 ; 5)`, `(−1 ; 4)`
- **Birlashmalar**: `(−∞ ; −2] ∪ (1 ; +∞)`
- **Maxsus to'plamlar**: `ℝ`, `∅`

## 🎯 Javoblar Statistikasi

### To'liq Test (88 savol)

```
A = 22 savol (25%)
B = 22 savol (25%)
C = 22 savol (25%)
D = 22 savol (25%)
```

**Mukammal muvozanatli taqsimot!**

## 💡 Misollar

### Oddiy Ratsional Tengsizlik

**Savol 1:**
```
x/(x−1) ≤ 0
```

**Variantlar:**
```
A) (−∞ ; 0] ∪ (1 ; +∞)
B) [0 ; 1)                    ← To'g'ri javob
C) (−∞ ; 0) ∪ (1 ; +∞)
D) [0 ; 1]
```

### Murakkab Kasrli Tengsizlik

**Savol 85:**
```
(2x²−14x+6)/(x²−4x+3) ≥ (3x−8)/(x−3)
```

**Variantlar:**
```
A) (−∞ ; 0] ∪ (1 ; 2] ∪ (3 ; +∞)    ← To'g'ri javob
B) [0 ; 1) ∪ [2 ; 3)
C) (−∞ ; 0) ∪ [1 ; 2) ∪ [3 ; +∞)
D) [0 ; 2] ∪ (3 ; +∞)
```

### Yuqori Daraja (Iqtidorlilar uchun)

**Savol 86:**
```
x⁵ + x³ + x ≥ 138
```

**Variantlar:**
```
A) [3 ; +∞)                   ← To'g'ri javob
B) (−∞ ; 3]
C) (3 ; +∞)
D) ℝ
```

## 🛠️ Texnik Tafsilotlar

### Format

- **Fayl turi**: Microsoft Word (.docx)
- **Matematik ifodalar**: OMML (Office Math Markup Language)
- **Sahifa o'lchami**: A4 (21cm × 29.7cm)
- **Shrift**: Times New Roman, 24pt (asosiy), 32pt (sarlavha)

### Python Dependencies

Faqat standart kutubxonalar:
- `zipfile` — .docx yaratish uchun
- `collections.Counter` — statistika uchun
- `random` — javoblarni aralashtirish uchun

**Qo'shimcha kutubxona o'rnatish talab qilinmaydi!**

## 📝 Kodda O'zgartirish

### Yangi savol qo'shish

`ratsional_tengsizliklar_test.py` faylidagi `QUESTIONS` ro'yxatiga qo'shing:

```python
{"num": 57,
 "q": mg(mf("x+5", "x−3"), mr(" > 0")),
 "A": mr("(−∞ ; −5) ∪ (3 ; +∞)"),
 "B": mr("(−5 ; 3)"),
 "C": mr("[−5 ; 3]"),
 "D": mr("(−∞ ; −5] ∪ [3 ; +∞)"),
 "ans": "A"},
```

### OMML Funksiyalari

- `mr(text)` — Oddiy matematik matn
- `mf(num, den)` — Kasr (numerator/denominator)
- `msup(base, sup)` — Daraja (superscript)
- `mg(*parts)` — Bir nechta elementni birlashtirish

## 📚 Manbalar

Ushbu test quyidagi fayllardan olingan:

1. **[Ratsional tengsizlik 1.pdf](Ratsional%20tengsizlik%201.pdf)** — 1-56 masalalar
2. **[Ratsional tengsizliklar 2.pdf](Ratsional%20tengsizliklar%202.pdf)** — 57-88 masalalar (qo'shimcha)

## 📈 Versiyalar Tarixi

| Versiya | Sana | Savollar soni | Tavsif |
|---------|------|---------------|--------|
| **2.0** | 2026-07-08 | **88** | To'liq versiya (2 PDF birlashtirilgan) ⭐ |
| 1.0 | 2026-07-08 | 56 | Birinchi nashr |

## 👨‍💻 Muallif

**Toʻxtayev Ulugʻbek**  
Matematika bo'yicha professional test generatori

## 📄 Litsenziya

Ushbu loyiha ochiq kodli va ta'lim maqsadlarida erkin foydalanish uchun.

---

## 🔗 Bog'lanish

- GitHub: [@toxtayevulugbek120-source](https://github.com/toxtayevulugbek120-source)
- Repository: [test-yaratish-uchun](https://github.com/toxtayevulugbek120-source/test-yaratish-uchun)

---

**Oxirgi yangilanish**: 2026-yil 8-iyul  
**Versiya**: 2.0 (88 savol — TO'LIQ) ⭐  
**Birinchi nashr**: 1.0 (56 savol)
