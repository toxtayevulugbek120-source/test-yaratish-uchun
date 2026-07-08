# Ratsional Tengsizliklar Test Generator

Professional matematik test generatori — Ratsional tengsizliklar uchun Microsoft Word formatida.

## 📋 Tavsif

Ushbu dastur matematik tengsizliklardan professional A/B/C/D variantli testlar yaratadi. Barcha matematik ifodalar **Microsoft Word Equation (OMML)** formatida yozilgan.

## ✨ Xususiyatlar

- ✅ **56 ta professional savol** — Ratsional tengsizliklar bo'yicha
- ✅ **OMML formatdagi matematik ifodalar** — Word'da to'liq tahrir qilinadigan
- ✅ **Muvozanatli javoblar taqsimoti** — A, B, C, D har biri 14 martadan
- ✅ **Interval yozuvi** — Professional matematika kitoblari uslubida
- ✅ **Javoblar kaliti** — Jadval ko'rinishida, statistika bilan
- ✅ **Tayyor nashrga** — Print-ready formatda

## 📁 Fayllar

- `ratsional_tengsizliklar_test.py` — Test generatori (Python dasturi)
- `Ratsional_Tengsizliklar_Test_56q.docx` — Tayyor test (56 savol)
- `Ratsional tengsizlik 1.pdf` — Asl manba (PDF format)

## 🚀 Ishlatish

### Tayyorlangan testni ochish

```bash
# Microsoft Word yoki LibreOffice'da oching
Ratsional_Tengsizliklar_Test_56q.docx
```

### Yangi test yaratish

```bash
# Python 3 kerak
python3 ratsional_tengsizliklar_test.py
```

Natija: `Ratsional_Tengsizliklar_Test.docx` fayli yaratiladi.

## 📊 Test Tuzilishi

### Savollar taqsimoti (56 ta)

1. **1-14**: Oddiy ratsional tengsizliklar
2. **15-30**: Kvadratik va ko'p hadli tengsizliklar  
3. **31-45**: Murakkab kasrli tengsizliklar
4. **46-56**: Maxsus holatlar (diskriminant bilan)

### Javoblar formati

Barcha javoblar quyidagi ko'rinishlarda:

- **Intervallar**: `(−∞ ; 3]`, `[2 ; 5)`, `(−1 ; 4)`
- **Birlashmalar**: `(−∞ ; −2] ∪ (1 ; +∞)`
- **Maxsus to'plamlar**: `ℝ`, `∅`

## 🎯 Javoblar Statistikasi

```
A = 14 savol (25%)
B = 14 savol (25%)
C = 14 savol (25%)
D = 14 savol (25%)
```

**To'liq muvozanatli taqsimot!**

## 💡 Misol

**Savol:**
```
1.  x/(x−1) ≤ 0
```

**Variantlar:**
```
A) (−∞ ; 0] ∪ (1 ; +∞)
B) [0 ; 1)                    ← To'g'ri javob
C) (−∞ ; 0) ∪ (1 ; +∞)
D) [0 ; 1]
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

## 📚 Manba

Ushbu test [Ratsional tengsizlik 1.pdf](https://github.com/toxtayevulugbek120-source/test-yaratish-uchun) faylidan olingan.

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
**Versiya**: 1.0 (56 savol)
