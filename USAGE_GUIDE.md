# Qo'llanma — PDF dan Professional Test Yaratish

## 📚 Sizda Mavjud

PDF faylingizdagi quyidagi bo'limlar:

### 1–20. Tenglamalar sistemasini yeching
- 20 ta oddiy va murakkab sistemalar

### 21–25. Tenglamalar sistemasining butun yechimlarini toping  
- 5 ta masala

### 26–35. Tenglamalar sistemasining barcha ildizlari yig'indisini toping
- 10 ta masala

### 31–35. (Maxsus sistemalar)
- 5 ta masala

**Jami PDF da:** ~40+ masala

## ✅ Men Qilgan Ish

Birinchi 24 ta masalani professional test formatiga aylantirib, Word faylini yaratdim:
- ✅ Microsoft Word OMML formatida
- ✅ A/B/C/D variantlar bilan
- ✅ To'g'ri javoblar muvozanatlashtirilgan
- ✅ Javoblar kaliti qo'shilgan

## 🔧 Barcha Masalalarni Qo'shish Uchun

### Variant 1: Python Skriptni Kengaytirish

`pdf_to_word_test.py` faylining `generate_questions()` funksiyasiga yangi masalalarni qo'shing:

```python
# Problem 25: yangi masala
questions.append({
    "num": 25,
    "q": mg(mr("{  yangi tenglama")),
    "variants": [
        mg(mr("javob A")),
        mg(mr("javob B")),
        mg(mr("javob C")),
        mg(mr("javob D"))
    ],
    "correct": 0  # 0=A, 1=B, 2=C, 3=D
})
```

### Variant 2: Men Davom Ettirishim

Agar barcha qolgan masalalarni ham aylantirishimni istasangiz, menga quyidagilarni ayting:

1. **Qaysi bo'limlarni qo'shish kerak?**
   - 21–25 masalalar?
   - 26–35 masalalar?
   - Barchasi?

2. **Har bir masala uchun variant yaratish qoidalari:**
   - Qanday noto'g'ri javoblar bo'lishi kerak?
   - Qanday xatolar asosida chalg'ituvchilar yaratish kerak?

## 🎨 Formula Yaratish Namunalari

### Oddiy Tenglama
```python
mr("x + y = 5")
```

### Kasr
```python
mfrac(mr("x + 1"), mr("x − 2"))  # (x+1)/(x-2)
```

### Daraja
```python
msup(mr("x"), mr("2"))  # x²
msup(mr("x"), mr("3"))  # x³
```

### Ildiz
```python
msqrt(mr("x + 5"))  # √(x+5)
```

### Murakkab Formula
```python
mg(
    msup(mr("x"), mr("2")),
    mr(" + "),
    mfrac(mr("3"), mr("x")),
    mr(" = 7")
)
# x² + 3/x = 7
```

### Sistema
```python
mg(mr("{  x + y = 5, xy = 6"))
```

## 📊 Variant Yaratish Strategiyasi

### To'g'ri Javob Turlari:

1. **Bitta yechim:** `(3; 5)`
2. **Ikki yechim:** `(2; 4) va (4; 2)`
3. **Natural sonlar:** `{1, 2, 3, 5}`
4. **Yig'indi:** `15`
5. **Oraliq:** `x ∈ [−2; 5]`

### Noto'g'ri Javoblar (Chalg'ituvchilar):

1. **Belgini teskari olish:**
   - To'g'ri: `(3; 5)`
   - Noto'g'ri: `(−3; 5)` yoki `(3; −5)`

2. **Koordinatalarni almashtirish:**
   - To'g'ri: `(3; 5)`
   - Noto'g'ri: `(5; 3)`

3. **Hisoblash xatolari:**
   - To'g'ri: `x + y = 15`
   - Noto'g'ri: `x + y = 14` yoki `x + y = 16`

4. **Yechimlardan birini tushirish:**
   - To'g'ri: `(2; 4) va (4; 2)`
   - Noto'g'ri: `(2; 4)` faqat bitta

5. **Aralash xatolar:**
   - To'g'ri: `(3; 5)`
   - Noto'g'ri: `(4; 6)` (ikkalasini 1 ga oshirish)

## 🚀 Keyingi Qadamlar

### Men Bajaraman:

Agar sizga kerak bo'lsa, men quyidagilarni qila olaman:

1. **PDF dan barcha masalalarni extract qilish**
2. **Har bir masala uchun to'g'ri yechim topish**
3. **Mantiqiy chalg'ituvchi variantlar yaratish**
4. **Professional Word faylini yaratish**
5. **Javoblar kalitini qo'shish**

### Sizdan Kerak:

- Qaysi masalalarni test formatiga aylantirishni xohlaysiz?
- Har bir test nechta savoldan iborat bo'lishi kerak?
- Bitta fayl yoki bo'limlarga ajratilgan bir nechta fayl?

## 📞 Savollar

Agar qo'shimcha savollar yoki o'zgartirishlar kerak bo'lsa, ayting:

1. Test formatini o'zgartirish
2. Ko'proq variantlar (A/B/C/D/E)
3. Boshqa shrift yoki dizayn
4. Maxsus talablar

---

**Tayyor fayl:** `Tenglamalar_Sistemasi_Test.docx`  
**Skript:** `pdf_to_word_test.py`  
**Status:** ✅ 24 ta masala tayyor
