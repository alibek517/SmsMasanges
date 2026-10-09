# 👥 Odamlar Ro'yxati Dashboard (Supabase + Python Streamlit)

Ushbu loyiha Supabase bulutli ma'lumotlar bazasi bilan integratsiya qilingan, zamonaviy va qulay dashboard hisoblanadi. Unda har bir shaxsning to'liq ma'lumotlari saqlanadi, qidiriladi va tahrirlanadi.

---

## 📋 Shaxs Ma'lumotlari Maydonlari (8 ta asosiy parametr)

Har bir shaxs (fuqaro) uchun quyidagi maydonlar mavjud:

1. **`DistrictName`** — Tuman / Shahar nomi (Masalan: *Chilonzor tumani*)
2. **`LocalityName`** — Aholi punkti / Mahalla (Masalan: *Do'mbirobod MFY*)
3. **`PinFL`** — JSHSHIR (14 xonali unikal shaxsiy raqam)
4. **`inFL`** — INN / STIR yoki boshqa identifikator
5. **`Yoshi`** — Shaxsning yoshi (Masalan: *29*)
6. **`FullName`** — To'liq F.I.Sh (Familiya Ism Sharif)
7. **`PhoneNumberi`** — Telefon raqami (Masalan: *+998901234567*)
8. **`Documenti`** — Pasport yoki ID-karta seriya va raqami (Masalan: *AA1234567*)

---

## 🚀 Ishga Tushirish

### 1. Talab qilinadigan kutubxonalarni o'rnatish:
```bash
pip install -r requirements.txt
```

### 2. Dashboardni ishga tushirish:
```bash
streamlit run app.py
```
Buyruq bajarilgach, brauzeringizda avtomatik ravishda `http://localhost:8501` ochiladi.

---

## 🗄️ Supabase Bazasini Sozlash (3 Qadam)

1. [supabase.com](https://supabase.com) saytiga kiring va yangi loyiha yarating.
2. Chap menyudan **SQL Editor** bo'limiga kiring va loyihangizdagi `supabase_schema.sql` fayli ichidagi barcha kodni nusxalab, **RUN** tugmasini bosing.
   - Bu buyruq `people` jadvalini, tezkor qidiruv indekslarini va xavfsizlik qoidalarini (RLS) avtomatik yaratadi.
3. Chap menyudan **Project Settings -> API** bo'limiga kiring:
   - **Project URL** va **anon public API Key** ni oling.
   - Loyihadagi `.env` fayliga kiriting:
     ```env
     SUPABASE_URL=https://your-project.supabase.co
     SUPABASE_KEY=eyJhbGciOi...
     ```
   - *Yoki ushbu kalitlarni to'g'ridan-to'g'ri Dashboard saydbari (chap paneli) orqali kiritishingiz ham mumkin!*

---

## 🌟 Asosiy Imkoniyatlar

- **Jonli Qidiruv:** F.I.Sh, PINFL, Telefon, Hujjat bo'yicha lahzalik qidiruv.
- **Filtrlar:** Tuman, Mahalla va Yosh oralig'i bo'yicha saralash.
- **Yangi Shaxs Qo'shish:** 14 ta raqamli PINFL va maydonlar validatsiyasi bilan oson kiritish.
- **Tahrirlash va O'chirish:** Har bir shaxs ma'lumotlarini to'g'ridan-to'g'ri o'zgartirish yoki o'chirish.
- **Eksport:** Barcha yoki filtrlangan odamlar ro'yxatini 1 ta tugma orqali **CSV** (Excel) formatida yuklab olish.
- **Statistika:** Tumanlar va yosh guruhlari bo'yicha avtomatik diagrammalar.
- **Offline / Sinov rejimi:** Supabase kalitlari ulanmagan taqdirda ham dastur to'xtab qolmaydi, qulay sinov ma'lumotlari bilan ishlaydi.
