# Reklama Havolalar Bazasi

## O'rnatish va ishga tushirish

### 1. Kerakli kutubxonalarni o'rnating
```bash
pip install -r requirements.txt
```

### 2. Dasturni ishga tushiring
```bash
python app.py
```

### 3. Brauzerda oching
```
http://10.170.115.120:5003/
```

---

## Fayllar tarkibi

```
link_tracker/
├── app.py              # Asosiy server (Flask)
├── requirements.txt    # Kerakli kutubxonalar
├── links.db            # Ma'lumotlar bazasi (avtomatik yaratiladi)
├── static/
│   └── index.html      # Veb interfeys
└── README.md
```

---

## Imkoniyatlar

- ✅ Havola qo'shish, tahrirlash, o'chirish
- ✅ Status belgilash: Faol / O'chgan
- ✅ Sana oralig'i bo'yicha filtrlash
- ✅ Like soni bo'yicha filtrlash
- ✅ Qidiruv (link bo'yicha)
- ✅ Excel (.xlsx) ga export
- ✅ Jami / Faol / O'chgan statistikasi (sarlavhada)

---

## Tarmoqda ulashish

Dastur `0.0.0.0:5002` da ishga tushadi, ya'ni LAN tarmog'idagi barcha hamkasblar
`http://10.170.115.120:5002/` orqali ulanib ishlata olishadi.
