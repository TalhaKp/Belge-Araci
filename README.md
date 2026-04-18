# Belge Aracı 📎📄📊

Günlük belge işlerini kolaylaştırmak için yapılmış sade bir Windows masaüstü uygulaması. Tek bir arayüzden PDF ve PowerPoint işlemlerini, Word dönüşümlerini yapabilirsin. Klasörü seç, başlat — hepsi bu.

---

## Özellikler

- 📎 **PDF Birleştirici** — Seçilen klasördeki tüm PDF dosyalarını tek bir dosyada birleştirir. Dosyalar akıllı sıralama ile doğru sıraya dizilir (1, 2, 10 sıralaması 1, 10, 2 olmaz).
- 📑 **PPTX Birleştirici** — Seçilen klasördeki tüm PowerPoint dosyalarını tek bir sunumda birleştirir. Aynı akıllı sıralama uygulanır.
- 📄 **Word → PDF Dönüştürücü** — Klasördeki tüm `.doc` ve `.docx` dosyalarını PDF'e çevirir. Zaten PDF'i olan dosyaları atlar. İstersen orijinal Word dosyalarını dönüşüm sonrası otomatik silebilir.
- 📊 **PPTX → PDF Dönüştürücü** — Klasördeki tüm `.ppt` ve `.pptx` dosyalarını PDF'e çevirir. Aynı atlama ve silme mantığı geçerlidir.
- 🔄 **Otomatik motor seçimi** — Word/PowerPoint → LibreOffice → indirme teklifi sırasıyla dener.
- 🖥️ **DPI ve ekran boyutuna duyarlı arayüz** — 4K'dan küçük laptop ekranlarına kadar düzgün görünür.
- 🪟 **Sade arayüz** — Teknik bilgi gerektirmez.

---

## Kurulum ve Kullanım

### Hazır `.exe` ile (önerilen)

1. Sağ üstteki **Releases** sekmesine tıkla.
2. En son sürümün altındaki `BelgeAraci.exe` dosyasını indir.
3. Dosyaya çift tıkla, çalıştır. Python veya başka bir kurulum gerekmez.

> **Not:** Word/PowerPoint → PDF özellikleri için bilgisayarında **Microsoft Office** veya **[LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)** kurulu olmalıdır. PDF ve PPTX birleştirme için herhangi bir ek programa gerek yoktur.

> **Windows Smart App Control:** İmzasız `.exe` dosyalarını ilk açışta Windows engelleyebilir. Ayarlar → Gizlilik ve Güvenlik → Windows Güvenliği → Uygulama ve Tarayıcı Denetimi → Akıllı Uygulama Denetimi'ni geçici olarak kapatarak çalıştırabilirsin.

---

### Kaynaktan çalıştırmak istersen

**Gereksinimler:**
- Python 3.10+
- Windows

**Adımlar:**

```bash
# Gerekli kütüphaneleri kur
pip install pypdf pywin32 python-pptx

# Uygulamayı başlat
python belge_araci.py
```

**`.exe` kendin oluşturmak istersen:**

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=belge_araci.ico --name "BelgeAraci" belge_araci.py
```

Oluşan dosya `dist/BelgeAraci.exe` konumunda olacaktır.

---

## Dosya Yapısı

```
/
├── belge_araci.py   # Uygulamanın tüm kaynak kodu
├── belge_araci.ico  # Uygulama ikonu
└── README.md        # Bu dosya
```

---

## Changelog

### v1.3.0
- PPTX Birleştirici eklendi
- PPTX → PDF Dönüştürücü eklendi

### v1.1.1
- Küçük ekranlarda font ölçekleme düzeltildi
- Pencere boyutu genişletildi
- PPTX dönüştürücüde silme sorusu düzeltildi

### v1.1.0
- Uygulama ikonu eklendi
- DPI ölçeklendirme eklendi

### v1.0.0
- İlk sürüm: PDF Birleştirici ve Word → PDF Dönüştürücü

---

## Geliştirici Notları

Yeni bir araç eklemek için `ToolBase` sınıfından türeyen yeni bir sınıf yaz ve dosyanın altındaki `TOOLS` listesine ekle — arayüz geri kalanını otomatik halleder. Dil desteği eklemek için `STRINGS` sözlüğüne yeni bir dil anahtarı eklemen yeterli.
