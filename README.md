# Belge Aracı 📎📄📊

Günlük belge işlerini kolaylaştırmak için yapılmış sade bir Windows masaüstü uygulaması. Tek bir arayüzden PDF ve PowerPoint işlemlerini, Word dönüşümlerini yapabilirsin. Klasörü seç, başlat — hepsi bu.

---

## Özellikler

- 📎 **PDF Birleştirici** — Seçilen klasördeki tüm PDF dosyalarını tek bir dosyada birleştirir. Dosyalar akıllı sıralama ile doğru sıraya dizilir (1, 2, 10 sıralaması 1, 10, 2 olmaz).
- 📑 **PPTX Birleştirici** — Klasördeki tüm PowerPoint dosyalarını tek bir sunumda birleştirir. Medya ve referanslar PowerPoint'in kendi motoru ile taşınır, bozulma olmaz.
- 📄 **Word → PDF Dönüştürücü** — Klasördeki tüm `.doc` ve `.docx` dosyalarını PDF'e çevirir. Zaten PDF'i olan dosyaları atlar. İstersen orijinal dosyaları dönüşüm sonrası otomatik silebilir.
- 📊 **PPTX → PDF Dönüştürücü** — Klasördeki tüm `.ppt` ve `.pptx` dosyalarını PDF'e çevirir. Aynı atlama ve silme mantığı geçerlidir.
- 🔍 **Klasör önizleme** — Klasör seçilince ilgili dosyalar anında listelenir. Klasörü Windows Explorer'da açmak için tek tıklama yeterli.
- 🔄 **Otomatik motor seçimi** — Word/PowerPoint → LibreOffice → indirme teklifi sırasıyla dener.
- 🖥️ **DPI ve ekran boyutuna duyarlı arayüz** — 4K'dan küçük laptop ekranlarına kadar düzgün görünür.

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

```bash
# Gerekli kütüphaneleri kur
pip install pypdf pywin32

# Uygulamayı başlat
python main.py
```

**`.exe` oluşturmak istersen:**

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=belge_araci.ico --add-data "belge_araci.ico;." --name "BelgeAraci" main.py
```

Oluşan dosya `dist/BelgeAraci.exe` konumunda olacaktır.

---

## Proje Yapısı

```
/
├── main.py              # Giriş noktası
├── belge_araci.ico      # Uygulama ikonu
├── README.md
├── core/
│   ├── helpers.py       # get_resource_path, DPI ölçekleme, sanitize_filename
│   └── config.py        # STRINGS, THEME, LanguageManager, sc(), sf()
├── gui/
│   ├── app.py           # Ana pencere ve ekran yönetimi
│   └── components.py    # HoverButton, form bileşenleri
└── tools/
    ├── tool_base.py     # ToolBase ABC, LibreOffice tespiti
    ├── pdf_merger.py    # PDF Birleştirici
    ├── pptx_merger.py   # PPTX Birleştirici
    ├── word2pdf.py      # Word → PDF
    └── pptx2pdf.py      # PPTX → PDF
```

---

## Changelog

### v1.4.0
- Proje modüler yapıya taşındı (`core/`, `gui/`, `tools/`)
- Klasör seçiminde dosya önizleme eklendi
- Klasörü Explorer'da açma butonu eklendi
- Font ölçekleme küçük ekranlar için düzeltildi

### v1.3.0
- PPTX Birleştirici eklendi (COM `InsertFromFile` ile, bozulma yok)
- PPTX → PDF Dönüştürücü eklendi

### v1.1.1
- Çıktı dosya adına `.pdf` uzantısı zorlandı
- Dosya adı injection güvenliği eklendi
- Pencere boyutu genişletildi

### v1.1.0
- Uygulama ikonu eklendi
- DPI ölçeklendirme eklendi

### v1.0.0
- İlk sürüm: PDF Birleştirici ve Word → PDF Dönüştürücü

---

## Geliştirici Notları

Yeni araç eklemek için `tools/tool_base.py` içindeki `ToolBase` sınıfından türeyen yeni bir sınıf yaz, `tools/__init__.py` içindeki `TOOLS` listesine ekle. Dil desteği için `core/config.py` içindeki `STRINGS` sözlüğüne yeni dil anahtarı ekle.
