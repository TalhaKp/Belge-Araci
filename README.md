# Belge Aracı 📎📄

Günlük belge işlerini kolaylaştırmak için yapılmış sade bir Windows masaüstü uygulaması. Tek bir arayüzden PDF birleştirme ve Word → PDF dönüştürme işlemlerini yapabilirsin. Klasörü seç, başlat — hepsi bu.

Uygulama Microsoft Word varsa onu, yoksa LibreOffice'i otomatik olarak kullanır. İkisi de yoksa LibreOffice indirme sayfasını açmayı teklif eder.

---

## Özellikler

- 📎 **PDF Birleştirici** — Seçilen klasördeki tüm PDF dosyalarını tek bir dosyada birleştirir. Dosyalar akıllı sıralama ile doğru sıraya dizilir (1, 2, 10 sıralaması 1, 10, 2 olmaz).
- 📄 **Word → PDF Dönüştürücü** — Klasördeki tüm `.doc` ve `.docx` dosyalarını PDF'e çevirir. Zaten PDF'i olan dosyaları atlar. İstersen orijinal Word dosyalarını dönüşüm sonrası otomatik silebilir.
- 🔄 **Otomatik motor seçimi** — Microsoft Word → LibreOffice → indirme teklifi sırasıyla dener.
- 🪟 **Sade arayüz** — Teknik bilgi gerektirmez.

---

## Kurulum ve Kullanım

### Hazır `.exe` ile (önerilen)

1. Sağ üstteki **Releases** sekmesine tıkla.
2. En son sürümün altındaki `BelgeAraci.exe` dosyasını indir.
3. Dosyaya çift tıkla, çalıştır. Python veya başka bir kurulum gerekmez.

> **Not:** Word → PDF özelliği için bilgisayarında **Microsoft Word** veya **[LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)** kurulu olmalıdır. PDF birleştirme için herhangi bir ek programa gerek yoktur.

---

### Kaynaktan çalıştırmak istersen

**Gereksinimler:**
- Python 3.10+
- Windows

**Adımlar:**

```bash
# Gerekli kütüphaneleri kur
pip install pypdf pywin32

# Uygulamayı başlat
python belge_araci.py
```

**`.exe` kendin oluşturmak istersen:**

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "BelgeAraci" belge_araci.py
```

Oluşan dosya `dist/BelgeAraci.exe` konumunda olacaktır.

---

## Dosya Yapısı

```
/
├── belge_araci.py   # Uygulamanın tüm kaynak kodu
└── README.md        # Bu dosya
```

---

## Geliştirici Notları

Yeni bir araç eklemek için `ToolBase` sınıfından türeyen yeni bir sınıf yaz ve dosyanın altındaki `TOOLS` listesine ekle — arayüz geri kalanını otomatik halleder. Dil desteği eklemek için `STRINGS` sözlüğüne yeni bir dil anahtarı eklemen yeterli.
