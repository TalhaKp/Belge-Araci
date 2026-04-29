from core.helpers import DPI_SCALE, FONT_SCALE


def sc(value: int) -> int:
    """Verilen piksel değerini DPI ölçeğine göre ölçeklendirir."""
    return max(1, round(value * DPI_SCALE))


def sf(size: int) -> int:
    """Font boyutunu ekran boyutu ve DPI'a göre ölçeklendirir."""
    return max(8, round(size * FONT_SCALE))


STRINGS = {
    "tr": {
        "app_title":        "Belge Aracı",
        "select_tool":      "Ne yapmak istersiniz?",
        "btn_merge":        "📎  PDF Birleştir",
        "btn_word2pdf":     "📄  Word → PDF",
        "btn_back":         "← Geri",
        "btn_start":        "Başlat",
        "btn_browse":       "Klasör Seç",
        "label_folder":     "Klasör:",
        "label_outname":    "Çıktı dosya adı:",
        "label_delete":     "Orijinal Word dosyaları dönüşüm sonrası silinsin mi?",
        "label_delete_pptx":"Orijinal PowerPoint dosyaları dönüşüm sonrası silinsin mi?",
        "running":          "İşlem devam ediyor…",
        "done":             "✅  İşlem tamamlandı!",
        "err_no_folder":    "Lütfen bir klasör seçin.",
        "err_no_pdf":       "Seçilen klasörde hiç PDF yok.",
        "err_no_word":      "Seçilen klasörde hiç Word dosyası yok.",
        "err_word_app":     "Microsoft Word başlatılamadı. Word kurulu mu?",
        "merge_title":      "PDF Birleştirici",
        "pptx_merge_title": "PPTX Birleştirici",
        "word2pdf_title":   "Word → PDF Dönüştürücü",
        "merge_desc":       "Seçilen klasördeki tüm PDF'leri tek dosyada birleştirir.",
        "pptx_merge_desc":  "Seçilen klasördeki tüm PPTX dosyalarını tek sunumda birleştirir.",
        "label_outname_pptx": "Çıktı dosya adı:",
        "err_no_pptx_merge": "Seçilen klasörde hiç PPTX dosyası yok.",
        "err_pypptx":        "python-pptx eksik. Terminale: pip install python-pptx",
        "word2pdf_desc":    "Klasördeki tüm .doc/.docx dosyalarını PDF'e çevirir.",
        "log_scanning":     "Klasör taranıyor…",
        "log_merging":      "Birleştiriliyor: {name}",
        "log_converting":   "Dönüştürülüyor: {name}",
        "log_skipped":      "Es geçildi (zaten var): {name}",
        "log_ok":           "✓ {name}",
        "log_err":          "✗ {name} — {err}",
        "log_saved":        "Kaydedildi → {path}",
        "log_summary":      "Toplam: {ok} başarılı, {skip} atlandı.",
        "log_deleted":      "Silindi: {name}",
        "yes":              "Evet",
        "no":               "Hayır",
        "engine_word":      "Microsoft Word kullanılıyor…",
        "engine_ppt":       "Microsoft PowerPoint kullanılıyor…",
        "engine_libre":     "LibreOffice kullanılıyor…",
        "err_no_engine":    "Dönüştürücü bulunamadı.",
        "pptx2pdf_title":   "PPTX → PDF Dönüştürücü",
        "pptx2pdf_desc":    "Klasördeki tüm .ppt/.pptx dosyalarını PDF'e çevirir.",
        "err_no_pptx":      "Seçilen klasörde hiç PowerPoint dosyası yok.",
        "err_ppt_app":      "Microsoft PowerPoint başlatılamadı.",
        "libre_offer_title":"LibreOffice Bulunamadı",
        "libre_offer_msg":  (
            "Word → PDF dönüşümü için Microsoft Word veya LibreOffice gereklidir.\n\n"
            "Bilgisayarınızda ikisi de bulunamadı.\n\n"
            "LibreOffice ücretsizdir. İndirme sayfasını açmamı ister misiniz?"
        ),
    }
}


class LanguageManager:
    def __init__(self, lang="tr"):
        self.lang = lang

    def t(self, key, **kwargs):
        text = STRINGS.get(self.lang, STRINGS["tr"]).get(key, key)
        return text.format(**kwargs) if kwargs else text


LM = LanguageManager("tr")
t = LM.t

THEME = {
    "bg":           "#F5F0EB",
    "card":         "#FFFFFF",
    "accent":       "#C0392B",
    "accent_dark":  "#922B21",
    "text":         "#1A1A1A",
    "subtext":      "#6B6B6B",
    "border":       "#E0D8D0",
    "log_bg":       "#1E1E1E",
    "log_fg":       "#D4D4D4",
    "log_ok":       "#6A9F6A",
    "log_err":      "#E06C75",
    "log_info":     "#61AFEF",
    "btn_bg":       "#C0392B",
    "btn_fg":       "#FFFFFF",
    "btn_hover":    "#922B21",
    "font_main":    ("Segoe UI", sf(10)),
    "font_title":   ("Segoe UI", sf(18), "bold"),
    "font_sub":     ("Segoe UI", sf(9)),
    "font_btn":     ("Segoe UI", sf(10), "bold"),
    "font_mono":    ("Consolas", sf(9)),
}
