# Yazgaç

Word (.docx) ve Excel (.xlsx / .xls / .csv) dosyalarını açıp düzenleyebilen, kaydedebilen basit bir masaüstü uygulama.

## Başlatma

- **Kolay yol:** `baslat.bat` dosyasına çift tıklayın.
- **Komut satırı:** proje klasöründe `npm start`

İlk çalıştırmada `node_modules` yoksa `npm install` otomatik çalışır (internet gerekir).

## Özellikler

**Word**
- .docx aç / kaydet / farklı kaydet
- Kalın, italik, altı çizili, üstü çizili
- Başlık 1-2-3, normal paragraf
- Sola / ortaya / sağa hizalama, iki yana yaslama
- Madde işaretli ve numaralı liste
- Geri al / yinele

**Excel**
- .xlsx / .xls / .csv aç, .xlsx veya .csv olarak kaydet
- Hücreye yazma, formül çubuğu, formüller (`=A1+B1`, `=SUM(A1:A5)`)
- Satır / sütun ekle-sil
- Birden fazla sayfa: ekle, yeniden adlandır, sil
- Enter / Tab / ok tuşlarıyla hücreler arasında gezinme
- Excel'den kopyalanan çok hücreli veriyi yapıştırma

## Kısayollar

| Kısayol | İşlev |
|---|---|
| Ctrl+O | Dosya aç |
| Ctrl+S | Kaydet |
| Ctrl+Shift+S | Farklı kaydet |
| Ctrl+N | Yeni Word belgesi |
| Ctrl+Shift+N | Yeni Excel tablosu |
| Ctrl+B / I / U | Kalın / İtalik / Altı çizili |

## Sınırlamalar

Asgari bir editördür: Word'de resim, tablo, sayfa düzeni, yazı tipi/renk ayarları (hizalama kaydedilir ancak dosya yeniden açıldığında sola döner); Excel'de hücre biçimlendirme, grafik ve formül hesaplama (formüller kaydedilir, Excel'de açıldığında hesaplanır) desteklenmez.

## Dosyalar

- `main.js` – Electron ana süreç, menü ve dosya diyalogları
- `office.js` – Word/Excel okuma-yazma mantığı
- `renderer.js` – Arayüz (Word ve Excel editörleri)
- `index.html`, `style.css` – Arayüz yapısı ve görünümü
- `ornekler/` – Test için oluşturulmuş örnek belgeler
