# mask_pdf.py — PDF Üzerinde Anahtar Kelime Maskeleme

`anahtar_kelimeler.csv` dosyasındaki anahtar kelimeleri, PDF'in görsel
yerleşimini bozmadan karşılıklarıyla **tamamen siler ve yerine yeniden yazar**
(yalnızca üzerini boyamaz — bkz. aşağıdaki "Silme gerçekten silme mi?"
bölümü).

## Kurulum

```bash
pip install pymupdf pikepdf
```

## 1. Adım — Eşleştirme dosyasını doldurun

`anahtar_kelimeler.csv` şu an 32 satır placeholder içeriyor:

```
anahtar_kelime,karsilik
ANAHTAR_KELIME_01,1
ANAHTAR_KELIME_02,2
...
```

`anahtar_kelime` sütununa PDF'lerde geçen gerçek kelime/ifadeleri (şube adı,
kod, vb.), `karsilik` sütununa neyle değiştirileceğini yazın:

```
anahtar_kelime,karsilik
```

## 2. Adım — Ayarları düzenleyin ve çalıştırın

Komut satırı argümanı **yok**. `mask_pdf.py` dosyasını açın, en üstteki
"AYARLAR" bölümündeki dört satırı kendi durumunuza göre düzenleyin:

```python
# PDF_INPUT_PATH: Tek bir PDF dosyası ya da PDF'lerin bulunduğu bir klasör
#                 olabilir (klasörse içindeki tüm .pdf dosyaları işlenir).
PDF_INPUT_PATH = "raporlar/"

# ANAHTAR_KELIMELER_CSV: Anahtar kelime/karşılık eşleştirme tablosunun yolu.
ANAHTAR_KELIMELER_CSV = "anahtar_kelimeler.csv"

# OUTPUT_DIR: Maskelenmiş PDF'lerin yazılacağı klasör (yoksa oluşturulur).
OUTPUT_DIR = "maskeli/"

# DRY_RUN: True yapılırsa dosyalar DEĞİŞTİRİLMEZ, yalnızca kaç eşleşme
#          bulunduğu konsolda gösterilir.
DRY_RUN = False
```

- Tek bir dosyayı işlemek isterseniz `PDF_INPUT_PATH`'i o dosyanın yoluna
  ayarlayın (örn. `"rapor.pdf"`).
- Klasördeki tüm PDF'leri (30 raporunuz için) işlemek isterseniz
  `PDF_INPUT_PATH`'i klasör yoluna ayarlayın (örn. `"raporlar_klasoru/"`).
- Önce sadece kaç eşleşme bulunduğunu görmek isterseniz (dosyaları
  değiştirmeden) `DRY_RUN = True` yapın; sonuçtan emin olunca `False`'a
  çevirip gerçek çıktıyı üretin.

Ayarları kaydettikten sonra, hiçbir argüman vermeden çalıştırın:

```bash
python mask_pdf.py
```

Çalıştırdıktan sonra konsolda her anahtar kelime için kaç eşleşme bulunduğunu
gösteren bir özet çıkar — "hiç bulunamadı" uyarısı çıkan kelimeler için
yazımı/boşlukları PDF'teki haliyle birebir eşleştiğinizden emin olun.

## Bilinen kısıt: karşılık, orijinal kelimeden çok uzunsa

PDF'lerde metin sabit konumda durur; Word gibi "reflow" (metnin otomatik
kayması) yoktur. Script bunu şöyle yönetir:

- **Karşılık orijinalden kısa veya benzer uzunluktaysa** : sorunsuz çalışır, en fazla küçük bir boşluk
  kalabilir.
- **Karşılık orijinalden çok uzunsa**:
  script font boyutunu otomatik küçültür; yine de sığmıyorsa **konsolda
  uyarı verir** ("orijinal alana sığmıyor — PDF'i gözle kontrol edin") ve o
  belirli örnekte komşu metinle görsel çakışma olabilir.

**Öneri:** `karsilik` değerlerini mümkün olduğunca kısa tutun. Konsolda uyarı çıkan dosyaları mutlaka açıp gözle kontrol
edin.

## Silme 

İşlem, eşleşen metnin PDF'in **içerik akışından tamamen
çıkarılmasıdır**. PyMuPDF varsayılan
ayarlarla kaydettiğinde, silinen orijinal metin dosyanın içinde
"kullanılmayan/referanssız obje" olarak (görünmez ama pikepdf gibi bir
araçla çıkarılabilir şekilde) kalabiliyordu. Bu, aşağıdaki iki önlemle
engellenir:

1. Dosya kaydedilirken agresif "garbage collection" (`garbage=4`) uygulanır
   — bu tür referanssız objeler dosyadan tamamen silinir.
2. Script, kaydettiği **her dosyayı otomatik olarak yeniden açıp** orijinal
   anahtar kelimelerin dosyanın hiçbir yerinde (görünür metin, gizli
   objeler, ham baytlar) kalmadığını doğrular ve sonucu konsolda raporlar.
   Bir kelime hâlâ bulunursa **"KRİTİK UYARI"** ile açıkça belirtilir ve o
   dosyayı kullanmamanız söylenir.

Bu davranış pikepdf ile ham dosya içeriği taranarak doğrulanmıştır;
düzeltme sonrasında hiçbir iz kalmadığı teyit edilmiştir.

## Büyük/küçük harf ve Türkçe İ/I/ı/i

Türkçe'de İngilizce'den farklı olarak 4 ayrı "i" harfi vardır (İ/i noktalı,
I/ı noktasız). Hem PyMuPDF'in hem Python'un standart büyük/küçük harf
dönüşümü bu ayrımı hatalı yapar (örn. "İstanbul".lower() Türkçe'de olması
gerekenden farklı bir sonuç verir). Bu nedenle script'e Türkçe'ye özel bir
varyant üretici eklenmiştir:

- 

## Türkçe karakterler (görüntüleme)

Script, ş/ğ/ı/İ/ç/ö/ü karakterlerinin doğru görünmesi için sistemde bulunan
bir Unicode font (DejaVu Sans) kullanır. Farklı bir ortamda çalıştırırsanız
ve bu font bulunamazsa, script bir uyarı basar — bu durumda `TR_FONT_CANDIDATES`
listesine kendi sisteminizdeki bir Unicode TTF font yolunu ekleyin.

---

# mask_excel.py — Excel Üzerinde Anahtar Kelime Maskeleme

`mask_pdf.py` ile AYNI mantık ve AYNI `anahtar_kelimeler.csv` formatı,
Excel (.xlsx) dosyaları için. Bir dosyadaki **tüm sayfaları** (tek sayfayla
sınırlı değil) tarar; hücre değerleri, sayfa adları, hücre yorumları ve
sayfa üstbilgi/altbilgileri dahil, eşleşen her yeri karşılığıyla değiştirir.

## Kurulum

```bash
pip install openpyxl
```

## Kullanım

Komut satırı argümanı **yok** — aynı `mask_pdf.py` gibi, dosyanın başındaki
"AYARLAR" bölümünü düzenleyip çalıştırın:

```python
EXCEL_INPUT_PATH = "istatistikler/"        # tek .xlsx ya da klasör
ANAHTAR_KELIMELER_CSV = "anahtar_kelimeler.csv"  # mask_pdf.py ile AYNI dosya
OUTPUT_DIR = "maskeli/"
DRY_RUN = False
```

```bash
python mask_excel.py
```

`ANAHTAR_KELIMELER_CSV` varsayılan olarak `mask_pdf.py` ile **aynı** dosyayı
gösterir — böylece bir şube adı/kodu hem PDF'te hem Excel'de aynı karşılığa
(örn. `SUBE_001`) maskelenir. Farklı bir eşleştirme istiyorsanız AYARLAR
bölümünden başka bir CSV yolu verebilirsiniz.

## Neyi tarar (tek sayfa değil, tüm sayfalar)

- Her sayfadaki her hücrenin değeri (metin veya sayı — örn. şube kodu Excel'de
  sayı olarak girilmişse de yakalanır)
- Sayfa adlarının kendisi (bir sayfa şube adıyla adlandırılmışsa)
- Hücre yorumları (comment)
- Sayfa üstbilgi/altbilgisi (varsa)

## Silme gerçekten silme mi?

Evet. xlsx dosyası aslında bir ZIP arşividir; hücre bazlı değiştirme
(yukarıdaki liste) çoğu durumu kapsar, ama örneğin bir **grafik başlığı**
gibi openpyxl'in hücre modelinin dışında kalan yerler olabilir. Bunun için:

1. Dosya kaydedildikten sonra, içindeki **her XML parçası** ham metin
   düzeyinde taranır (grafikler, belge özellikleri, ilişki dosyaları dahil).
2. Bir yerde hâlâ anahtar kelime bulunursa, doğrudan o XML parçası üzerinde
   onarım yapılır (bunu test ederken bizzat bir grafik başlığında yakaladık
   ve düzelttik).
3. Onarımdan sonra dosya tekrar taranıp doğrulanır; hâlâ bir şey bulunursa
   **"KRİTİK UYARI"** ile açıkça belirtilir.

**Bilinen bir tuzak (ve nasıl çözüldü):** karşılık değeri anahtar kelimeyi
alt dize olarak içerebilir (örn. `"İstanbul"` → `"SUBE_ISTANBUL"` karşılığının
içinde `"ISTANBUL"` geçer). Bunu test ederken, doğru maskelenmiş metnin
kendi içindeki bu çakışma yüzünden **tekrar** eşleşip `"SUBE_SUBE_ISTANBUL"`
gibi çifte maskelenmesi hatasını bulduk. Script artık hem taramayı (bir
metindeki her pozisyonu yalnızca bir kez işleyen tek geçişli algoritma) hem
doğrulamayı (zaten uygulanmış karşılık değerlerini önce ayıklayıp öyle
arama) bu soruna karşı korumalı şekilde yapıyor — bunu yine bizzat test
ederek doğruladık.

## Büyük/küçük harf ve Türkçe İ/I/ı/i

`mask_pdf.py` ile birebir aynı mantık — İstanbul/istanbul/ISTANBUL ve
ünzil/ÜNZİL gibi tüm gerçekçi yazım varyantları eşleşir; `unzil` (Türkçe
karaktersiz, gerçekten farklı bir kelime) eşleşmez. Bunu Excel üzerinde de
ayrıca test ettik.

## Bilinen kısıtlar

- Yalnızca `.xlsx` desteklenir. Eski `.xls` formatındaki dosyalar için
  script bir uyarı basar — önce Excel'de "Farklı Kaydet" ile `.xlsx`'e
  dönüştürün.
- Anahtar kelimeniz bir hücre referansına benziyorsa (örn. `"A1"`), formül
  içeren hücrelerde teorik olarak yanlış eşleşme riski vardır — gerçek şube
  adı/kodu gibi anahtar kelimelerde bu risk yok denecek kadar azdır.
