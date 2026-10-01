# Sesli Sunum

Windows masaüstü PowerPoint ve Python 3.11/3.12 (64 bit, tkinter dahil) gerekir.
Bu bir ilk prototiptir. Gerçek şirket dosyalarından önce üç slaytlık sahte içerikle deneyin.

## Bir defalık kurulum

Şirketinizin izin verdiği Python kurulumunu kullanın. Bu klasörde PowerShell açıp çalıştırın:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
New-Item -ItemType Directory -Force models
.\.venv\Scripts\python.exe -m piper.download_voices de_DE-thorsten-medium --data-dir models
```

Python 3.11 kullanıyorsanız ilk komutta `-3.11` yazın. Kurulum ve model indirme internet gerektirir;
bu adımlarda herhangi bir sunum veya Word dosyası işlenmez. İnternetsiz şirkete kurulum için IT,
uygun Windows/Python sürümüne ait paketleri ve model dosyalarını onaylı ortamda hazırlayıp aktarabilir.
Piper GPL-3.0 lisanslıdır; seçilen sesin MODEL_CARD lisansını da IT ile inceleyin.
Kaynaklar: https://github.com/OHF-Voice/piper1-gpl ve https://huggingface.co/rhasspy/piper-voices

## Kullanım

1. İnternet bağlantısını kapatın. Dosyaları OneDrive gibi senkronize klasörler yerine yerel klasörde tutun.
2. `Baslat.cmd` dosyasını açın.
3. PPTX (veya eski PPT) ile DOCX dosyasını seçin. Eski DOC dosyasını Word'de DOCX olarak kaydedin.
4. `models` içindeki `.onnx` dosyasını seçin. Yanında aynı adlı `.onnx.json` bulunmalıdır.
5. Yerel çıktı klasörünü seçin, “Metinleri önizle” düğmesine basıp eşleştirmeyi inceleyin.
6. “Sesli PPTX oluştur” düğmesine basın. İşlem boyunca uygulamayı açık tutun.
7. Yeni oluşturulan tarihli klasördeki `_sesli.pptx` dosyasını PowerPoint'te açıp F5 ile deneyin.

Her Word başlığı ayrı paragraf olmalı:

```text
FOLIE - 1.
Guten Tag. Heute stellen wir unser Beispielprojekt vor.

FOLIE - 2.
Im zweiten Schritt betrachten wir die wichtigsten Ergebnisse.

FOLIE - 3.
Vielen Dank für Ihre Aufmerksamkeit.
```

Başlık büyük/küçük harfi, tire çevresindeki boşluklar ve son nokta esnektir.
Slayt numarası fiziksel sıra numarasıdır. Gizli slaytlar da sayılır; denemeden önce görünür yapın.
Eksik, fazla, tekrarlanan veya boş metinler hata verir. Metin kutuları, üstbilgiler ve dipnotlar
okunmaz; ana belge paragrafları ve tablolardaki paragraflar okunur. İlk başlıktan önce metin olmamalıdır.

## Çıktı ve sınırlar

Sesler PPTX içine gömülür; `sesler` klasörü kontrol için ayrıca saklanır. Kaynak değiştirilmez.
Ses efekti animasyon sırasının başına, sıfır gecikmeyle yerleştirilir.
Geçiş süresi ses süresi + mevcut slayt geçiş efekti süresi + 2 saniyedir.
Bu, sesin bitmesini bekleyen bir olay dinleyicisi değildir; süreye dayalı geçiştir.
İlk sürüm normal model hızını kullanır.
Mevcut sesler silinmez: kaynakta anlatım, müzik veya video sesi varsa çakışabilir.
Tıklamayla çalışan animasyonlar kelimelerle eşzamanlanmaz; bunları ayrıca kontrol edin.
MP4 için PowerPoint'te Dosya > Dışa Aktar > Video Oluştur yolunu kullanıp kayıtlı zamanlamaları seçin.
Bir hata sonrası oluşmuş eksik klasörü kullanmayın; hatayı giderip tekrar çalıştırın.

Uygulama kodunda ağ isteği, bulut API'si veya model indirme yoktur. Ses üretimi yerel modelle yapılır.
Office ve diğer yazılımların ağ davranışları uygulamadan bağımsızdır; tam çevrimdışı kullanım
internet bağlantısı kapalıyken doğrulanmalıdır. Tanılama ekranı dosya yollarını gösterebilir.

Bu paket kaynak kod ve başlatıcı içerir; bağımsız EXE değildir. PowerPoint'teki ses başlangıcı,
animasyon etkileşimi ve geçişler hedef bilgisayarda deneme sunumuyla doğrulanmalıdır.
