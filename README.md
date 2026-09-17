# EN-TR-Speech-Suite

Windows üzerinde panodaki metni seslendiren ve mikrofondan algıladığı konuşmayı aktif yazı alanına aktaran Python/Tkinter uygulaması.

Metinden sese dönüşüm için **Windows SAPI**, konuşmadan metne dönüşüm için **Vosk** kullanır. Türkçe ve İngilizce dil seçimi, okuma ve dinleme için ortaktır.

## Özellikler

- Panoya kopyalanan yeni metni otomatik seslendirme.
- **SESLİ OKU** düğmesiyle panodaki metni elle okutma.
- **SESLİ YAZ** düğmesiyle mikrofon üzerinden dikte.
- Türkçe ve İngilizce arasında düğmeyle veya sesli komutla geçiş.
- Okuma hızını ayarlama.
- Tanınan metni `Ctrl+V` ile aktif yazı alanına yapıştırma.
- Dikte sırasında yapılan pano değişikliklerinin otomatik okumayı tetiklemesini önleme.
- SAPI seslendirmesi sırasında mikrofon verisinin işlenmesini duraklatma.
- Modelleri başlangıçta bir kez yükleme; dinlemeyi kapatıp açarken yeniden yüklememe.

## Gereksinimler

- Windows 10 veya Windows 11.
- Tercihen 64 bit Python ve aynı mimariye uygun bağımlılıklar.
- Çalışan bir mikrofon ve Windows mikrofon erişim izni.
- Windows SAPI tarafından kullanılabilen Türkçe ve İngilizce sesler.
- Türkçe ve İngilizce Vosk model dosyaları.

Model ve bağımlılıklar indirildikten sonra okuma ve dikte için internet bağlantısı gerekmez. Python paketleri ve modeller hedef bilgisayara çevrimdışı yöntemle de taşınabilir.

> Bu proje Windows'a özeldir; SAPI, COM ve Windows klavye/pano API'lerini kullanır.

## Proje dosyaları

| Dosya / klasör | Açıklama |
| --- | --- |
| `READ.py` | Tkinter arayüzü ve uygulama kodu |
| `BUILD.cmd` | PyInstaller ile tek dosya EXE oluşturur |
| `model/tr/` | Türkçe Vosk modelinin içeriği |
| `model/en/` | İngilizce Vosk modelinin içeriği |
| `KURULUM.txt` | Kısa kurulum notları |

Modeller dağıtılan kaynak paketine dahil değildir; ayrıca indirilmelidir.

## Python ile kurulum

### 1. Projeyi hazırlayın

Bu depoyu indirin veya klonlayın. PowerShell'de `READ.py` dosyasının bulunduğu klasörü açın.

Sanal ortam oluşturun:

```powershell
python -m venv .venv
```

Paketleri bu ortama kurun:

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install vosk sounddevice pyperclip pywin32 pyinstaller
```

`tkinter`, standart Windows Python kurulumunun Tcl/Tk bileşeniyle gelir. `tkinter` bulunamıyorsa Python kurulumunda Tcl/Tk desteğini etkinleştirin.

### 2. Vosk modellerini yerleştirin

[Vosk model listesinden](https://alphacephei.com/vosk/models) bir Türkçe ve bir İngilizce model indirin ve arşivlerini açın.

Model klasörlerini aşağıdaki konumlara yerleştirin:

| Dil | Python ile çalıştırırken | EXE ile çalıştırırken |
| --- | --- | --- |
| Türkçe | `READ.py` yanında `model/tr/` | `READ.exe` yanında `model/tr/` |
| İngilizce | `READ.py` yanında `model/en/` | `READ.exe` yanında `model/en/` |

**`tr` ve `en` klasörlerinin içinde modelin kendi dosya ve klasörleri doğrudan bulunmalıdır.** İndirilen modelin adını taşıyan fazladan bir alt klasör bırakmayın. Modelin iç yapısını koruyun.

Örnek uygulama konumu:

```text
C:\Konusan\READ.exe
C:\Konusan\model\tr\
C:\Konusan\model\en\
```

### 3. SAPI seslerini kontrol edin

Kod şu ses sırasını varsayar:

```python
ENGLISH_VOICE_INDEX = 0  # Microsoft David Desktop
TURKISH_VOICE_INDEX = 3  # Microsoft Tolga
```

Bu indeksler bilgisayara göre değişebilir. Uygulamanın kullanacağı Python ortamında sesleri listeleyin:

```powershell
.\.venv\Scripts\python.exe -c "import win32com.client; voices = win32com.client.Dispatch('SAPI.SpVoice').GetVoices(); [print(i, voices.Item(i).GetDescription()) for i in range(voices.Count)]"
```

`READ.py` içindeki iki indeks değerini çıktınıza göre düzenleyin. Windows ayarlarında görünen her sesin SAPI listesinde bulunacağını varsaymayın; yukarıdaki çıktıyı esas alın.

### 4. Uygulamayı başlatın

```powershell
.\.venv\Scripts\python.exe .\READ.py
```

Başlangıçta iki model arka planda yüklenir. **Sesle yazma hazır** mesajı gelince dikte düğmesi etkinleşir.

## Kullanım

| Kontrol | İşlev |
| --- | --- |
| **TR TÜRKÇE** | Türkçe SAPI sesi ve Vosk modelini seçer |
| **EN ENGLISH** | İngilizce SAPI sesi ve Vosk modelini seçer |
| **Panoyu otomatik oku** | Panoya kopyalanan yeni metni otomatik okur; başlangıçta açıktır |
| **SESLİ OKU** | Panodaki mevcut metni seslendirir |
| **SESLİ YAZ** | Dikteyi açar; açıldığında düğme **BİTİR** olur |
| **BİTİR** | Dikteyi kapatır |
| **DURDUR** | Seslendirmeyi durdurur; dikteyi kapatmaz |
| **Okuma Hızı** | SAPI konuşma hızını değiştirir |

Dikte için **SESLİ YAZ** düğmesine basın, ardından yazılmasını istediğiniz uygulamanın metin alanına tıklayın ve konuşun. Metin aktif pencereye yapıştırılır.

Uygulama varsayılan mikrofonu, 16 kHz mono sesle kullanır. Dinleme kapalıyken ses verisi işlenmez; mevcut uygulamada mikrofon akışı model yüklemesinden sonra açık kalır.

### Sesli komutlar

Komutlar yalnızca dikte açıkken ve konuşma sonucu tanındığında işlenir.

| Geçerli dil | Tanınan komut | Sonuç |
| --- | --- | --- |
| Türkçe | `ingilizce`, `english`, `ingiliz`, `ingilizceye` | İngilizceye geçer |
| İngilizce | `turkce`, `turkish`, `turk`, `turkceye` | Türkçeye geçer |
| Her iki dil | `kendini kapat` veya `kill the program` | Uygulamayı kapatır |

Türkçe karakterler komut kontrolünden önce sadeleştirilir. Dil değiştirme sözcüğü bir cümlenin içinde geçerse de komut olarak değerlendirilir; o tanıma sonucu yazı alanına yapıştırılmaz. Algılama başarısı kullanılan modele ve söylenen komutun mevcut dilde tanınmasına bağlıdır.

## Tek dosya EXE oluşturma

Önce Python ile uygulamanın çalıştığını doğrulayın. Ardından proje klasöründe:

```powershell
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --collect-all vosk --name READ READ.py
```

Çıktı:

```text
dist\READ.exe
```

`model` klasörünü `dist` içine kopyalayın:

```powershell
Copy-Item -LiteralPath .\model -Destination .\dist -Recurse -Force
```

Dağıtırken **`READ.exe` ve yanındaki `model` klasörünü birlikte** taşıyın. Hedef bilgisayarda Python gerekmez; uygun SAPI sesleri ve mikrofon gerekir.

### BUILD.cmd kullanımı

`BUILD.cmd` aynı derleme işlemini `python` komutuyla yapar. Bu nedenle kullandığı Python ortamında tüm bağımlılıklar bulunmalıdır.

Yukarıdaki sanal ortamı kullanmak için Komut İstemi'nde (CMD):

```bat
.venv\Scripts\activate.bat
BUILD.cmd
```

Dosyayı doğrudan çift tıklamak PATH üzerindeki Python'u kullanır; sanal ortamın otomatik seçildiği varsayılmamalıdır. `BUILD.cmd` model klasörlerini kopyalamaz.

### Auto-py-to-exe kullanımı

- Script olarak `READ.py` seçin.
- **One File** ve **Window Based** seçeneklerini kullanın.
- Ek parametrelere `--collect-all vosk` ekleyin.
- Model klasörlerini EXE içine eklemeyin; çıktıdaki EXE'nin yanında tutun.

## Dosya yolları nasıl çalışır?

| Çalıştırma biçimi | Model kökü |
| --- | --- |
| Normal Python | `READ.py` dosyasının bulunduğu klasör |
| PyInstaller EXE | `sys.executable` ile bulunan EXE klasörü |
| Thonny, `__file__` olmadan editör çalıştırması | Geçerli çalışma klasörü |

Onefile EXE çalışırken PyInstaller'ın `_MEI...` adlı geçici bir klasöre açılması normaldir. **Vosk kütüphanesi ve DLL'leri paket içinden, modeller ise EXE'nin yanından yüklenir.** Kısayolun “Başlangıç yeri” EXE modundaki model yolunu değiştirmez.

[PyInstaller çalışma zamanı ve dosya yolları](https://pyinstaller.org/en/stable/runtime-information.html)

## Sorun giderme

### `FileNotFoundError: ... _MEI...\vosk`

Bu hata model yüklenmeden, `import vosk` sırasında oluşur. Vosk'un gerekli dosyaları pakete eksik alınmış olabilir.

EXE'yi `--clean --collect-all vosk` seçenekleriyle yeniden oluşturun. Yalnızca `model` klasörünü taşımak veya çalışma dizinini değiştirmek bu paketleme hatasını çözmez.

### Model klasörü bulunamadı

Hata mesajında gösterilen tam yolu kontrol edin. `model/tr` ve `model/en` doğru kökte bulunmalıdır. EXE'yi `dist` içinden çalıştırıyorsanız modeller de `dist/model` içinde olmalıdır.

### Klasör var ama model yüklenmiyor

Model arşivinin tam açıldığını ve fazladan bir klasör seviyesi olmadığını kontrol edin. `tr` ve `en` içinde modelin kendi dosyaları bulunmalıdır.

### Yanlış ses kullanılıyor veya ses seçimi hatası oluşuyor

SAPI ses listesini yeniden alın ve `ENGLISH_VOICE_INDEX` / `TURKISH_VOICE_INDEX` değerlerini düzeltin. Kaynak kod değiştiyse EXE'yi yeniden derleyin.

### Mikrofon açılamadı

Windows'ta varsayılan giriş cihazını ve masaüstü uygulamalarının mikrofon erişimini kontrol edin. Uygulama 16 kHz mono giriş açmayı dener; seçilen cihazın bu ayarı desteklemesi gerekir.

### Konuşma tanınıyor ama yazı yazılmıyor

Hedef metin alanına tıklayın. Uygulama metni klavye simülasyonu ve `Ctrl+V` ile aktarır. Hedef uygulama yapıştırmayı engelliyorsa veya yönetici yetkisiyle çalışıyorsa aktarım başarısız olabilir.

### Pano otomatik okunuyor

Bu özellik başlangıçta açıktır. İstemiyorsanız **Panoyu otomatik oku** seçimini kaldırın.

## Doğrulama durumu

Kaynak kodun sözdizimi ve EXE / normal Python / Thonny için kök yol seçimi kontrol edilmiştir. Windows SAPI, mikrofon ve oluşturulan EXE'nin uçtan uca testi hedef bilgisayarda yapılmalıdır.
