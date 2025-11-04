# GABİM KONTROL PROGRAMI - Kurulum ve Dağıtım Kılavuzu

**Yapımcı:** A.R.E.A. GAYRİMENKUL DEĞERLEME VE DANIŞMANLIK A.Ş.

Bu dokümanda programın nasıl derlenip dağıtılacağı anlatılmaktadır.

## İçindekiler
1. [Hızlı Başlangıç](#hızlı-başlangıç)
2. [Detaylı Kurulum](#detaylı-kurulum)
3. [Profesyonel Installer Oluşturma](#profesyonel-installer-oluşturma)
4. [Dağıtım Seçenekleri](#dağıtım-seçenekleri)
5. [Sorun Giderme](#sorun-giderme)

---

## Hızlı Başlangıç

### Basit EXE Oluşturma (Önerilen)

En kolay yöntem:

```bash
basit_build.bat
```

Bu komut:
- Gerekli tüm kütüphaneleri yükler
- Tek bir `.exe` dosyası oluşturur
- `dist\GabimKontrolProgram.exe` olarak kaydeder

**Avantajları:**
- Python yüklü olmayan bilgisayarlarda çalışır
- Tek dosya, kolayca paylaşılabilir
- Kurulum gerektirmez

---

## Detaylı Kurulum

### Ön Gereksinimler

#### 1. Python Kurulumu
- Python 3.7 veya üzeri gereklidir
- İndirme: https://www.python.org/downloads/

**Python kurulum kontrolü:**
```bash
python --version
```

#### 2. Gerekli Kütüphaneleri Yükleme
```bash
pip install -r requirements.txt
```

Bu komut şu kütüphaneleri yükler:
- `openpyxl` - Excel dosyası işleme
- `pyinstaller` - EXE dosyası oluşturma

### EXE Dosyası Oluşturma

#### Yöntem 1: Basit Build (Önerilen)
```bash
basit_build.bat
```

#### Yöntem 2: Spec Dosyası ile Build
```bash
pyinstaller excel_kontrolor.spec --clean
```

#### Yöntem 3: Manuel PyInstaller Komutu
```bash
pyinstaller --onefile --windowed --name "ExcelKontrolProgram" excel_kontrolor.py
```

**Parametreler:**
- `--onefile`: Tek EXE dosyası oluştur
- `--windowed`: Konsol penceresi gösterme
- `--name`: EXE dosya adı

---

## Profesyonel Installer Oluşturma

### Gereksinimler

**Inno Setup 6** yüklü olmalıdır:
- İndirme: https://jrsoftware.org/isdl.php
- Ücretsiz ve açık kaynaklıdır

### Installer Oluşturma Adımları

#### 1. Tam Build (EXE + Installer)
```bash
build.bat
```

Bu script:
1. Eski build dosyalarını temizler
2. PyInstaller ile EXE oluşturur
3. Inno Setup ile profesyonel installer oluşturur

#### 2. Manuel Inno Setup Kullanımı

```bash
# Önce EXE oluştur
pyinstaller excel_kontrolor.spec --clean

# Sonra Inno Setup ile derle
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup_script.iss
```

### Installer Özellikleri

Oluşturulan installer:
- ✓ Program Files klasörüne kurulum
- ✓ Başlat Menüsü kısayolu
- ✓ Desktop kısayolu (isteğe bağlı)
- ✓ Hızlı başlatma kısayolu (isteğe bağlı)
- ✓ Kaldırma (uninstall) desteği
- ✓ Türkçe arayüz
- ✓ Yönetici hakları ile kurulum

**Çıktı:**
`installer_output\ExcelKontrolProgram_Setup_v1.0.0.exe`

---

## Dağıtım Seçenekleri

### Seçenek 1: Sadece EXE Dosyası (Basit)

**Avantajlar:**
- Kurulum gerektirmez
- Tek dosya paylaşımı
- USB'den çalıştırılabilir

**Nasıl Dağıtılır:**
1. `basit_build.bat` çalıştır
2. `dist\ExcelKontrolProgram.exe` dosyasını kopyala
3. Hedef bilgisayara yapıştır
4. Çift tıkla ve çalıştır

**Uygun Durumlar:**
- Az sayıda kullanıcı
- Hızlı test ve demo
- Taşınabilir kullanım

### Seçenek 2: Profesyonel Installer (Önerilen)

**Avantajlar:**
- Profesyonel görünüm
- Otomatik kısayol oluşturma
- Kolay kaldırma
- Merkezi yönetim desteği

**Nasıl Dağıtılır:**
1. `build.bat` çalıştır
2. `installer_output\ExcelKontrolProgram_Setup_v1.0.0.exe` dosyasını paylaş
3. Kullanıcılar setup'ı çalıştırır
4. Kurulum sihirbazı her şeyi otomatik yapar

**Uygun Durumlar:**
- Şirket içi dağıtım
- Çok sayıda kullanıcı
- Standart kurulum gereksinimi
- Merkezi güncelleme ihtiyacı

### Seçenek 3: Network Paylaşımı

**Kurulum:**
1. Setup dosyasını ağ paylaşımına koy
2. Kullanıcılara paylaşım yolunu bildir
3. Kullanıcılar ağdan kurulumu çalıştırır

**Örnek:**
```
\\SUNUCU\Paylaşım\ExcelKontrolProgram_Setup_v1.0.0.exe
```

### Seçenek 4: Portable Versiyon

**Özellikler:**
- Kurulum gerektirmez
- Registry değişikliği yapmaz
- USB'den çalışabilir

**Hazırlama:**
1. `dist\ExcelKontrolProgram.exe` dosyasını al
2. Bir klasöre koy
3. İsteğe bağlı README.md ekle
4. Klasörü ZIP'le
5. Dağıt

---

## Build Dosyaları ve Klasör Yapısı

```
GABİM/
├── excel_kontrolor.py          # Ana program
├── excel_kontrolor.spec        # PyInstaller yapılandırma
├── version_info.txt            # Versiyon bilgisi
├── setup_script.iss            # Inno Setup script
├── build.bat                   # Tam build script
├── basit_build.bat             # Basit build script
├── requirements.txt            # Python bağımlılıkları
├── README.md                   # Kullanım kılavuzu
├── KURULUM_VE_DAGITIM.md       # Bu dosya
│
├── build/                      # Geçici build dosyaları (silinebilir)
├── dist/                       # Oluşturulan EXE dosyası
│   └── ExcelKontrolProgram.exe
└── installer_output/           # Oluşturulan setup dosyası
    └── ExcelKontrolProgram_Setup_v1.0.0.exe
```

---

## Sorun Giderme

### Python Bulunamıyor Hatası

**Sorun:** `'python' is not recognized as an internal or external command`

**Çözüm:**
1. Python'un yüklü olduğundan emin olun
2. PATH'e eklendiğini kontrol edin
3. Veya tam yolu kullanın: `C:\Python39\python.exe`

### PyInstaller Yüklenmiyor

**Sorun:** `No module named 'pyinstaller'`

**Çözüm:**
```bash
pip install --upgrade pip
pip install pyinstaller
```

### EXE Oluşturulamıyor

**Sorun:** PyInstaller hata veriyor

**Çözüm:**
```bash
# Cache'i temizle
pip cache purge

# PyInstaller'ı yeniden yükle
pip uninstall pyinstaller
pip install pyinstaller

# Temiz build
rmdir /s /q build dist
pyinstaller excel_kontrolor.spec --clean
```

### Inno Setup Bulunamıyor

**Sorun:** `Inno Setup bulunamadı!`

**Çözüm:**
1. https://jrsoftware.org/isdl.php adresinden indirin
2. Varsayılan konuma yükleyin: `C:\Program Files (x86)\Inno Setup 6\`
3. Veya `setup_script.iss` dosyasında yolu güncelleyin

### EXE Çalışmıyor

**Sorun:** EXE dosyası açılmıyor veya hata veriyor

**Çözüm:**
1. Windows Defender/Antivirüs kontrolü
2. Yönetici olarak çalıştırmayı deneyin
3. Konsol modunda test edin:
```bash
pyinstaller --onefile --console excel_kontrolor.py
```

### Antivirus False Positive

**Sorun:** Antivirüs programı EXE'yi engelliyor

**Çözüm:**
1. Bu, PyInstaller ile oluşturulan programlarda normaldir
2. EXE'yi istisnalara ekleyin
3. Veya dijital imza ekleyin (ücretli)
4. VirusTotal'de taratın ve sonuçları paylaşın

---

## Versiyonlama

Program versiyonunu güncellemek için:

1. **version_info.txt** dosyasını düzenle:
```python
filevers=(1, 1, 0, 0),  # 1.0.0.0 -> 1.1.0.0
```

2. **setup_script.iss** dosyasını düzenle:
```
#define MyAppVersion "1.1.0"
```

3. Yeniden build et:
```bash
build.bat
```

---

## Güncelleme Stratejisi

### Versiyon Kontrolü

**Tavsiye edilen versiyon numarası formatı:**
- `1.0.0` - İlk sürüm
- `1.0.1` - Hata düzeltmeleri
- `1.1.0` - Yeni özellikler (geriye uyumlu)
- `2.0.0` - Büyük değişiklikler (geriye uyumsuz olabilir)

### Kullanıcılara Güncelleme Dağıtımı

1. Yeni versiyon oluştur
2. Test et
3. CHANGELOG dosyası hazırla
4. Kullanıcılara duyur
5. Yeni setup dosyasını dağıt

---

## Lisanslama ve Dağıtım Hakları

Bu program şahsi ve ticari kullanım için serbestçe dağıtılabilir.

**Önerilen:** Dağıtım öncesi:
1. LICENSE.txt dosyası ekleyin
2. Copyright bilgisi ekleyin
3. İletişim bilgisi ekleyin

---

## Destek ve İletişim

Sorunlar için:
1. README.md dosyasını okuyun
2. Bu dokümandaki Sorun Giderme bölümünü inceleyin
3. Build loglarını kontrol edin

---

## Özet - Hızlı Komutlar

```bash
# Basit kullanım (EXE oluştur)
basit_build.bat

# Profesyonel kullanım (EXE + Installer)
build.bat

# Manuel PyInstaller
pyinstaller excel_kontrolor.spec --clean

# Sadece Installer oluştur (EXE zaten varsa)
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup_script.iss
```

---

**Son Güncelleme:** 30 Ekim 2025
**Versiyon:** 1.0.0
