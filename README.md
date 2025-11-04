# GABİM KONTROL PROGRAMI

Excel dosyalarındaki belirli sütun ve hücrelerde otomatik kontrol ve düzeltme işlemleri yapan Python tabanlı GUI uygulaması.

**Yapımcı:** A.R.E.A. GAYRİMENKUL DEĞERLEME VE DANIŞMANLIK A.Ş.

## Özellikler

Program, yüklenen Excel dosyasında aşağıdaki kontrol ve düzeltmeleri otomatik olarak yapar:

1. **C=3 ve AD=1** olan satırların AD ve A hücrelerini kırmızı renk ile işaretler
2. **AU > mevcut yıl** ise AU hücrelerini kırmızı renk ile işaretler
3. **AU>2000 ve C=3** olan satırların AV değerini 7 yapar
4. **C=15** olan satırların U değerini 3 yapar
5. **C=3 ve CY<50** olan satırların BL değerini 2 yapar
6. **C=3 ve CY<100** olan satırların BN değerini 2 yapar
7. **C=3 ve CY<100** olan satırların BO değerini 1 yapar
8. Tüm satırlarda **BP değeri maksimum 3** olmasını sağlar (fazla ise 3'e düzeltir)
9. **BU > S** ise BU değerini S değerine eşitler
10. **DB sütunu** değerlerini virgülden sonra maksimum 2 haneye yuvarlar

## Gereksinimler

- Python 3.7 veya üzeri
- openpyxl kütüphanesi

## Kurulum ve Kullanım

### Seçenek 1: Hazır EXE Dosyası (Python Gerektirmez) ⭐ Önerilen

**En kolay yöntem - Python bilgisi gerektirmez!**

1. `GabimKontrolProgram.exe` dosyasını edinin
2. Herhangi bir klasöre kopyalayın
3. Çift tıklayarak çalıştırın
4. Excel dosyanızı seçin ve işlemleri başlatın

**Avantajları:**
- Python kurulumu gerektirmez
- Tek dosya, taşınabilir
- Tüm Windows bilgisayarlarda çalışır

### Seçenek 2: Setup Dosyası ile Kurulum (Profesyonel)

**Şirket içi dağıtım için ideal**

1. `GabimKontrolProgram_Setup_v1.0.0.exe` dosyasını çalıştırın
2. Kurulum sihirbazını takip edin
3. Program otomatik olarak kurulur
4. Başlat menüsünden veya masaüstü kısayolundan çalıştırın

**Özellikleri:**
- Otomatik kurulum
- Başlat menüsü kısayolu
- Masaüstü kısayolu (isteğe bağlı)
- Kolay kaldırma (uninstall)

### Seçenek 3: Python ile Kaynak Koddan Çalıştırma (Geliştiriciler İçin)

**Gereksinimler:**
- Python 3.7 veya üzeri

**Kurulum:**
```bash
# Gerekli kütüphaneleri yükle
pip install -r requirements.txt
```

**Çalıştırma:**
```bash
python excel_kontrolor.py
```

## Kullanım Adımları

1. Programı açın (çift tıklayın veya çalıştırın)
2. "Excel Dosyası Seç" butonuna tıklayın
3. İşlem yapmak istediğiniz Excel dosyasını seçin (.xlsx veya .xls)
4. "Kontrolleri Başlat" butonuna tıklayın
5. Program tüm kontrolleri yapacak ve değişiklikleri kaydedecektir
6. İşlem tamamlandığında, orijinal dosyanın yanına `_düzeltilmiş` eki ile yeni dosya oluşturulur

## Örnek

Orijinal dosya: `talepler.xlsx`
İşlenmiş dosya: `talepler_düzeltilmiş.xlsx`
Değişiklik raporu: `talepler_değişiklik_raporu.txt`

## Değişiklik Raporu

Program, her işlemden sonra otomatik olarak detaylı bir değişiklik raporu oluşturur. Bu rapor şunları içerir:

- Her değişikliğin satır ve sütun bilgisi
- Eski değer ve yeni değer
- Hangi kural nedeniyle değişiklik yapıldığı
- Toplam değişiklik sayısı
- İşlem tarihi ve saati

Rapor dosyası, orijinal Excel dosyasının bulunduğu klasörde `_değişiklik_raporu.txt` uzantısı ile kaydedilir.

## Önemli Notlar

- Program orijinal dosyayı değiştirmez, yeni bir dosya oluşturur
- İşlem sırasında ilerleme durumu gösterilir
- Hata durumlarında bilgilendirme mesajı gösterilir
- Boş veya geçersiz hücre değerleri güvenli şekilde ele alınır
- Tüm değişiklikler detaylı şekilde .txt dosyasına kaydedilir

## Teknik Detaylar

- **GUI Framework**: Tkinter
- **Excel İşleme**: openpyxl
- **Dosya Formatı**: .xlsx (Excel 2007+)
- **Derleme**: PyInstaller
- **Installer**: Inno Setup

## Programı Derleme (Geliştiriciler İçin)

Programdan EXE dosyası oluşturmak için:

### Hızlı Derleme (Basit)
```bash
basit_build.bat
```

Bu komut `dist\GabimKontrolProgram.exe` dosyasını oluşturur.

### Profesyonel Derleme (EXE + Setup)
```bash
build.bat
```

Bu komut:
- `dist\GabimKontrolProgram.exe` - Çalıştırılabilir dosya
- `installer_output\GabimKontrolProgram_Setup_v1.0.0.exe` - Kurulum dosyası

**Detaylı bilgi için:** [KURULUM_VE_DAGITIM.md](KURULUM_VE_DAGITIM.md)

## Dağıtım

### Kullanıcılara Dağıtım Seçenekleri

1. **Tek EXE Dosyası** (En Basit)
   - `GabimKontrolProgram.exe` dosyasını paylaşın
   - Kurulum gerektirmez
   - USB'den çalışabilir

2. **Setup Dosyası** (Profesyonel)
   - `GabimKontrolProgram_Setup_v1.0.0.exe` dosyasını paylaşın
   - Otomatik kurulum yapar
   - Kısayollar oluşturur

3. **Network Paylaşımı** (Şirket İçi)
   - Setup dosyasını ağ paylaşımına koyun
   - Kullanıcılar oradan kurulum yapabilir

## Sistem Gereksinimleri

- **İşletim Sistemi**: Windows 7/8/10/11 (64-bit)
- **RAM**: Minimum 2 GB
- **Disk Alanı**: 50 MB
- **Diğer**: .NET Framework (Windows ile birlikte gelir)

## Sık Sorulan Sorular

**S: Python yüklü olmayan bir bilgisayarda çalışır mı?**
C: Evet! EXE dosyası tüm gerekli bileşenleri içerir.

**S: Antivirüs programı EXE'yi engelliyor?**
C: Bu, PyInstaller ile derlenen programlarda normaldir. EXE'yi güvenilir programlar listesine ekleyin.

**S: Program güncellendiğinde ne yapmalıyım?**
C: Yeni EXE veya setup dosyasını indirin ve eski sürümün üzerine kurun.

**S: Birden fazla bilgisayara nasıl kurarım?**
C: Setup dosyasını kullanarak toplu kurulum yapabilir veya EXE'yi paylaşımlı bir klasöre koyabilirsiniz.

## Lisans ve Telif Hakkı

**GABİM KONTROL PROGRAMI**
© 2025 A.R.E.A. GAYRİMENKUL DEĞERLEME VE DANIŞMANLIK A.Ş.

Bu yazılım A.R.E.A. GAYRİMENKUL DEĞERLEME VE DANIŞMANLIK A.Ş. tarafından geliştirilmiştir.

## Versiyon Geçmişi

**v1.0.0** (30 Ekim 2025)
- İlk sürüm
- 10 farklı kontrol kuralı
- Otomatik değişiklik raporu
- Modern GUI arayüz
- EXE ve Setup desteği
