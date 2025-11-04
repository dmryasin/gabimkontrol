@echo off
chcp 65001 >nul
echo ╔════════════════════════════════════════════════════════════════════════════╗
echo ║                 GABİM KONTROL PROGRAMI - Basit Build                      ║
echo ╚════════════════════════════════════════════════════════════════════════════╝
echo.

echo Sadece çalıştırılabilir EXE dosyası oluşturuluyor...
echo.

echo [Adım 1/3] Gerekli kütüphaneler yükleniyor...
pip install -q -r requirements.txt
if errorlevel 1 (
    echo HATA: Kütüphaneler yüklenemedi!
    pause
    exit /b 1
)
echo ✓ Kütüphaneler yüklendi

echo.
echo [Adım 2/3] Eski dosyalar temizleniyor...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
echo ✓ Temizlik tamamlandı

echo.
echo [Adım 3/3] EXE dosyası oluşturuluyor...
echo (Bu işlem birkaç dakika sürebilir)
pyinstaller --noconfirm --onefile --windowed ^
    --name "GabimKontrolProgram" ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.cell._writer ^
    --hidden-import openpyxl.styles.stylesheet ^
    excel_kontrolor.py

if errorlevel 1 (
    echo.
    echo ✗ HATA: EXE dosyası oluşturulamadı!
    pause
    exit /b 1
)

echo.
echo ╔════════════════════════════════════════════════════════════════════════════╗
echo ║                          İŞLEM TAMAMLANDI!                                 ║
echo ╚════════════════════════════════════════════════════════════════════════════╝
echo.
echo ✓ Program başarıyla oluşturuldu!
echo.
echo Dosya konumu: dist\GabimKontrolProgram.exe
echo.
echo Bu EXE dosyasını istediğiniz bilgisayara kopyalayabilirsiniz.
echo Python yüklü olmayan bilgisayarlarda da çalışır.
echo.
pause
