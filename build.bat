@echo off
echo ================================================================================
echo GABiM KONTROL PROGRAMI - Build Script
echo ================================================================================
echo.

REM Renkli çıktı için
color 0A

echo [1/5] Gerekli kutuphaneleri kontrol ediliyor...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller bulunamadi. Yukleniyor...
    pip install pyinstaller
) else (
    echo PyInstaller zaten yuklu.
)

echo.
echo [2/5] Eski build dosyalari temizleniyor...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "installer_output" rmdir /s /q "installer_output"
echo Temizlik tamamlandi.

echo.
echo [3/5] PyInstaller ile EXE dosyasi olusturuluyor...
pyinstaller excel_kontrolor.spec --clean
if errorlevel 1 (
    echo HATA: EXE dosyasi olusturulamadi!
    pause
    exit /b 1
)
echo EXE dosyasi basariyla olusturuldu: dist\GabimKontrolProgram.exe

echo.
echo [4/5] Inno Setup kontrolu yapiliyor...
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo Inno Setup bulundu. Installer olusturuluyor...
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup_script.iss
    if errorlevel 1 (
        echo UYARI: Installer olusturulamadi!
    ) else (
        echo Installer basariyla olusturuldu!
    )
) else (
    echo.
    echo UYARI: Inno Setup bulunamadi!
    echo Installer olusturmak icin Inno Setup yuklemeniz gerekiyor.
    echo Indirme linki: https://jrsoftware.org/isdl.php
    echo.
    echo Sadece EXE dosyasi olusturuldu: dist\GabimKontrolProgram.exe
)

echo.
echo [5/5] Build islemi tamamlandi!
echo ================================================================================
echo.
echo Olusturulan dosyalar:
echo - EXE dosyasi: dist\GabimKontrolProgram.exe
if exist "installer_output\GabimKontrolProgram_Setup_v1.0.0.exe" (
    echo - Setup dosyasi: installer_output\GabimKontrolProgram_Setup_v1.0.0.exe
)
echo.
echo ================================================================================
pause
