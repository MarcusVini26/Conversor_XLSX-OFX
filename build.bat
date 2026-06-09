@echo off
chcp 65001 >nul
echo.
echo ============================================
echo   Build: Conversor SISPAG-OFX.exe
echo ============================================
echo.

pyinstaller --onefile --windowed --name "Conversor SISPAG-OFX" --add-data "xls_to_ofx.py;." gui.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] Build falhou. Verifique se o PyInstaller esta instalado:
    echo   pip install pyinstaller
    echo.
    pause
    exit /b 1
)

echo.
echo Limpando arquivos temporarios...
if exist build rmdir /s /q build
if exist "Conversor SISPAG-OFX.spec" del /q "Conversor SISPAG-OFX.spec"

echo.
echo ============================================
echo   Pronto! Executavel gerado em:
echo   dist\Conversor SISPAG-OFX.exe
echo ============================================
echo.
pause
