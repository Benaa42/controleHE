@echo off
:: Muda para a pasta onde o bat esta localizado
cd /d "%~dp0"

echo ============================================
echo  Build - Controle de Horas Extras
echo ============================================

set PYTHON=C:\Users\renat\AppData\Local\Python\bin\python.exe

:: Instala dependencias se necessario
echo [1/3] Verificando dependencias...
"%PYTHON%" -m pip install openpyxl pyinstaller pillow --quiet

:: Limpa cache antigo do PyInstaller
echo [2/3] Limpando cache anterior...
if exist "build"          rmdir /s /q "build"
if exist "dist"           rmdir /s /q "dist"
if exist "Controle HE.spec" del /q "Controle HE.spec"

:: Gera o executavel
echo [3/3] Gerando executavel...
"%PYTHON%" -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name "Controle HE" ^
  --icon=icon.ico ^
  --add-data "icon.ico;." ^
  --add-data "Icon_maior.png;." ^
  controle_he.py

:: Copia o exe para a pasta raiz
echo.
if exist "dist\Controle HE.exe" (
  copy /Y "dist\Controle HE.exe" "Controle HE.exe"
  echo  EXE gerado com sucesso: "Controle HE.exe"
) else (
  echo  ERRO: arquivo nao encontrado em dist\
)

echo.
pause
