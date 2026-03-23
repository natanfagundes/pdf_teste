@echo off
:: Muda automaticamente para a pasta onde este .bat esta
cd /d "%~dp0"

echo ============================================
echo    PDF - Instalador de Dependencias
echo    Pasta: %CD%
echo ============================================
echo.

:: Encontra Python
set PY_CMD=
where py >nul 2>&1
if %ERRORLEVEL% == 0 ( set PY_CMD=py & goto :found_python )
where python >nul 2>&1
if %ERRORLEVEL% == 0 ( set PY_CMD=python & goto :found_python )
where python3 >nul 2>&1
if %ERRORLEVEL% == 0 ( set PY_CMD=python3 & goto :found_python )

echo [ERRO] Python nao encontrado no sistema!
echo.
echo Baixe em: https://python.org/downloads
echo IMPORTANTE: Marque "Add Python to PATH" na instalacao!
echo.
pause & exit /b 1

:found_python
echo [OK] Python encontrado: %PY_CMD%
%PY_CMD% --version
echo.

echo [1/10] Atualizando pip...
%PY_CMD% -m pip install --upgrade pip --quiet

echo [2/10] Instalando PyMuPDF...
%PY_CMD% -m pip install PyMuPDF --quiet

echo [3/10] Instalando pillow e psutil...
%PY_CMD% -m pip install pillow psutil --quiet

echo [4/10] Instalando pyinstaller...
%PY_CMD% -m pip install pyinstaller --quiet

echo [5/10] Instalando opencv...
%PY_CMD% -m pip install opencv-python-headless --quiet

echo [6/10] Instalando imagehash...
%PY_CMD% -m pip install imagehash

echo [7/10] Instalando SCKIT...
%PY_CMD% -m pip install scikit-image

echo [8/10] Instalando Numpy...
%PY_CMD% -m pip install numpy

echo [9/10] Instalando Tesseract...
%PY_CMD% -m pip install pytesseract

echo.
echo Instalacao concluida!
echo.

set /p EXEC=Deseja executar o PDF_FINDER	 agora? (S/N): 
if /i "%EXEC%"=="S" (
    echo Iniciando...
    %PY_CMD% main.py
)
pause