@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ================================================================
echo    GERAR EXE - Dimensao Anuncio
echo    Pasta: %CD%
echo ================================================================
echo.

:: ── 1. Localiza Python ────────────────────────────────────────────
set PY_CMD=
where py >nul 2>&1 && set PY_CMD=py && goto :python_ok
where python >nul 2>&1 && set PY_CMD=python && goto :python_ok
where python3 >nul 2>&1 && set PY_CMD=python3 && goto :python_ok

echo [ERRO] Python nao encontrado!
echo Baixe em: https://www.python.org/downloads/
echo IMPORTANTE: marque "Add Python to PATH" na instalacao.
echo.
pause & exit /b 1

:python_ok
echo [OK] Python encontrado: %PY_CMD%
%PY_CMD% --version
echo.

:: ── 2. Atualiza pip ───────────────────────────────────────────────
echo [1/9] Atualizando pip...
%PY_CMD% -m pip install --upgrade pip --quiet
echo.

:: ── 3. Instala dependencias do app ───────────────────────────────
echo [2/9] Instalando PyMuPDF...
%PY_CMD% -m pip install PyMuPDF --quiet

echo [3/9] Instalando Pillow...
%PY_CMD% -m pip install Pillow --quiet

echo [4/9] Instalando opencv...
%PY_CMD% -m pip install opencv-python-headless --quiet

echo [5/9] Instalando imagehash...
%PY_CMD% -m pip install imagehash --quiet

echo [6/9] Instalando scikit-image...
%PY_CMD% -m pip install scikit-image --quiet

echo [7/9] Instalando numpy...
%PY_CMD% -m pip install numpy --quiet

echo [8/9] Instalando pytesseract...
%PY_CMD% -m pip install pytesseract --quiet

echo.

:: ── 4. Instala PyInstaller ────────────────────────────────────────
echo [9/9] Instalando PyInstaller...
%PY_CMD% -m pip install pyinstaller --quiet
echo.

:: ── 5. Verifica se main.py existe ─────────────────────────────────
if not exist "main.py" (
    echo [ERRO] Arquivo main.py nao encontrado em %CD%
    pause & exit /b 1
)

:: ── 6. Limpa builds anteriores ────────────────────────────────────
echo [BUILD] Limpando builds anteriores...
if exist "build"  rmdir /s /q "build"
if exist "dist"   rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"
echo.

:: ── 7. Gera o .exe ────────────────────────────────────────────────
echo [BUILD] Gerando executavel... (pode levar alguns minutos)
echo.

%PY_CMD% -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "DimensaoAnuncio" ^
    --collect-data fitz ^
    --collect-binaries fitz ^
    --copy-metadata PyMuPDF ^
    --collect-data cv2 ^
    --collect-binaries cv2 ^
    --collect-data skimage ^
    --collect-data imagehash ^
    --collect-data numpy ^
    --hidden-import fitz ^
    --hidden-import fitz._fitz ^
    --hidden-import PIL ^
    --hidden-import PIL.ImageDraw ^
    --hidden-import PIL.ImageTk ^
    --hidden-import PIL.ImageEnhance ^
    --hidden-import PIL.ImageFilter ^
    --hidden-import pytesseract ^
    --hidden-import difflib ^
    --hidden-import unicodedata ^
    --hidden-import collections ^
    --hidden-import hashlib ^
    --hidden-import pathlib ^
    --hidden-import tkinter ^
    --hidden-import tkinter.ttk ^
    --hidden-import tkinter.filedialog ^
    --hidden-import tkinter.messagebox ^
    --exclude-module commandline ^
    main.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERRO] Falha ao gerar o executavel!
    echo Verifique as mensagens acima.
    pause & exit /b 1
)

:: ── 8. Monta a pasta de entrega para o gestor ─────────────────────
echo.
echo [OK] Executavel gerado com sucesso!
echo.
echo [ENTREGA] Montando pasta para o gestor...

set PASTA_ENTREGA=%CD%\PARA_O_GESTOR
if exist "%PASTA_ENTREGA%" rmdir /s /q "%PASTA_ENTREGA%"
mkdir "%PASTA_ENTREGA%"

:: Copia o exe
copy "dist\DimensaoAnuncio.exe" "%PASTA_ENTREGA%\" >nul

:: Cria o bat de instalacao do Tesseract para o gestor
(
echo @echo off
echo cd /d "%%~dp0"
echo echo =====================================================
echo echo   Instalador do Tesseract-OCR ^(necessario para OCR^)
echo echo =====================================================
echo echo.
echo.
echo :: Verifica se ja esta instalado
echo if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" ^(
echo     echo [OK] Tesseract ja esta instalado!
echo     goto :fim
echo ^)
echo.
echo echo Baixando instalador do Tesseract... Por favor aguarde.
echo echo.
echo powershell -Command "Invoke-WebRequest -Uri 'https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.5.0.20241111.exe' -OutFile 'tesseract_setup.exe'"
echo.
echo if not exist "tesseract_setup.exe" ^(
echo     echo [ERRO] Nao foi possivel baixar o instalador.
echo     echo Acesse manualmente: https://github.com/UB-Mannheim/tesseract/wiki
echo     pause ^& exit /b 1
echo ^)
echo.
echo echo Instalando Tesseract... Siga as instrucoes na tela.
echo start /wait tesseract_setup.exe
echo del tesseract_setup.exe
echo.
echo :fim
echo echo.
echo echo Pronto! Agora execute o DimensaoAnuncio.exe
echo echo.
echo pause
) > "%PASTA_ENTREGA%\1_instalar_tesseract.bat"

:: Cria o LEIA_ME
(
echo =====================================================
echo   DIMENSAO ANUNCIO - Instrucoes para o Gestor
echo =====================================================
echo.
echo PASSO 1: Execute "1_instalar_tesseract.bat"
echo   - Isso instala o Tesseract-OCR (necessario para a
echo     funcao de leitura de texto em imagens/OCR^).
echo   - So precisa fazer isso UMA vez.
echo   - Requer conexao com a internet.
echo.
echo PASSO 2: Execute "DimensaoAnuncio.exe"
echo   - O programa abre normalmente, sem precisar de Python.
echo.
echo OBS: O Windows pode exibir um aviso de seguranca
echo   na primeira execucao. Clique em "Mais informacoes"
echo   e depois "Executar assim mesmo".
echo.
echo =====================================================
) > "%PASTA_ENTREGA%\LEIA_ME.txt"

:: ── 9. Resumo final ───────────────────────────────────────────────
echo.
echo ================================================================
echo   CONCLUIDO!
echo.
echo   Pasta gerada: %PASTA_ENTREGA%
echo.
echo   Conteudo:
echo     DimensaoAnuncio.exe       ^<-- programa principal
echo     1_instalar_tesseract.bat  ^<-- gestor executa primeiro
echo     LEIA_ME.txt               ^<-- instrucoes
echo.
echo   Compacte a pasta e envie para o seu gestor.
echo ================================================================
echo.

:: Abre a pasta no Explorer
explorer "%PASTA_ENTREGA%"

pause
