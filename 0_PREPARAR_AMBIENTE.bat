@echo off
chcp 65001 >nul
title Kit - preparar o ambiente
cd /d "%~dp0"

echo PREPARAR O AMBIENTE
echo Instala neste computador o que faltar para rodar o kit, nas versões com que os resultados foram gerados:
echo   1. Python 3.12.10, só para o usuário atual;
echo   2. Tesseract OCR 5.4.0, que faz o OCR das páginas escaneadas: o Windows vai pedir permissão de administrador;
echo   3. os pacotes do Python de requisitos.txt, numa pasta só do kit: ambiente_python\
echo O que já estiver instalado é mantido. Precisa de internet e pode levar alguns minutos.
echo.
echo Pressione qualquer tecla para começar, ou feche esta janela para cancelar.
pause >nul

echo.
echo ---------- 1. Python 3.12 ----------
call :achar_python_base
if not errorlevel 1 goto python_ok
call :tem_winget || goto sem_winget
echo Instalando o Python 3.12.10...
winget install --exact --id Python.Python.3.12 --version 3.12.10 --scope user --accept-package-agreements --accept-source-agreements
call :achar_python_base || goto erro_python
:python_ok
echo Python 3.12: %BASE%

echo.
echo ---------- 2. Tesseract OCR 5.4.0 ----------
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" goto tesseract_ok
call :tem_winget || goto sem_winget
echo Instalando o Tesseract OCR 5.4.0. Aceite o pedido de permissão de administrador do Windows.
winget install --exact --id UB-Mannheim.TesseractOCR --version 5.4.0.20240606 --accept-package-agreements --accept-source-agreements
if not exist "C:\Program Files\Tesseract-OCR\tesseract.exe" goto erro_tesseract
:tesseract_ok
echo Tesseract: C:\Program Files\Tesseract-OCR\tesseract.exe

echo.
echo ---------- 3. Pacotes do Python em ambiente_python\ ----------
set "VENV=ambiente_python\Scripts\python.exe"
if exist "%VENV%" "%VENV%" -c "import sys; sys.exit(sys.version_info[:2] != (3, 12))" >nul 2>nul && goto venv_ok
if exist ambiente_python rmdir /s /q ambiente_python
echo Criando ambiente_python\...
"%BASE%" -m venv ambiente_python || goto erro_venv
:venv_ok
"%VENV%" -m pip install --disable-pip-version-check -r requisitos.txt || goto erro_pip

echo.
echo ---------- Conferência ----------
"%VENV%" -I -X utf8 scripts\verificar_ambiente.py
if errorlevel 1 goto erro_conferencia
echo.
echo PRONTO. Agora use o 1_RODAR_TUDO.bat e depois o 2_ABRIR_REVISAO.bat.
goto fim

:achar_python_base
set "BASE="
for /f "delims=" %%p in ('py -3.12 -c "import sys; print(sys.executable)" 2^>nul') do set "BASE=%%p"
if defined BASE exit /b 0
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "BASE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if defined BASE exit /b 0
if exist "%ProgramFiles%\Python312\python.exe" set "BASE=%ProgramFiles%\Python312\python.exe"
if defined BASE exit /b 0
exit /b 1

:tem_winget
where winget >nul 2>nul && exit /b 0
exit /b 1

:sem_winget
echo *** Este Windows não tem o winget, que instala os programas sozinho. Instale à mão e rode este arquivo de novo:
echo     Python 3.12.10 ... https://www.python.org/downloads/release/python-31210/
echo     Tesseract 5.4.0 .. https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe
echo                        instalado na pasta padrão C:\Program Files\Tesseract-OCR
goto fim
:erro_python
echo *** O Python 3.12 não foi instalado. Veja a mensagem acima, ou instale à mão: https://www.python.org/downloads/release/python-31210/
goto fim
:erro_tesseract
echo *** O Tesseract não foi instalado em C:\Program Files\Tesseract-OCR. Veja a mensagem acima, ou instale à mão:
echo     https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe - na pasta padrão.
goto fim
:erro_venv
echo *** Não foi possível criar a pasta ambiente_python\. Veja a mensagem acima.
goto fim
:erro_pip
echo *** A instalação dos pacotes do Python falhou. Confira a internet e rode este arquivo de novo.
goto fim
:erro_conferencia
echo *** Ainda falta algo, marcado como FALTA acima.
goto fim

:fim
echo.
echo Pressione qualquer tecla para fechar.
pause >nul
