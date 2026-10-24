@echo off
chcp 65001 >nul
title Kit - rodar tudo
cd /d "%~dp0"

echo RODAR TUDO
echo Refaz o kit do zero: e-MEC, texto dos PDFs com OCR, busca dos termos, classificação automática por regras,
echo consolidação da classificação revisada, checagens e análise do capítulo 4. Leva cerca de 11 minutos.
echo A pasta resultado\ é apagada e refeita. A classificação revisada, em entradas\classificacao\classificacao.csv, não é alterada.
echo.
call :achar_python || goto fim
echo Pressione qualquer tecla para começar, ou feche esta janela para cancelar.
pause >nul
echo.

%PY% -I -X utf8 rodar_tudo.py
if errorlevel 1 goto erro
echo.
echo PRONTO. Os resultados estão em resultado\ e a revisão pode ser aberta com o 2_ABRIR_REVISAO.bat.
goto fim

:erro
echo.
echo *** PAROU COM ERRO. Veja a mensagem acima.
goto fim

:achar_python
set "PY=ambiente_python\Scripts\python.exe"
if exist "%PY%" %PY% -c "import openpyxl, pymupdf, matplotlib, numpy, PIL" >nul 2>nul && exit /b 0
set "PY=py -3.12"
%PY% -c "import openpyxl, pymupdf, matplotlib, numpy, PIL" >nul 2>nul && exit /b 0
set "PY=python"
%PY% -c "import openpyxl, pymupdf, matplotlib, numpy, PIL" >nul 2>nul && exit /b 0
echo *** Falta preparar o ambiente: rode antes o 0_PREPARAR_AMBIENTE.bat.
exit /b 1

:fim
echo.
echo Pressione qualquer tecla para fechar.
pause >nul
