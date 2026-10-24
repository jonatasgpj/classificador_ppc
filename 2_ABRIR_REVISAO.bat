@echo off
chcp 65001 >nul
title Kit - revisão da classificação
cd /d "%~dp0"

echo REVISÃO DA CLASSIFICAÇÃO
call :achar_python || goto fim
if not exist "resultado\revisao\fila.json" goto sem_resultado
echo O navegador abre sozinho em http://127.0.0.1:8780
echo Deixe esta janela aberta enquanto revisa; para encerrar, feche a janela.
echo Cada resposta é gravada na hora em entradas\classificacao\classificacao.csv.
echo Depois de revisar, rode o 1_RODAR_TUDO.bat para atualizar os resultados.
echo.

%PY% -I -X utf8 app_classificacao\passo04b_revisar.py --porta 8780 --abrir
goto fim

:sem_resultado
echo *** Ainda não há o que revisar: rode antes o 1_RODAR_TUDO.bat.
goto fim

:achar_python
set "PY=ambiente_python\Scripts\python.exe"
if exist "%PY%" %PY% -c "import openpyxl, pymupdf, PIL" >nul 2>nul && exit /b 0
set "PY=py -3.12"
%PY% -c "import openpyxl, pymupdf, PIL" >nul 2>nul && exit /b 0
set "PY=python"
%PY% -c "import openpyxl, pymupdf, PIL" >nul 2>nul && exit /b 0
echo *** Falta preparar o ambiente: rode antes o 0_PREPARAR_AMBIENTE.bat.
exit /b 1

:fim
echo.
echo Pressione qualquer tecla para fechar.
pause >nul
