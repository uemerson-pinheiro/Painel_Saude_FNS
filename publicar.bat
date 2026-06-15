@echo off
chcp 65001 > nul
set GIT=C:\Users\upinheiro\AppData\Local\Programs\Git\cmd\git.exe

echo.
echo =============================================
echo   Publicar alteracoes no GitHub
echo =============================================
echo.

cd /d "%~dp0"

%GIT% add .
%GIT% status

echo.
set /p MSG="Descricao da alteracao (ex: novo grafico de cobertura): "
if "%MSG%"=="" set MSG=Atualizacao do painel

%GIT% commit -m "%MSG%"
%GIT% push origin main

echo.
echo =============================================
echo   Publicado! O Streamlit Cloud atualiza
echo   automaticamente em ~1 minuto.
echo =============================================
echo.
pause
