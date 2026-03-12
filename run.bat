@echo off
REM Script para executar o Controle_DDU ativando o venv automaticamente
chcp 65001 >nul
setlocal enabledelayedexpansion

echo Ativando ambiente virtual...
call .venv\Scripts\activate.bat

echo Iniciando Controle_DDU...
python Controle_documentos.py

pause
