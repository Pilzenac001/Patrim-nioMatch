@echo off
cd /d "%~dp0"
if not exist .venv (
    echo Ambiente virtual nao encontrado! Certifique-se de estar na pasta correta.
    pause
    exit /b
)
echo Ativando ambiente virtual e iniciando Patrimonio Match...
call .venv\Scripts\activate.bat
python main.py
if %errorlevel% neq 0 (
    echo.
    echo A aplicacao encerrou com erro %errorlevel%.
    pause
)
