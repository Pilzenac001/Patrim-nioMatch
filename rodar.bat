@echo off
echo Iniciando o sistema...
:: Tenta ativar o ambiente virtual diretamente na pasta atual
call venv\Scripts\activate.bat

:: Verifica se a ativacao funcionou
if %errorlevel% neq 0 (
    echo ERRO: Nao foi possivel ativar o ambiente virtual.
    pause
    exit
)

:: Executa o programa
python main.py

:: Se o programa fechar, pausa para você ver o erro
pause