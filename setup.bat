@echo off
echo Criando ambiente virtual...
python -m venv venv
call venv\Scripts\activate

echo Instalando dependencias...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo Criando pastas de trabalho...
if not exist "entrada" mkdir entrada
if not exist "saida" mkdir saida
if not exist "logs" mkdir logs

echo SETUP CONCLUIDO!
pause