@echo off
echo Criando ambiente virtual...
python -m venv venv
echo Ativando ambiente e instalando dependencias...
call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
echo.
echo SETUP CONCLUIDO! Voce ja pode usar o rodar.bat
pause