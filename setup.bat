@echo off
echo Configurando o ambiente virtual...
python -m venv venv
call venv\Scripts\activate
echo Instalando as dependencias...
pip install -r requirements.txt
echo Configuracao concluida!
pause