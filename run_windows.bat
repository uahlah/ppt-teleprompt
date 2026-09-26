@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv" (
    echo Criando ambiente virtual Python...
    python -m venv .venv
)

call .venv\Scripts\activate.bat
pip install -q -r requirements.txt

echo.
echo Iniciando o Teleprompter para PowerPoint...
echo Deixe o PowerPoint aberto e a apresentacao em modo slideshow (F5) antes de continuar.
echo.
python main.py

pause
