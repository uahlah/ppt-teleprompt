#!/bin/bash
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
    echo "Criando ambiente virtual Python..."
    python3 -m venv .venv
fi

source .venv/bin/activate
pip install -q -r requirements.txt

echo ""
echo "Iniciando o Teleprompter para PowerPoint..."
echo "Deixe o PowerPoint aberto e a apresentacao em modo slideshow antes de continuar."
echo "(Na primeira execucao o macOS vai pedir permissao de Automacao para controlar o PowerPoint - autorize.)"
echo ""
python3 main.py
