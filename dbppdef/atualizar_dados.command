#!/bin/bash
# Muda para o diretório onde o script está localizado
cd "$(dirname "$0")"

echo "=== Atualizando Banco de Dados do Dashboard ==="
python3 atualizar_dados.py

echo ""
read -p "Pressione [Enter] para fechar esta janela..."
