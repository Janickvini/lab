#!/bin/bash
cd "$(dirname "$0")"
echo "========================================================"
echo "  SFCSED - Sincronização de Sementes de Futuro e Scopus"
echo "========================================================"
echo ""
python3 atualizar_dados.py
echo ""
echo "--------------------------------------------------------"
echo "Concluído! Pressione qualquer tecla para fechar."
read -n 1 -s
