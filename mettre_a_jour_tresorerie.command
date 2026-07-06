#!/bin/bash
cd "$(dirname "$0")"
echo "Mise à jour de la trésorerie en cours..."
echo ""
python3 main.py "$@"
echo ""
echo "Terminé. Appuie sur Entrée pour fermer cette fenêtre."
read
