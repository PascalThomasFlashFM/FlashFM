#!/bin/bash
cd "$(dirname "$0")"
echo "=== Rattrapage initial de la trésorerie ==="
echo ""
echo "ATTENTION : ce rattrapage AJOUTE les montants Pennylane à ce qui est"
echo "déjà écrit dans les cellules. Si un mois est déjà rempli à la main,"
echo "ne le fais pas rentrer dans la période de rattrapage, sinon les"
echo "montants seront comptés deux fois."
echo ""
echo "Depuis quelle date veux-tu récupérer les transactions Pennylane ?"
echo "(Format AAAA-MM-JJ, par exemple 2026-03-01 pour repartir du 1er mars)"
echo ""
read -p "Date de début : " START_DATE
echo ""
python3 main.py --since "$START_DATE"
echo ""
echo "Terminé. Appuie sur Entrée pour fermer cette fenêtre."
read
