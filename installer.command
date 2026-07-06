#!/bin/bash
cd "$(dirname "$0")"
echo "Installation des dépendances Python..."
pip3 install -r requirements.txt
if [ ! -f .env ]; then
  cp .env.example .env
  echo ""
  echo "Un fichier .env a été créé. Ouvre-le (Finder > Cmd+Maj+. pour voir les"
  echo "fichiers cachés) et renseigne ton token Pennylane et le chemin de ton"
  echo "fichier Excel avant de lancer 'mettre_a_jour_tresorerie.command'."
else
  echo ""
  echo "Le fichier .env existe déjà, rien à faire de ce côté."
fi
echo ""
echo "Appuie sur Entrée pour fermer cette fenêtre."
read
