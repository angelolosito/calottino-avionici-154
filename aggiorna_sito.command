#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
python3 scripts/build_site_data.py

echo
echo "Sito aggiornato nella cartella docs/."
echo "Se il sito è su GitHub Pages, pubblica le modifiche sul repository per aggiornare il link pubblico."
echo
read -r -n 1 -s -p "Premi un tasto per chiudere."
echo
