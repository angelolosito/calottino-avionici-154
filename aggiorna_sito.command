#!/usr/bin/env bash
set -u

cd "$(dirname "$0")"

echo
echo "========================================"
echo " Aggiornamento sito Calottino Avionici"
echo "========================================"
echo

pause() {
  echo
  read -r -n 1 -s -p "Premi un tasto per chiudere."
  echo
}

BUNDLED_PY="/Users/angelo.losito/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
VENV_PY=".venv/bin/python3"

if [[ -f '~$calottino_categoria_gestione.xlsx' ]]; then
  echo "Excel sembra ancora aperto."
  echo "Se hai appena modificato il file, assicurati di aver premuto Salva prima di continuare."
  echo
  read -r -p "Premi Invio per continuare oppure chiudi questa finestra per annullare."
fi

if [[ -x "$BUNDLED_PY" ]]; then
  PYTHON="$BUNDLED_PY"
else
  if [[ ! -x "$VENV_PY" ]]; then
    echo "Preparo l'ambiente Python locale. Serve solo la prima volta."
    python3 -m venv .venv || {
      echo "Errore: non riesco a creare l'ambiente Python."
      pause
      exit 1
    }
    "$VENV_PY" -m pip install --upgrade pip
    "$VENV_PY" -m pip install -r requirements.txt || {
      echo "Errore: non riesco a installare le librerie necessarie."
      pause
      exit 1
    }
  fi
  PYTHON="$VENV_PY"
fi

echo "Leggo l'Excel e preparo il JSON del sito..."
"$PYTHON" scripts/excel_to_json.py || {
  echo
  echo "Errore: aggiornamento dati non riuscito."
  pause
  exit 1
}

echo
echo "Preparo le modifiche per GitHub..."
git add calottino_categoria_gestione.xlsx docs/data.json

if git diff --cached --quiet; then
  echo "Nessuna modifica da pubblicare."
else
  git commit -m "Aggiorna dati sito" || {
    echo
    echo "Errore: non sono riuscito a creare il commit."
    pause
    exit 1
  }
fi

echo
echo "Provo a pubblicare online..."
if git push origin main; then
  echo
  echo "Fatto. GitHub Actions convertira' l'Excel e aggiornera' il sito entro circa 1-2 minuti."
  echo "Link: https://angelolosito.github.io/calottino-avionici-154/"
else
  echo
  echo "Non riesco a pubblicare: probabilmente manca internet o GitHub non e' raggiungibile."
  echo "La modifica e' pronta sul Mac. Appena torna internet, apri GitHub Desktop e premi 'Push origin'."
  echo "Dopo il push, GitHub Actions aggiornera' il sito in circa 1-2 minuti:"
  echo "https://angelolosito.github.io/calottino-avionici-154/"
  open -a "GitHub Desktop" . >/dev/null 2>&1 || true
fi

pause
echo
