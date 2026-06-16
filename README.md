# Calottino Avionici 154

Sito statico in sola lettura per consultare dati del calottino, quote, patch, spese e statuto.

## Anteprima locale

```bash
python3 -m pip install -r requirements.txt
python3 scripts/build_site_data.py
python3 -m http.server 8000 --directory docs
```

Poi apri `http://localhost:8000`.

## Aggiornamento dati

Metodo semplice su Mac:

1. Salva il file `calottino_categoria_gestione.xlsx`.
2. Fai doppio clic su `aggiorna_sito.command`.
3. Se il Terminale non riesce a pubblicare da solo, apri GitHub Desktop e premi `Push origin`.
4. Dopo 1-2 minuti aggiorna la pagina del sito.

Il link resta sempre:

`https://angelolosito.github.io/calottino-avionici-154/`

Il comando rigenera `docs/data/site-data.json`, crea il commit e prova a pubblicare online.

## Pubblicazione gratis con link stabile

Il sito è pubblicato gratuitamente con GitHub Pages dalla cartella `docs/`.

Finché il repository e l'utente GitHub restano gli stessi, il link non cambia.
