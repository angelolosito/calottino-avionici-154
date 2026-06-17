# Calottino Avionici 154

Sito statico in sola lettura per consultare dati del calottino, quote, patch, spese e statuto.

## Anteprima locale

```bash
python3 -m pip install -r requirements.txt
python3 scripts/excel_to_json.py
python3 -m http.server 8000 --directory docs
```

Poi apri `http://localhost:8000`.

## Aggiornamento dati

Metodo semplice su Mac:

1. Salva il file `calottino_categoria_gestione.xlsx`.
2. Pubblica la modifica con GitHub Desktop: `Commit` e poi `Push origin`.
3. GitHub Actions converte l'Excel in `docs/data.json` e pubblica `docs/` su GitHub Pages.
4. Dopo 1-2 minuti aggiorna la pagina del sito.

Il link resta sempre:

`https://angelolosito.github.io/calottino-avionici-154/`

Il sito legge `docs/data.json`, generato automaticamente dal workflow `.github/workflows/deploy-pages.yml`.

Se non hai internet, il salvataggio Excel resta solo sul Mac e GitHub Actions non può partire. Appena torna la connessione, fai `Commit` e `Push origin` da GitHub Desktop; in alternativa apri la tab `Actions` del repository e rilancia manualmente il workflow `Aggiorna sito da Excel`.

## Pubblicazione gratis con link stabile

Il sito è pubblicato gratuitamente con GitHub Pages tramite GitHub Actions e l'artifact della cartella `docs/`.

Finché il repository e l'utente GitHub restano gli stessi, il link non cambia.

Impostazione una tantum su GitHub: `Settings` -> `Pages` -> `Build and deployment` -> `Source` -> `GitHub Actions`.
