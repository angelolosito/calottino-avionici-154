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

Modifica `calottino_categoria_gestione.xlsx`, poi esegui:

```bash
python3 scripts/build_site_data.py
```

Il sito mantiene gli stessi file e lo stesso link. Al refresh della pagina caricherà il nuovo `data/site-data.json` pubblicato.

Su Mac puoi anche fare doppio clic su `aggiorna_sito.command`.

## Pubblicazione gratis con link stabile

La soluzione più semplice è GitHub Pages:

1. Crea un repository GitHub pubblico.
2. Carica questa cartella nel repository.
3. In GitHub abilita Pages usando GitHub Actions oppure la cartella `docs/`.
4. Il link sarà del tipo `https://TUO-UTENTE.github.io/NOME-REPOSITORY/`.

Finché il repository e l'utente GitHub restano gli stessi, il link non cambia.
