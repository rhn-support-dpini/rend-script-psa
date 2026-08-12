# PROJECT_STATE — rend-script-psa

Ultimo aggiornamento: 2026-08-12 (sessione Cursor su branch `Cursor`).

Documento di handoff per sessioni successive: cosa c’è nel repo, come funziona, cosa è stato toccato di recente, e cosa evitare.

---

## Panoramica del progetto

Automazione per elaborare export **PSA** (assegnazioni / pianificazione) e **MIRO** (Kanban JBOSS) copiati da Google Drive, con output Excel, HTML e CSV di storico.

| Componente | Ruolo |
|------------|--------|
| `goAll` | Orchestrazione: copia Drive → elabora solo file nuovi → rinomina sorgenti |
| `elabora_progetti.py` | Export PSA → Excel multi-foglio + HTML |
| `elabora_JKAN.py` | Export MIRO Kanban → Excel + `dbJKAN.csv` + report HTML Scrum/Kanban |
| `elabora_Ass.py` | Export RHCC assignments → Excel semplificato |
| `copia_drive-PSA.sh` / `copia_drive-Miro.sh` | Copia incrementale da Drive |
| `script.config` / `cust.config` | Indici colonne input e parametri report PSA |
| `.prjIgnore` | Codici progetto esclusi dalla **Tabella di Export** (non dai calcoli) |

**Esecuzione consigliata:** `cd rend-script-psa/venv && ../goAll`

**Dati e output:** `venv/Data/` (non tracciati in git, vedi `.gitignore`).

**Repo:** `https://github.com/rhn-support-dpini/rend-script-psa.git` — branch attivo di sviluppo: `Cursor`.

---

## Implementato / modificato (cronologia recente)

### Sessione 2026-08-12 — foglio `data-export` in `elabora_JKAN.py`

Commits: `025ef4e`, `64041d9`.

1. **Colonna E — Tags**
   - Intestazione `Tags`, valore dal campo `Tags` della card MIRO (colonna 9 del CSV Kanban).
   - `COLONNE_EXPORT` = Title, Status, Start Date, End Date, Tags.

2. **Allineamento**
   - Tutte le celle del foglio `data-export` centrificate in orizzontale e verticale (`wrap_text` attivo).

3. **Bordi visibili su sfondi colorati**
   - Perimetro di ogni riga: bordo blu `medium` (`#2563EB`) — funzione `applica_bordo_riga_export`.
   - Divisori tra colonne: bordo `thin` blu chiaro (`#93C5FD`) sul bordo destro di A–D.
   - Sfondo per Status (col. B) esteso su colonne A–E.

### Prima della sessione (branch `Cursor`, ultimi ~15 commit)

**elabora_JKAN.py**

- Calcolo giorni lavorativi con fine esclusa su tag a due date; validazione struttura tag `# Working -`, `# Waiting -`, ecc.
- Palette colori `data-export` per Status (pastello + verde acceso per Acronimi done).
- Stati esclusi da `data-export` aggiornati al nuovo formato CSV aggregato.
- Report HTML omonimo dell’`.xlsx` (CFD, burnup, WIP, velocità, griglia colonne Kanban).
- `dbJKAN.csv` con upsert per data snapshot; migrazione colonne legacy.

**elabora_progetti.py**

- Foglio **VERIFICA** con diagnostica ref keys e righe omesse da `.prjIgnore`.
- `.prjIgnore`: filtro solo sulla stampa tabellare **Tabella di Export**; calcoli su tutti i dati.
- Codice interno colonna A: chiavi `CodiceInterno` in config; lookup `ref_key` normalizzato (col. M/N).
- Evidenziazione gialla righe **assegnazioni non conformi** nel foglio `dati`.
- Evidenziazione grigio chiaro per progetti in `.prjIgnore` nel dettaglio ruoli.

**goAll / infra**

- Rielaborazione PSA se manca file `RHCC-*` corrispondente.
- `script.config` e `cust.config` tracciati in git.
- Documentazione in `docs/` (JKAN, progetti, Assignment, goAll).

---

## Decisioni di architettura e convenzioni

### Struttura repo

- **Whitelist in `.gitignore`:** solo script, config, docs, regole Cursor; artefatti (xlsx, csv di dati, html generati) fuori repo.
- **Config:** `script.config` (layout colonne PSA, WeeksLimit), `cust.config` (mapping contratti / codice interno). Risoluzione path: cartella script, poi `venv/`.
- **Virtualenv:** `venv/` opzionale; `goAll` usa `venv/bin/python3` se presente.

### elabora_JKAN.py

- Input CSV MIRO con sezioni Kanban ripetute (header standard 9 colonne); **nessun filtro per nome kanban**.
- Data snapshot estratta dal **nome file** (`yyyy-mm-dd` o `yyyy/mm/dd`); obbligatoria.
- Output: `<input>.xlsx` (fogli `data-all`, `data-export`, `stat`), `<input>.html`, aggiornamento `dbJKAN.csv` nella cartella script.
- **Estimate (col. G):** sempre da tag `# Estimate` in Description, non dal CSV.
- **Espansione righe:** ogni tag `#` in Description (eccetto Estimate) → sotto-riga da colonna N; merge A–M per Title.
- **Giorni lavorativi:** lun–ven; intervalli a due date `[inizio, fine)` con fine esclusa; senza seconda data → fino a oggi incluso (eccetto Done).
- **Modifiche layout colonne `data-all`:** aggiornare insieme `COL_*`, `COL_CARD_END`, `COL_LAST`, `CENTER_COLS`, `colonne_output()`, `LEGENDA_COLONNE`, docstring iniziale.
- **Regola Cursor:** `.cursor/rules/elabora-jkan.mdc` (allineata al codice).

### elabora_progetti.py

- Filtro cliente default `Intesa`; input Excel/CSV PSA.
- Output tipico: `YYYY-MM-DD-Result-PSA.xlsx` + HTML.
- `.prjIgnore` → solo esclusione visiva/export tabella codice interno; **non** alterare pivot e totali.
- Assegnazioni: conformità controllata con evidenziazione nel foglio `dati`.

### Commit

- Stile: `feat(jkan): …`, `feat(progetti): …`, `fix(…): …`.
- Commit e push **solo su richiesta esplicita** dell’utente.

### Formattazione Excel (openpyxl)

- Colori pastello centralizzati come `PatternFill` + `Side` in costanti.
- Bordi per gruppi: `applica_bordo_gruppo` (perimetro rettangolo).
- `data-export`: `applica_bordo_riga_export` (perimetro riga + divisori colonna) — **non** riusare `applica_bordo_gruppo` se servono entrambi.

---

## Bug e lezioni apprese

| Problema | Causa | Come evitare |
|----------|--------|----------------|
| Tag temporali con Warning a runtime | Tag MIRO non conformi al pattern `# <tag> - <data> - <data> -` (es. `# Working - 30/06/2026 - 01/07/2026` senza trattino finale, spazi in date `29/06/2026` vs `29-5-2026`) | Validare su board MIRO; `valida_struttura_tag_due_date` stampa Warning ma non blocca; estendere regex solo se il formato diventa standard |
| Colonna Tags vuota in export | Nel CSV Intesa-JBOSS il campo `Tags` è spesso vuoto | La colonna è corretta; i tag operativi stanno in Description (`# …`); non confondere con TAG Temporali col. O |
| `.prjIgnore` alterava i totali | Filtro applicato troppo presto nel pipeline | Filtro solo in `formatta_tabella_codice_interno` / export tabellare; `df_per_calc` su dati completi |
| Ref key codice interno non trovata | Normalizzazione chiavi (spazi, maiuscole) incoerente tra somme e lookup | Usare `norm_ref_key` / `lookup_somma_ref`; verificare foglio VERIFICA |
| PSA non rielaborato | File `RHCC-*` già presente ma `Result-PSA` mancante | `goAll` controlla entrambi i pattern (fix `81252f6`) |
| Bordi colonne persi | `applica_bordo_gruppo` sostituisce l’intero `Border` della cella | Per `data-export` usare `applica_bordo_riga_export` con logica left/right diversa per divisori |
| Layout colonne JKAN incoerente | Modifica di una costante `COL_*` senza aggiornare legenda/merge | Checklist in regola `elabora-jkan.mdc` |
| Card Status non mappato in snapshot | Status aggregati MIRO (`Attività …`) o Abandoned non entrano in `dbJKAN` | Conteggio stampa `Card con Status non mappato: N`; rivedere `classifica_colonna_kanban` se cambiano label |

---

## Stato attuale noto (smoke test)

- `python elabora_JKAN.py 2026-07-10-Intesa-JBOSS.csv` → OK (43 card, warning su tag x1pr0).
- Snapshot 2026-07-10: Under analysis 10, backlog 11, in progress 1, waiting for fab. 1, Fab test 4, Acronimi done 0, 16 non mappate.
- `data-export`: 44 righe (header + 43 card filtrate), 5 colonne, bordi verificati via openpyxl.

---

## Prossimi passi suggeriti

### Priorità alta

1. **Allineare `docs/elabora_JKAN.md`** alla tabella `data-export` attuale (Tags, bordi, 5 colonne A–E) — la doc ancora elenca solo 4 colonne senza formattazione.
2. **Status non mappati (16 card)** — valutare mapping per `Abandoned` e stati aggregati se devono entrare in `dbJKAN` o restare solo in `data-all`/`data-export`.
3. **Tag MIRO malformati** — guida operativa per chi edita Description (formato `# Waiting - dd/mm/yyyy - dd/mm/yyyy - causale`).

### Priorità media

4. **Tracciare `PROJECT_STATE.md` in git** — attualmente non è nella whitelist `.gitignore`; aggiungere `!PROJECT_STATE.md` se deve essere versionato.
5. **Test automatici** — almeno smoke su parsing tag, `righe_export`, bordi export (openpyxl senza Excel GUI).
6. **elabora_Ass.py** — verificare allineamento con pipeline `goAll` (oggi goAll non lo invoca).

### Priorità bassa / idee

7. Larghezza colonne auto-fit su `data-export` per Tags lunghi.
8. Esportare tag da Description (non solo campo Tags CSV) in colonna E se il campo è vuoto.
9. Sincronizzare palette colori Status tra `data-export` e legenda HTML JKAN.

---

## Eredità per la sessione successiva

**Branch:** `Cursor` (allineato a `origin/Cursor` dopo push `64041d9`).

**File toccati di recente:**

- `elabora_JKAN.py` — `COLONNE_EXPORT`, `EXPORT_COL_LAST`, `EXPORT_COL_DIVIDER`, `applica_bordo_riga_export`, `formatta_foglio_export`
- `.cursor/rules/elabora-jkan.mdc` — descrizione foglio `data-export`

**Non modificati in questa sessione ma rilevanti:**

- `elabora_progetti.py` — VERIFICA, prjIgnore, codice interno
- `goAll`, `copia_drive-*.sh`
- `docs/*` — possono essere stale rispetto al codice JKAN export

**Comandi utili:**

```bash
# JKAN manuale
python elabora_JKAN.py path/to/YYYY-MM-DD-Intesa-JBOSS.csv

# Pipeline completa
cd venv && ../goAll

# PSA manuale
python elabora_progetti.py Intesa RHCC-All-Assignments-Full-....csv
```

**Contesto utente:** report per cliente Intesa / JBOSS Kanban; target scope 30 card (`SCOPE_TOTALE_CARD`); focus su leggibilità Excel (colori, bordi) oltre ai calcoli temporali.

---

*Aggiornare questo file alla fine di ogni sessione significativa o prima di un handoff a un altro agente.*
