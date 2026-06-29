# elabora_JKAN.py — Documentazione

Script Python che elabora un export CSV di una board **MIRO**, estrae le card da **tutte le sezioni Kanban** presenti nel file, produce un file Excel con metriche temporali e mantiene uno **storico snapshot** delle colonne Kanban con report HTML interattivo (burnup, WIP, velocità).

---

## Indice

1. [Panoramica](#panoramica)
2. [Requisiti](#requisiti)
3. [Utilizzo](#utilizzo)
4. [Input: export CSV MIRO](#input-export-csv-miro)
5. [Output: file Excel](#output-file-excel)
6. [Storico `dbJKAN.csv`](#storico-dbjkancsv)
7. [Report HTML `dbJKAN.html`](#report-html-dbjkanhtml)
8. [Colonne dei fogli `data-all` e `data-check`](#colonne-dei-fogli-data-all-e-data-check)
9. [Tag temporali nella Description](#tag-temporali-nella-description)
10. [Calcolo dei giorni](#calcolo-dei-giorni)
11. [Formattazione Excel](#formattazione-excel)
12. [Fogli `stat` e `graph`](#fogli-stat-e-graph)
13. [Estendere lo script](#estendere-lo-script)
14. [Esempi](#esempi)

---

## Panoramica

Il flusso di elaborazione è:

```
CSV MIRO  →  estrazione tutte le Kanban  →  espansione tag temporali  →  Excel (.xlsx)
                                        ↘  snapshot colonne Kanban  →  dbJKAN.csv
                                                                    →  dbJKAN.html
```

Per ogni card MIRO lo script:

1. Legge i campi standard (Title, Description, Status, …).
2. Interpreta le righe della Description che iniziano con `#` come **tag temporali**.
3. Espande ogni card in una o più righe (una per tag temporale).
4. Calcola metriche aggregate (giorni in lavorazione, periodo in Progress, stima vs consuntivo).
5. Applica formattazione visiva (merge, bordi, colori) e genera fogli di riepilogo.
6. Conta le card per colonna Kanban e aggiorna lo storico `dbJKAN.csv`.
7. Rigenera `dbJKAN.html` con grafici di burnup e metriche di avanzamento.

---

## Requisiti

- Python 3
- Dipendenze: `pandas`, `openpyxl`

```bash
pip install pandas openpyxl
```

---

## Utilizzo

```bash
python elabora_JKAN.py [input.csv]
python elabora_JKAN.py -h
```

### Parametri

| Parametro | Obbligatorio | Default | Descrizione |
|-----------|--------------|---------|-------------|
| `input.csv` | No | `2026-06-06-WIP.csv` | Export CSV di una board MIRO. Percorso relativo alla cartella dello script o assoluto. **Il nome deve contenere una data** nel formato `yyyy-mm-dd` o `yyyy/mm/dd` (es. `2026-06-26-WIP.csv`). |

### Opzioni

| Opzione | Descrizione |
|---------|-------------|
| `-h`, `--help` | Mostra l'help da riga di comando (parametri, esempi, output). |

**Output:**

| File | Percorso | Descrizione |
|------|----------|-------------|
| `<input>.xlsx` | Stesso percorso del CSV | Excel con fogli `data-all`, `data-check`, `stat`, `graph` |
| `dbJKAN.csv` | Cartella dello script | Storico snapshot colonne Kanban |
| `dbJKAN.html` | Cartella dello script | Report HTML con grafici burnup e metriche |

### Esempi

```bash
python elabora_JKAN.py
# → 2026-06-06-WIP.xlsx (default)

python elabora_JKAN.py 2026-06-06-WIP.csv
# → 2026-06-06-WIP.xlsx

python elabora_JKAN.py /percorso/export-miro.csv
```

In console vengono stampati: numero sezioni kanban trovate, numero card, numero righe output, percorso Excel, conteggio snapshot per colonna Kanban, eventuali card con Status non mappato, percorsi di `dbJKAN.csv` e `dbJKAN.html`.

---

## Input: export CSV MIRO

### Struttura attesa

Il CSV può contenere più sezioni Kanban. Lo script riconosce ogni blocco con l'header:

```
Title, Description, Status, Assignee, Start Date, End Date, Estimate, Priority, Tags
```

Non viene applicato alcun filtro per nome kanban: tutte le sezioni con questo header vengono elaborate (utile quando l'export MIRO non riporta in modo affidabile i titoli delle kanban).

### Card estratte

Dopo ogni header vengono lette tutte le righe valide fino al prossimo header Kanban o alla fine della sezione. Ogni riga con almeno un campo valorizzato diventa una card; le card di sezioni diverse confluiscono nello stesso Excel.

### Campi card (colonne A–I)

| Col | Campo CSV | Note |
|-----|-----------|------|
| A | Title | Nome acronimo / card |
| B | Description | Testo libero + tag temporali `#` |
| C | Status | Stato Kanban MIRO |
| D | Assignee | Assegnatario |
| E | Start Date | Data inizio (dal CSV) |
| F | End Date | Data fine (dal CSV) |
| G | Estimate | **Sovrascritto** dal tag `# Estimate` in Description, se presente |
| H | Priority | Priorità |
| I | Tags | Tag MIRO |

---

## Output: file Excel

Il file `.xlsx` contiene quattro fogli:

| Foglio | Contenuto |
|--------|-----------|
| `data-all` | Tutte le card espansi, metriche, legenda |
| `data-check` | Come `data-all`, esclusi gli stati chiusi o da non monitorare (vedi sotto) |
| `stat` | Conteggio card per Status (da `data-all`) |
| `graph` | Grafico a barre «Tempo mancante per acronimo» (da `data-all`) |

### Foglio `data-check`

Include tutte le righe di `data-all` **tranne** le card il cui Status (col. C) è uno dei seguenti:

- `Complete`
- `Abandoned`
- `Probably dismissed / delayed to 2027`
- `new - to be verified`

Stesse colonne, merge, colori e footer di `data-all`.

---

## Storico `dbJKAN.csv`

Ad ogni esecuzione lo script salva (o aggiorna) uno **snapshot** dello stato della Kanban nella data indicata nel **nome del file di input**.

### Estrazione della data

La data viene letta dal nome file con regex `yyyy-mm-dd` o `yyyy/mm/dd`:

| Nome file | Data snapshot |
|-----------|---------------|
| `2026-06-06-WIP.csv` | `2026-06-06` |
| `export-2026/06/26.csv` | `2026-06-26` |

Se la data non è riconoscibile, lo script termina con errore.

### Struttura del CSV

| Colonna | Descrizione |
|---------|-------------|
| `data` | Data dello snapshot (`yyyy-mm-dd`) |
| `Backlog` | Card con Status mappato a Backlog |
| `In Progress` | Card in lavorazione |
| `Waiting` | Card in attesa |
| `Test in progress` | Card in carico a fabbrica / test |
| `Done` | Card completate |

Esempio:

```csv
data,Backlog,In Progress,Waiting,Test in progress,Done
2026-06-06,3,1,2,0,1
2026-06-13,2,2,1,1,2
```

### Comportamento upsert

- Se `dbJKAN.csv` **non esiste**, viene creato con intestazione e prima riga.
- Se esiste già una riga per la **stessa data**, i valori vengono **sovrascritti**.
- Altrimenti viene **aggiunta** una nuova riga; le righe sono ordinate per data.

### Mappatura Status → colonna Kanban

Il conteggio usa il campo **Status** (col. C) di ogni card. La classificazione avviene in ordine di priorità:

| Colonna | Status riconosciuti (case-insensitive) |
|---------|----------------------------------------|
| Done | contiene `complete` o `done` |
| Test in progress | contiene `test in progress`, oppure valore esatto `test` |
| In Progress | contiene `in progress` o `lavorazione` |
| Waiting | contiene `waiting` o `attesa` |
| Backlog | contiene `backlog` |

Card con Status non mappato vengono escluse dal conteggio e segnalate in console (`Card con Status non mappato: N`).

> **Nota:** la mappatura opera sul valore del campo Status MIRO, non sulle etichette delle colonne fisiche della board.

---

## Report HTML `dbJKAN.html`

Dopo ogni aggiornamento di `dbJKAN.csv`, lo script rigenera un report HTML nella cartella dello script. I grafici usano **Chart.js** (CDN, nessuna dipendenza Python aggiuntiva). Aprire il file in un browser.

### KPI in testata

| Indicatore | Significato |
|------------|-------------|
| Scope totale | Somma di tutte le colonne Kanban nell'ultimo snapshot |
| Done | Card completate nell'ultimo snapshot |
| WIP | In Progress + Waiting + Test in progress |
| Snapshot registrati | Numero di righe in `dbJKAN.csv` |

### Grafici

| Sezione | Tipo | Descrizione |
|---------|------|-------------|
| **Burnup** | Linee | Done vs scope totale (somma colonne) nel tempo |
| **Distribuzione stati** | Barre impilate | Breakdown Backlog / In Progress / Waiting / Test / Done per ogni snapshot |
| **WIP** | Linea | Andamento del lavoro in corso (escluse Backlog e Done) |
| **Velocità** | Barre | Incremento di Done rispetto allo snapshot precedente |
| **Tabella storico** | Tabella | Contenuto completo di `dbJKAN.csv` con colonna Totale |

Con un solo snapshot i grafici mostrano un punto; diventano significativi dopo più esecuzioni con file di date diverse (es. export settimanali).

---

## Colonne dei fogli `data-all` e `data-check`

### Colonne card e derivate (A–M) — merge verticali

Le colonne A–M sono unite verticalmente per ogni gruppo di righe con lo stesso **Title** (stessa card).

| Col | Titolo | Livello | Descrizione |
|-----|--------|---------|-------------|
| A–I | *(campi CSV)* | Card | Dati originali della card |
| J | InizioLavorazione(GG) | Card | Giorni dalla data nel tag `# Inizio Attivita' -` a oggi |
| K | Waiting # | Card | Giorni dall'ultimo tag `# Waiting -` a oggi; sfondo giallo se valorizzato |
| L | Totale Lavorazione | Card | Somma col. N per tag con «Lavorazione» o «in Progress»; sfondo in base a % su Estimate |
| M | Period SUM | Card | Somma col. N per tag con «in Progress» |

### Colonne per riga tag (N–O)

| Col | Titolo | Livello | Descrizione |
|-----|--------|---------|-------------|
| N | Giorni | Riga tag | Durata in giorni del singolo tag temporale |
| O | TAG Temporali | Riga tag | Testo completo del tag `#` |

### Schema visivo

```
Card «Acronimo X» (3 tag temporali)
┌─────────────────────────────────────────────────────────────┐
│ A–M (merge)  │ N (Giorni) │ O (TAG Temporali)              │
├──────────────┼────────────┼────────────────────────────────┤
│   merge      │     2      │ # 1/6/2026 - Mancano sorgenti  │
│   merge      │     1      │ # 2/6/2026 - Iniziata lav.     │
│   merge      │     5      │ # 3/6/2026 - Attesa fab.       │
└──────────────┴────────────┴────────────────────────────────┘
         ↑ bordo rosso pastello attorno all'intera card
```

---

## Tag temporali nella Description

I tag sono righe nella Description che iniziano con `#` (eventuali spazi iniziali vengono ignorati).

### Tag con prefisso strutturato

Questi tag hanno un significato specifico oltre all'espansione righe:

| Prefisso | Esempio | Effetto |
|----------|---------|---------|
| `# Estimate` | `# Estimate: 5` | Valorizza col. G; **non** genera riga espansa |
| `# Inizio Attivita' -` | `# Inizio Attivita' - 1/6/2026` | Valorizza col. J (giorni a oggi) |
| `# Waiting -` | `# Waiting - 3/6/2026` | Valorizza col. K **solo se è l'ultimo tag temporale** |

### Tag generici

Qualsiasi altra riga `#` genera una sotto-riga con colonne N e O valorizzate:

```
# 24/5/2026 - Mancano i sorgenti
# 26/5/2026 - Sorgenti a disposizione
# 27/5/2026 - Iniziata la lavorazione
```

### Formati data supportati

- `dd/mm/yyyy` — es. `3/6/2026`
- `dd/mm/yy` — es. `3/6/26`
- `yyyy-mm-dd` — es. `2026-06-03`
- ISO con timestamp — es. `2026-06-01T00:00:00.000Z` (nel CSV Start Date)

---

## Calcolo dei giorni

### Colonna N — Giorni (per tag)

Per ogni tag temporale:

1. Se il tag contiene **due date** → differenza tra seconda e prima data.
2. Se contiene **una sola data** e esiste un tag successivo → giorni fino alla data del tag successivo.
3. Se contiene **una sola data** e non ci sono tag successivi → giorni fino a **oggi**, tranne per i tag **Done** (solo data di chiusura, nessun calcolo verso oggi).

### Colonna J — InizioLavorazione(GG)

Cerca il tag `# Inizio Attivita' - <data>` e calcola:

```
oggi − data_nel_tag
```

Se il tag non è presente, la colonna resta vuota.

### Colonna K — Waiting #

Considera solo l'**ultimo** tag temporale della Description (escluso `# Estimate`):

- Se inizia con `# Waiting -` → calcola `oggi − data_nel_tag` e applica sfondo giallo pastello (`#FFFDE7`).
- Altrimenti → cella vuota, senza colore.

> **Importante:** il controllo è sul prefisso esatto `# Waiting -`, non su parole come «Attesa» nel testo del tag.

### Colonna L — Totale Lavorazione

Somma dei Giorni (col. N) delle righe il cui tag contiene **«Lavorazione»** o **«in Progress»** (case-insensitive).

#### Colorazione col. L (percentuale Working)

Calcola la percentuale `(L / Estimate) × 100` e applica lo sfondo:

| Percentuale | Colore |
|-------------|--------|
| 0 – 50% | Verde pastello (`#D9EAD3`) |
| 51 – 80% | Giallo pastello (`#FFFDE7`) |
| 81 – 100% | Rosso pastello (`#FFEBEE`) |
| > 100% | Rosso acceso (`#E53935`) |

Se manca Estimate (col. G) o Totale Lavorazione, la cella resta senza sfondo.

### Colonna M — Period SUM

Somma dei Giorni (col. N) delle righe il cui tag contiene **«in Progress»**.

### Confronto Period SUM con Estimate (col. G)

La colonna **M** viene colorata rispetto all'Estimate:

| Condizione | Colore |
|------------|--------|
| Estimate ≥ Period SUM | Verde pastello (`#D9EAD3`) |
| Estimate < Period SUM | Rosso pastello (`#FFEBEE`) |

---

## Formattazione Excel

### Allineamento

- Intestazioni e colonne C, D, G, J, K, L, M: centrate.
- Altre colonne: allineamento verticale centrale con testo a capo.

### Raggruppamento card

- Righe consecutive con lo stesso Title formano un gruppo.
- Colonne A–M: celle unite verticalmente nel gruppo.
- Intero gruppo: bordo perimetrale rosso pastello (`#E8A0A0`).

### Footer (in fondo ai fogli `data-all` e `data-check`)

1. **Riga totali** — numero card (col. A), somme di Estimate (G), Totale Lavorazione (L), Period SUM (M).
2. **Timestamp** — data/ora di generazione.
3. **Legenda colonne** — tabella con lettera, titolo e descrizione di ogni colonna derivata.

---

## Fogli `stat` e `graph`

### `stat`

Tabella con conteggio card per valore di **Status** (col. C). Status vuoti riportati come `(vuoto)`.

### `graph`

Per ogni acronimo (Title):

```
Tempo mancante = max(0, Estimate − Period SUM)
```

Viene generato un grafico a barre «Tempo mancante per acronimo» (asse X: acronimo, asse Y: giorni).

---

## Estendere lo script

Per aggiungere una nuova colonna derivata, aggiornare **sempre insieme**:

| Elemento | File / sezione |
|----------|----------------|
| Costante nome colonna | es. `NUOVA_COL = "..."` |
| Indici `COL_*` | Tutte le colonne da quella in poi si spostano |
| `COL_CARD_END` | Se la colonna è a livello card (merge A–…) |
| `COL_LAST` | Ultima colonna del foglio |
| `CENTER_COLS` | Se la colonna va centrata |
| `colonne_output()` | Ordine colonne nel DataFrame |
| `LEGENDA_COLONNE` | Voce in legenda footer |
| Docstring iniziale | Riepilogo colonne |
| Logica valorizzazione | `espandi_card_con_tag()` e/o `applica_totali_gruppo()` |
| Formattazione | Colori, merge, allineamento |

---

## Esempi

### Description di esempio

```
# Estimate: 10
# Inizio Attivita' - 1/6/2026
# 1/6/2026 - Mancano i sorgenti
# 2/6/2026 - Iniziata la lavorazione
# Waiting - 5/6/2026
```

### Risultato atteso (oggi = 8/6/2026)

| Col | Valore |
|-----|--------|
| G (Estimate) | 10 |
| J (InizioLavorazione) | 7 |
| K (Waiting #) | 3 *(giallo)* |
| L, M | calcolati dalla somma dei Giorni filtrati |
| Righe N–O | 4 righe (solo `# Estimate` è escluso dall'espansione) |

### Esecuzione

```bash
cd /percorso/rend-script-psa
python elabora_JKAN.py 2026-06-06-WIP.csv
```

Output atteso in console:

```
Sezioni kanban: 2
Card estratte: 7
Righe output: 13
Output: /percorso/rend-script-psa/2026-06-06-WIP.xlsx
Snapshot 2026-06-06: {'Backlog': 3, 'In Progress': 1, 'Waiting': 2, 'Test in progress': 0, 'Done': 1}
Database: /percorso/rend-script-psa/dbJKAN.csv
Grafici: /percorso/rend-script-psa/dbJKAN.html
```

### Storico multi-snapshot

Per costruire l'andamento nel tempo, eseguire lo script su export con date diverse nel nome:

```bash
python elabora_JKAN.py 2026-06-06-WIP.csv
python elabora_JKAN.py 2026-06-13-WIP.csv
python elabora_JKAN.py 2026-06-20-WIP.csv
```

Ogni run aggiunge (o aggiorna) una riga in `dbJKAN.csv` e rigenera `dbJKAN.html`.

---

## Riferimenti nel repository

| File | Ruolo |
|------|-------|
| `elabora_JKAN.py` | Script principale (`python elabora_JKAN.py -h` per l'help CLI) |
| `.cursor/rules/elabora-jkan.mdc` | Regole Cursor per modifiche future |
| `docs/elabora_JKAN.md` | Questa documentazione |
| `docs/elabora_progetti.md` | Documentazione dello script progetti correlato |
