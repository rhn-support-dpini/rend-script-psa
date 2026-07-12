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
7. [Report HTML `<input>.html`](#report-html-inputhtml)
8. [Colonne dei fogli `data-all` e `data-export`](#colonne-dei-fogli-data-all-e-data-export)
9. [Tag temporali nella Description](#tag-temporali-nella-description)
10. [Calcolo dei giorni](#calcolo-dei-giorni)
11. [Formattazione Excel](#formattazione-excel)
12. [Foglio `stat`](#foglio-stat)
13. [Estendere lo script](#estendere-lo-script)
14. [Esempi](#esempi)

---

## Panoramica

Il flusso di elaborazione è:

```
CSV MIRO  →  estrazione tutte le Kanban  →  espansione tag temporali  →  Excel (.xlsx)
                                        ↘  snapshot colonne Kanban  →  dbJKAN.csv
                                                                    →  <input>.html
```

Per ogni card MIRO lo script:

1. Legge i campi standard (Title, Description, Status, …).
2. Interpreta le righe della Description che iniziano con `#` come **tag temporali**.
3. Espande ogni card in una o più righe (una per tag temporale).
4. Calcola metriche aggregate (giorni in lavorazione, periodo in Progress, stima vs consuntivo).
5. Applica formattazione visiva (merge, bordi, colori) e genera fogli di riepilogo.
6. Conta le card per colonna Kanban e aggiorna lo storico `dbJKAN.csv`.
7. Rigenera `<input>.html` (omonimo del file Excel) con report Scrum/Kanban e metriche di avanzamento.

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
| `<input>.xlsx` | Stesso percorso del CSV | Excel con fogli `data-all`, `data-export`, `stat` |
| `dbJKAN.csv` | Cartella dello script | Storico snapshot colonne Kanban |
| `<input>.html` | Stesso percorso del file `.xlsx` prodotto | Report HTML Scrum/Kanban (CFD, evoluzione colonne, burnup, WIP, velocità) |

### Esempi

```bash
python elabora_JKAN.py
# → 2026-06-06-WIP.xlsx (default)

python elabora_JKAN.py 2026-06-06-WIP.csv
# → 2026-06-06-WIP.xlsx

python elabora_JKAN.py /percorso/export-miro.csv
```

In console vengono stampati: numero sezioni kanban trovate, numero card, numero righe output, percorso Excel, conteggio snapshot per colonna Kanban, eventuali card con Status non mappato, percorsi di `dbJKAN.csv` e del report HTML (`.html` omonimo dell'Excel).

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

Il file `.xlsx` contiene tre fogli:

| Foglio | Contenuto |
|--------|-----------|
| `data-all` | Tutte le card espanse per tag temporali, metriche, merge, legenda |
| `data-export` | Una riga per card, sole colonne Title / Status / Start Date / End Date |
| `stat` | Conteggio card per Status (da `data-all`) |

---

## Tabelle di output: struttura e calcolo valori

Le tabelle seguenti replicano il layout dei fogli Excel prodotti. Al posto dei valori numerici o testuali, ogni cella riporta **come** quel dato viene determinato dallo script.

**Convenzioni comuni**

| Elemento | Regola |
|----------|--------|
| Giornate | Tutti i calcoli in giorni usano **giornate lavorative** (lunedì–venerdì), non giorni di calendario. |
| Espansione card | Ogni riga `#` nella Description (escluso `# Estimate`) genera una sotto-riga con colonne N–O valorizzate. |
| Livello card | Colonne A–M sono unite verticalmente per ogni gruppo di righe con lo stesso **Title**; bordo rosso pastello attorno all'intera card. |
| Livello tag | Colonne N–O hanno un valore per ogni sotto-riga (tag temporale). |
| Colonne L e M | Calcolate **dopo** la scrittura Excel, sommando i valori del gruppo card (stesso Title). |

---

### Foglio `data-all`

Una o più righe per card MIRO (espansione tag). Intestazioni riga 1; dati da riga 2.

| Title (A) | Description (B) | Status (C) | Assignee (D) | Start Date (E) | End Date (F) | Estimate (G) | Priority (H) | Tags (I) | InizioLavorazione(GG) (J) | Waiting # (K) | Totale Lavorazione (L) | Period SUM (M) | Giorni (N) | TAG Temporali (O) |
|-----------|-----------------|------------|--------------|----------------|--------------|--------------|--------------|----------|---------------------------|---------------|------------------------|----------------|------------|-------------------|
| Nome card dal CSV; chiave di merge. | Testo Description dal CSV (merge). | Status Kanban dal CSV (merge). | Assignee dal CSV (merge). | Start Date dal CSV (merge). | End Date dal CSV (merge). | Valore numerico dal tag `# Estimate` in Description; **non** dal CSV (merge). | Priority dal CSV (merge). | Tags dal CSV (merge). | Giornate lavorative dalla data nel tag `# Inizio Attivita' - <data>` a oggi; vuoto se tag assente (merge). | Se l'**ultimo** tag temporale inizia con `# Waiting -`: giornate lavorative dalla data nel tag a oggi, sfondo giallo; altrimenti vuoto (merge). | Somma giornate lavorative di tutti i tag `# Working - <start> - <end>` del gruppo; senza end date: da start a oggi; sfondo per % su Estimate (merge). | Somma colonna N delle righe tag il cui testo contiene «in Progress» (case-insensitive); colore verde/rosso vs Estimate (merge). | Per ogni riga tag: con **due date** nel tag → giornate lavorative tra le date; con **una data** e tag successivo → fino alla data del tag successivo; con **una data** senza successivo → fino a oggi (tranne tag **Done**); tag `# Working -` segue la stessa logica Working. | Testo completo del tag `#` della riga. |

**Righe footer** (dopo i dati):

| Riga | Colonna A | Colonna G | Colonna L | Colonna M |
|------|-----------|-----------|-----------|-----------|
| Totali | Numero di card (gruppi Title distinti). | Somma Estimate di tutte le card. | Somma Totale Lavorazione di tutte le card. | Somma Period SUM di tutte le card. |
| Timestamp | Data/ora di generazione (`dd/mm/yyyy HH:MM:SS`). | — | — | — |
| Legenda | Tabella lettera / titolo / descrizione di ogni colonna derivata. | — | — | — |

---

### Foglio `data-export`

Una riga per card, **senza** espansione tag. Escluse le card con Status:

- `Complete`
- `Abandoned`
- `Probably dismissed / delayed to 2027`
- `new - to be verified`

| Title | Status | Start Date | End Date |
|-------|--------|------------|----------|
| Campo Title della card dal CSV. | Campo Status della card. | Campo Start Date della card. | Campo End Date della card. |

Nessun merge, colori o footer aggiuntivi.

---

### Foglio `stat`

Una riga per ogni valore distinto di **Status** tra le card del foglio `data-all` (una riga per Title, non per sotto-riga tag).

| Status | Conteggio |
|--------|-----------|
| Valore Status della card; se vuoto → `(vuoto)`. | Numero di card con quel Status. |

---

### File `dbJKAN.csv` (storico snapshot)

Non è un foglio Excel, ma viene aggiornato a ogni esecuzione insieme al `.xlsx`.

| data | Under analysis | backlog | in progress | waiting for fab. | Fab test in progress | Acronimi done |
|------|----------------|---------|-------------|------------------|----------------------|---------------|
| Data estratta dal nome file input (`yyyy-mm-dd`). | Conteggio card il cui Status (e Description) mappa a questa colonna Kanban. | Idem. | Idem. | Idem (include attesa test/fab in Status, Description o ultimo tag). | Idem. | Idem (`Acronimi done` o `done`). |

**Scope tracciato** (usato nei grafici HTML): somma delle sei colonne Kanban per ogni riga. Card con Status non mappato o `Complete` non entrano nel conteggio.

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
| `Under analysis` | Card in fase di analisi iniziale |
| `backlog` | Card in attesa di lavorazione |
| `in progress` | Card in lavorazione |
| `waiting for fab.` | Card in attesa test/fabbrica |
| `Fab test in progress` | Card in carico a fabbrica / test |
| `Acronimi done` | Card completate (unica colonna «done») |

Esempio:

```csv
data,Under analysis,backlog,in progress,waiting for fab.,Fab test in progress,Acronimi done
2026-06-06,1,3,1,1,0,1
2026-06-13,2,2,2,1,1,2
```

### Scope tracciato e target

Lo **scope tracciato** per ogni snapshot è la **somma** delle sei colonne Kanban sopra (Under analysis + backlog + in progress + waiting for fab. + Fab test in progress + Acronimi done). Il progetto prevede **30 card** totali: nel grafico burnup compare una linea di riferimento orizzontale a 30, mentre la linea «Scope tracciato» mostra quante card risultano mappate nello snapshot. Le card completate contano **solo** in `Acronimi done`.

### Migrazione colonne legacy

Se `dbJKAN.csv` contiene ancora le intestazioni del vecchio schema (`Backlog`, `Waiting`, `In Progress`, `Waiting Test`, `Test in progress`, `Done`), lo script le converte automaticamente al caricamento:

| Colonna legacy | Nuova colonna |
|----------------|---------------|
| `Backlog` | `backlog` |
| `Waiting` | `Under analysis` |
| `In Progress` | `in progress` |
| `Waiting Test` | `waiting for fab.` |
| `Test in progress` | `Fab test in progress` |
| `Done` | `Acronimi done` |

### Comportamento upsert

- Se `dbJKAN.csv` **non esiste**, viene creato con intestazione e prima riga.
- Se esiste già una riga per la **stessa data**, i valori vengono **sovrascritti**.
- Altrimenti viene **aggiunta** una nuova riga; le righe sono ordinate per data.

### Mappatura Status → colonna Kanban

Il conteggio usa il campo **Status** (col. C) di ogni card. La classificazione avviene in ordine di priorità:

| Colonna | Status riconosciuti (case-insensitive) |
|---------|----------------------------------------|
| Acronimi done | valore esatto `Acronimi done` o `done` (case-insensitive) |
| Fab test in progress | contiene `fab test in progress` o `test in progress`, oppure valore esatto `test` |
| in progress | contiene `in progress` o `lavorazione` |
| waiting for fab. | Status con `waiting for fab` / `waiting test` / `attesa fab` / `attesa test`, oppure Description/ultimo tag con «Attesa test» / «Attesa fab.» |
| backlog | contiene `backlog` |
| Under analysis | contiene `under analysis`, oppure `waiting` / «attesa lavorazione/intesa» (esclusa attesa test/fab) |

Card con Status `Complete` (colonna distinta su MIRO, non equivalente ad Acronimi done) e altri stati chiusi non entrano nello snapshot. Card con Status non mappato vengono escluse dal conteggio e segnalate in console (`Card con Status non mappato: N`).

> **Nota:** la mappatura opera sul valore del campo Status MIRO, non sulle etichette delle colonne fisiche della board.

---

## Report HTML `<input>.html`

Dopo ogni aggiornamento di `dbJKAN.csv`, lo script genera un report HTML **omonimo del file Excel prodotto** (es. `2026-07-08-Intesa-JBOSS.csv` → `2026-07-08-Intesa-JBOSS.xlsx` → `2026-07-08-Intesa-JBOSS.html`). Ogni esecuzione produce un file HTML dedicato allo snapshot corrente, senza sovrascrivere report di export precedenti con nome diverso. I grafici usano **Chart.js** (CDN). Aprire il file in un browser.

### KPI in testata

| Indicatore | Significato |
|------------|-------------|
| Scope tracciato | Somma delle sei colonne Kanban nell'ultimo snapshot |
| Target card | 30 (card previste nel progetto) |
| Non avviato | Under analysis + Backlog nell'ultimo snapshot |
| In delivery (WIP) | in progress + Waiting for fabric + Fab. Test in progress |
| Acronimi Done | Card completate nell'ultimo snapshot |
| Snapshot | Numero di righe in `dbJKAN.csv` |

### Grafici Scrum / Kanban

| Sezione | Tipo | Descrizione |
|---------|------|-------------|
| **Cumulative Flow Diagram (CFD)** | Aree impilate | Evoluzione cumulativa di tutte le colonne Kanban — strumento Agile per colli di bottiglia |
| **Evoluzione colonne — vista comparata** | Linee multiple | Andamento simultaneo di Under analysis, Backlog, In progress, Waiting for fabric, Fab. Test in progress, Acronimi Done |
| **Evoluzione per colonna** | Griglia di 6 linee | Un grafico dedicato per ciascuna colonna del board |
| **Pipeline Scrum** | Linee | Aggregazione: non avviato (upstream) / in delivery (WIP) / completato |
| **Burnup** | Linee | Acronimi Done, scope tracciato e target 30 card |
| **Distribuzione stati** | Barre impilate | Breakdown per colonna Kanban a ogni snapshot |
| **WIP** | Linea | Andamento del lavoro in corso |
| **Velocità** | Barre | Incremento di Acronimi Done rispetto allo snapshot precedente |
| **Tabella storico** | Tabella | Contenuto completo di `dbJKAN.csv` con colonna Totale |

Sull'**asse delle ascisse** di ogni grafico, sotto la data (`dd/mm/yyyy`), viene mostrato il **numero di settimana ISO** (es. `W23`).

Con un solo snapshot i grafici mostrano un punto; diventano significativi dopo più esecuzioni con file di date diverse (es. export settimanali).

---

## Colonne dei fogli `data-all` e `data-export`

Per il dettaglio tabellare con descrizione del calcolo di ogni cella vedi [Tabelle di output: struttura e calcolo valori](#tabelle-di-output-struttura-e-calcolo-valori).

### Colonne card e derivate (A–M) — merge verticali

Le colonne A–M sono unite verticalmente per ogni gruppo di righe con lo stesso **Title** (stessa card).

| Col | Titolo | Livello | Descrizione |
|-----|--------|---------|-------------|
| A–I | *(campi CSV)* | Card | Dati originali della card |
| J | InizioLavorazione(GG) | Card | Giornate lavorative (lun–ven) dalla data nel tag `# Inizio Attivita' -` a oggi |
| K | Waiting # | Card | Giornate lavorative dall'ultimo tag `# Waiting -` a oggi; sfondo giallo se valorizzato |
| L | Totale Lavorazione | Card | Somma giornate lavorative dei tag `# Working - start - end`; sfondo in base a % su Estimate |
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
| `# Inizio Attivita' -` | `# Inizio Attivita' - 1/6/2026` | Valorizza col. J (giornate lavorative a oggi) |
| `# Waiting -` | `# Waiting - 3/6/2026` | Valorizza col. K **solo se è l'ultimo tag temporale** |
| `# Working -` | `# Working - 1/6/2026 - 5/6/2026` | Contribuisce a col. L (giornate lavorative nel periodo); senza end date: da start a oggi |

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

Per ogni tag temporale (giornate lavorative lun–ven):

1. Se il tag contiene **due date** → giornate lavorative tra prima e seconda data (inclusi).
2. Se contiene **una sola data** e esiste un tag successivo → giornate lavorative fino alla data del tag successivo.
3. Se contiene **una sola data** e non ci sono tag successivi → giornate lavorative fino a **oggi**, tranne per i tag **Done** (solo data di chiusura, nessun calcolo verso oggi).

I tag `# Working -` seguono la stessa logica quando contribuiscono a **Totale Lavorazione** (col. L).

### Colonna J — InizioLavorazione(GG)

Cerca il tag `# Inizio Attivita' - <data>` e calcola:

```
giornate lavorative (lun–ven) tra data_nel_tag e oggi
```

Se il tag non è presente, la colonna resta vuota.

### Colonna K — Waiting #

Considera solo l'**ultimo** tag temporale della Description (escluso `# Estimate`):

- Se inizia con `# Waiting -` → calcola giornate lavorative da `data_nel_tag` a oggi e applica sfondo giallo pastello (`#FFFDE7`).
- Altrimenti → cella vuota, senza colore.

> **Importante:** il controllo è sul prefisso esatto `# Waiting -`, non su parole come «Attesa» nel testo del tag.

### Colonna L — Totale Lavorazione

Somma delle **giornate lavorative** di tutti i tag `# Working - <start> - <end>` della card. Se manca la data di fine, si usano i giorni lavorativi da start a oggi.

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

### Footer (in fondo al foglio `data-all`)

1. **Riga totali** — numero card (col. A), somme di Estimate (G), Totale Lavorazione (L), Period SUM (M).
2. **Timestamp** — data/ora di generazione.
3. **Legenda colonne** — tabella con lettera, titolo e descrizione di ogni colonna derivata.

---

## Foglio `stat`

Tabella con conteggio card per valore di **Status** (col. C). Status vuoti riportati come `(vuoto)`. Una card = un Title distinto nel foglio `data-all` (non conta le sotto-righe dei tag).

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
Snapshot 2026-06-06: {'Under analysis': 1, 'backlog': 3, 'in progress': 1, 'waiting for fab.': 1, 'Fab test in progress': 0, 'Acronimi done': 1}
Database: /percorso/rend-script-psa/dbJKAN.csv
Grafici: /percorso/venv/Data/2026-06-06-WIP.html   # omonimo del .xlsx prodotto
```

### Storico multi-snapshot

Per costruire l'andamento nel tempo, eseguire lo script su export con date diverse nel nome:

```bash
python elabora_JKAN.py 2026-06-06-WIP.csv
python elabora_JKAN.py 2026-06-13-WIP.csv
python elabora_JKAN.py 2026-06-20-WIP.csv
```

Ogni run aggiunge (o aggiorna) una riga in `dbJKAN.csv` e genera il report HTML omonimo del file Excel (`.html`).

---

## Automazione con goAll

Lo script `goAll` copia da Google Drive i file `*-Intesa-JBOSS.csv`, li elabora con `elabora_JKAN.py` e rinomina i sorgenti con prefisso `x-` per segnalarli come già trattati. Vedi [goAll.md](goAll.md) per il flusso completo (PSA + MIRO).

---

## Riferimenti nel repository

| File | Ruolo |
|------|-------|
| `elabora_JKAN.py` | Script principale (`python elabora_JKAN.py -h` per l'help CLI) |
| `goAll` | Automazione copia Drive ed elaborazione incrementale |
| `.cursor/rules/elabora-jkan.mdc` | Regole Cursor per modifiche future |
| `docs/elabora_JKAN.md` | Questa documentazione |
| `docs/elabora_progetti.md` | Documentazione dello script progetti correlato |
| `docs/goAll.md` | Documentazione automazione `goAll` |
