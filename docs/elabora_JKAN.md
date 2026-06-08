# elabora_JKAN.py — Documentazione

Script Python che elabora un export CSV di una board **MIRO**, estrae le card della Kanban **JBOSS-Barison** e produce un file Excel con metriche temporali, totali e grafici di avanzamento.

---

## Indice

1. [Panoramica](#panoramica)
2. [Requisiti](#requisiti)
3. [Utilizzo](#utilizzo)
4. [Input: export CSV MIRO](#input-export-csv-miro)
5. [Output: file Excel](#output-file-excel)
6. [Colonne del foglio `data`](#colonne-del-foglio-data)
7. [Tag temporali nella Description](#tag-temporali-nella-description)
8. [Calcolo dei giorni](#calcolo-dei-giorni)
9. [Formattazione Excel](#formattazione-excel)
10. [Fogli `stat` e `graph`](#fogli-stat-e-graph)
11. [Estendere lo script](#estendere-lo-script)
12. [Esempi](#esempi)

---

## Panoramica

Il flusso di elaborazione è:

```
CSV MIRO  →  estrazione Kanban JBOSS-Barison  →  espansione tag temporali  →  Excel (.xlsx)
```

Per ogni card MIRO lo script:

1. Legge i campi standard (Title, Description, Status, …).
2. Interpreta le righe della Description che iniziano con `#` come **tag temporali**.
3. Espande ogni card in una o più righe (una per tag temporale).
4. Calcola metriche aggregate (giorni in lavorazione, periodo in Progress, stima vs consuntivo).
5. Applica formattazione visiva (merge, bordi, colori) e genera fogli di riepilogo.

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
```

| Parametro | Default | Descrizione |
|-----------|---------|-------------|
| `input.csv` | `2026-06-06-WIP.csv` | Percorso del CSV esportato da MIRO (relativo alla cartella dello script o assoluto) |

**Output:** file con lo stesso nome e percorso del CSV, estensione `.xlsx`.

Esempio:

```bash
python elabora_JKAN.py 2026-06-06-WIP.csv
# → 2026-06-06-WIP.xlsx
```

In console vengono stampati: nome Kanban trovata, numero card, numero righe output e percorso del file generato.

---

## Input: export CSV MIRO

### Struttura attesa

Il CSV può contenere più sezioni Kanban. Lo script cerca l'header:

```
Title, Description, Status, Assignee, Start Date, End Date, Estimate, Priority, Tags
```

e identifica la sezione **JBOSS-Barison** analizzando le righe precedenti l'header (ricerca di testi come `Barison`, `JBoss`, `JBOSS-Barison`).

### Card estratte

Dopo l'header vengono lette tutte le righe valide fino al prossimo header Kanban o alla fine della sezione. Ogni riga con almeno un campo valorizzato diventa una card.

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
| `data` | Dati card espansi, metriche, legenda |
| `stat` | Conteggio card per Status |
| `graph` | Grafico a barre «Tempo mancante per acronimo» |

---

## Colonne del foglio `data`

### Colonne card e derivate (A–M) — merge verticali

Le colonne A–M sono unite verticalmente per ogni gruppo di righe con lo stesso **Title** (stessa card).

| Col | Titolo | Livello | Descrizione |
|-----|--------|---------|-------------|
| A–I | *(campi CSV)* | Card | Dati originali della card |
| J | InizioLavorazione(GG) | Card | Giorni dalla data nel tag `# Inizio Attivita' -` a oggi |
| K | Waiting # | Card | Giorni dall'ultimo tag `# Waiting -` a oggi; sfondo giallo se valorizzato |
| L | Totale Lavorazione | Card | Somma col. N per tag con «Lavorazione» o «in Progress» |
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

### Colonna M — Period SUM

Somma dei Giorni (col. N) delle righe il cui tag contiene **«in Progress»**.

### Confronto con Estimate (col. G)

Le colonne L e M vengono colorate rispetto all'Estimate:

| Condizione | Colore |
|------------|--------|
| Estimate ≥ valore | Verde pastello (`#D9EAD3`) |
| Estimate < valore | Rosso pastello (`#FFEBEE`) |

---

## Formattazione Excel

### Allineamento

- Intestazioni e colonne C, D, G, J, K, L, M: centrate.
- Altre colonne: allineamento verticale centrale con testo a capo.

### Raggruppamento card

- Righe consecutive con lo stesso Title formano un gruppo.
- Colonne A–M: celle unite verticalmente nel gruppo.
- Intero gruppo: bordo perimetrale rosso pastello (`#E8A0A0`).

### Footer (in fondo al foglio `data`)

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
Kanban: JBOSS-Barison
Card estratte: 6
Righe output: 12
Output: /percorso/rend-script-psa/2026-06-06-WIP.xlsx
```

---

## Riferimenti nel repository

| File | Ruolo |
|------|-------|
| `elabora_JKAN.py` | Script principale |
| `.cursor/rules/elabora-jkan.mdc` | Regole Cursor per modifiche future |
| `docs/elabora_JKAN.md` | Questa documentazione |
