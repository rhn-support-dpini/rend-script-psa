# elabora_progetti.py — Documentazione

Script Python che legge un export Excel o CSV da PSA/pianificazione e produce un report multi-foglio con pivot settimanali, dettaglio ruoli e riepiloghi per cliente.

---

## Utilizzo

```bash
python elabora_progetti.py [cliente] [input] [output.xlsx]
python elabora_progetti.py <input>
python elabora_progetti.py --list-kl-combos [cliente] [input]
python elabora_progetti.py -h
```

### Parametri posizionali (tutti opzionali)

| Parametro | Default | Descrizione |
|-----------|---------|-------------|
| `cliente` | `Intesa` | Filtro sulla colonna **Cliente** (case-insensitive). Passare `""` da shell per includere tutti i clienti. |
| `input` | `input.xlsx` | File di input (`.xlsx`, `.xlsm`, `.xls` o `.csv`), export PSA/pianificazione, nella cartella dello script o percorso assoluto. |
| `output.xlsx` | `output_elaborato.xlsx` | File Excel generato. |

**Caso particolare:** se viene passato **un solo argomento** ed è un file `.xlsx`, `.xlsm`, `.xls` o `.csv`, viene interpretato come input; `cliente` resta `Intesa` e `output.xlsx` il default.

### Opzioni

| Opzione | Descrizione |
|---------|-------------|
| `-h`, `--help` | Mostra l'help da riga di comando (parametri, esempi, configurazione). |
| `--list-kl-combos` | Elenca le coppie uniche colonna K×L (scheduling × commit/exclude) nel sorgente e termina **senza** generare Excel. |

### Esempi

```bash
# Default: cliente Intesa, input.xlsx → output_elaborato.xlsx
python elabora_progetti.py

# Solo file Excel di input (cliente = Intesa, output = default)
python elabora_progetti.py export-psa.xlsx

# Solo file CSV di input (stessi dati, formato diverso)
python elabora_progetti.py export-psa.csv

# Tutti i parametri espliciti
python elabora_progetti.py Intesa input.xlsx output_elaborato.xlsx

# Nessun filtro cliente, input CSV
python elabora_progetti.py "" input.csv

# Analisi combinazioni K×L senza output Excel
python elabora_progetti.py --list-kl-combos Intesa input.xlsx
```

---

## Configurazione

File nella cartella dello script (o percorso assoluto dove previsto dal codice):

| File | Ruolo |
|------|-------|
| `script.config` | Impostazioni generiche (sempre letto). |
| `cust.config` | Configurazione specifica del cliente. |

---

## Output Excel

Il file di output contiene i fogli:

| Foglio | Contenuto |
|--------|-----------|
| `dati` | Dati sorgente arricchiti con colonne derivate |
| `progetti` | Riepilogo contratti: giorni consuntivati vs. riscattati |
| `Riepilogo Settimanale` | Pivot actual/estimated per attività e settimana |
| `Dettaglio Ruoli` | Pivot estimated con breakdown per ruolo/milestone |
| `Tentative` | Stati ≠ Scheduled/Commit e aggregazioni miste K×L |
| `Tabella di Export` | Riepilogo giornate per codice ordine |

Viene anche generato un file **HTML** omonimo (stesso nome dell'output Excel, estensione `.html`) con tabelle e grafici Chart.js.

---

## Milestone Non-Billable Labor

Le righe con milestone **Non-Billable Labor** (colonna ruolo/milestone del sorgente, confronto case-insensitive) sono **sempre tracciate** nei fogli di dettaglio:

- `dati`, pivot Actual/Estimated, `Riepilogo Settimanale`, `Dettaglio Ruoli`, `Tentative`
- colonna **J** della `Tabella di Export` (giornate consuntivate totali)

Non entrano invece nel calcolo dei **giorni disponibili / rimanenti**:

| Dove | Comportamento |
|------|----------------|
| Foglio `progetti` — colonne *Days Used* (G/H) e *Days remaining* (I/J) | Escluse dalla detrazione sui giorni riscattati |
| `Tabella di Export` — colonna K (residuo) | In J compaiono tutte le giornate; in K si detrae solo la parte billable |
| Grafici HTML — trend giorni rimasti (contratti RH e voci Intesa) | Cumulo solo ore billable |

---

## Elaborazione dati sorgente

Dopo la lettura dell'Excel, lo script espande la **colonna assegnazione** (indice 2, colonna C del sorgente PSA) in sette campi separati dal carattere `|`:

| Campo derivato | Contenuto |
|----------------|-----------|
| `Nome risorsa` | Nome della risorsa |
| `OPA@profilo` | Codice OPA e profilo |
| `Cliente` | Cliente |
| `Sotto progetto` | Sotto-progetto |
| `Riferimento tabella 1` | Primo codice numerico del campo riferimento |
| `Sotto Riferimento tabella 1` | Secondo codice, se presente (formato `6&4`) |
| `Commento` | Commento libero |

Il campo riferimento nell'assegnazione può contenere uno o due valori numerici separati da `&` (es. `6&4` o `6 & 4`): il primo va in **Riferimento tabella 1**, il secondo in **Sotto Riferimento tabella 1**.

---

## Tabella di Export

Il foglio **Tabella di Export** (e la sezione omonima nel report HTML) riepiloga le giornate consuntivate per codice ordine, incrociando le righe del sorgente con le voci definite in `script.config` / `cust.config` (`Export3`, `Export4`, …).

**Regola di esclusione:** le righe con **Riferimento tabella 1** vuoto, mancante o pari a **0** restano nel flusso normale (pivot, foglio `progetti`, riepiloghi, Tentative, grafici) ma **non** entrano nelle somme del tab Export — né per il riferimento principale né per il sotto-riferimento di quella riga.

Le righe con riferimento valorizzato e diverso da zero contribuiscono alle colonne J/K (giornate consuntivate e residuo) incrociando sia `Riferimento tabella 1` sia `Sotto Riferimento tabella 1` con l'ultimo campo di ciascuna voce `ExportN` in config.

- **Colonna J:** somma di tutte le ore consuntivate (in giornate), incluse quelle con milestone **Non-Billable Labor**.
- **Colonna K:** giornate acquisite (config) meno solo le ore **billable**; le ore Non-Billable Labor non riducono il residuo.

---

## Riferimenti nel repository

| File | Ruolo |
|------|-------|
| `elabora_progetti.py` | Script principale |
| `script.config` | Configurazione generale (non versionata) |
| `cust.config` | Configurazione cliente (non versionata) |
| `docs/elabora_progetti.md` | Questa documentazione |
