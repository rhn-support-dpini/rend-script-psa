# elabora_Ass.py — Documentazione

Script Python che legge un export CSV da PSA/pianificazione (assignment) e produce un file Excel omonimo con tre fogli: dati grezzi, riepilogo per risorsa/milestone/progetto e riepilogo per status.

---

## Panoramica

```
CSV PSA (assignment)  →  lettura e normalizzazione  →  Excel (.xlsx)
                              ↳ foglio data              (copia integrale)
                              ↳ foglio Assignment        (aggregazione per terna)
                              ↳ foglio Assignment Status (aggregazione per status)
```

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
python elabora_Ass.py input.csv
python elabora_Ass.py -h
```

### Parametri

| Parametro | Obbligatorio | Descrizione |
|-----------|--------------|-------------|
| `input.csv` | Sì | Export CSV PSA/pianificazione. Percorso relativo alla cartella dello script o assoluto. |

### Opzioni

| Opzione | Descrizione |
|---------|-------------|
| `-h`, `--help` | Mostra l'help da riga di comando (parametri, esempi, output). |

### Output

| File | Percorso | Descrizione |
|------|----------|-------------|
| `<input>.xlsx` | Stesso percorso del CSV | Excel con fogli `data`, `Assignment`, `Assignment Status` |

In console viene stampato il percorso del file generato.

---

## Tabelle di output: struttura e calcolo valori

Le tabelle seguenti replicano il layout dei fogli Excel prodotti. Al posto dei valori numerici o testuali, ogni cella riporta **come** quel dato viene determinato dallo script.

**Convenzioni comuni**

| Elemento | Regola |
|----------|--------|
| Ore | I campi ore sono normalizzati: virgola decimale → punto; valori non numerici → 0. |
| Nomi colonna | Lo script risolve le intestazioni in modo flessibile (case-insensitive, alias alternativi). |
| Raggruppamento | Le somme sono calcolate su tutte le righe del CSV che condividono la stessa chiave di gruppo. |
| Valori descrittivi | Nei fogli aggregati, `Project: Project Name` e `Assignment: Assignment Name` usano il **primo** valore incontrato nel gruppo (`first`). |

---

### Foglio `data`

Copia integrale del CSV di input: stesse colonne, stesso ordine, stesse righe. Nessuna trasformazione o colonna aggiuntiva.

| *(ogni colonna del CSV)* | Valore grezzo dal file di input, senza aggregazioni né ricalcoli. |

---

### Foglio `Assignment`

Una riga per ogni combinazione unica di **Resource: Full Name** + **Assignment: Milestone: Milestone Name** + **Project: OPA Project Number**.

| Resource: Full Name | Assignment: Milestone: Milestone Name | Project: OPA Project Number | Assignment: Milestone: Planned Hours | Estimated Hours | Actual Hours | Project: Project Name | Assignment: Assignment Name |
|---------------------|--------------------------------------|----------------------------|--------------------------------------|-----------------|--------------|----------------------|----------------------------|
| Chiave di gruppo: nome risorsa dal sorgente. | Chiave di gruppo: milestone dell'assignment. | Chiave di gruppo: codice OPA progetto. | Somma di **Planned Hours** su tutte le righe del gruppo. | Somma di **Estimated Hours** su tutte le righe del gruppo. | Somma di **Actual Hours** su tutte le righe del gruppo. | Primo valore di **Project: Project Name** nel gruppo. | Primo valore di **Assignment: Assignment Name** nel gruppo. |

---

### Foglio `Assignment Status`

Una riga per ogni combinazione unica di **Resource: Full Name** + **Assignment: Status** + **Assignment: Forecast Category**.

| Resource: Full Name | Assignment: Status | Assignment: Forecast Category | Actual Hours |
|---------------------|-------------------|------------------------------|--------------|
| Chiave di gruppo: nome risorsa dal sorgente. | Chiave di gruppo: stato assignment (colonna K del PSA). | Chiave di gruppo: categoria forecast (colonna L del PSA). | Somma di **Actual Hours** su tutte le righe del gruppo. |

### Esempi

```bash
python elabora_Ass.py export-assignment.csv
# → export-assignment.xlsx

python elabora_Ass.py /percorso/export-assignment.csv
# → /percorso/export-assignment.xlsx
```

---

## Input: export CSV PSA

Lo script si aspetta un CSV con almeno le colonne riportate di seguito. I nomi colonna vengono risolti in modo flessibile (case-insensitive, alias alternativi).

| Colonna attesa | Alias accettati | Uso |
|----------------|-----------------|-----|
| `Resource: Full Name` | `Full Name` | Chiave di raggruppamento |
| `Assignment: Milestone: Milestone Name` | — | Chiave di raggruppamento |
| `Project: OPA Project Number` | — | Chiave di raggruppamento |
| `Assignment: Milestone: Planned Hours` | `Planned Hours` | Somma nel foglio Assignment |
| `Estimated Hours` | — | Somma nel foglio Assignment |
| `Actual Hours` | — | Somma in entrambi i fogli di riepilogo |
| `Project: Project Name` | — | Valore descrittivo nel foglio Assignment |
| `Assignment: Assignment Name` | `Assignment Name` | Valore descrittivo nel foglio Assignment |
| `Assignment: Status` | `Assignment Status` | Chiave di raggruppamento (foglio Assignment Status) |
| `Assignment: Forecast Category` | — | Chiave di raggruppamento (foglio Assignment Status) |

**Lettura CSV:** encoding provati in ordine `utf-8-sig`, `utf-8`, `latin-1`; separatore rilevato automaticamente oppure `;` / `,`. I valori numerici delle ore accettano la virgola decimale (convertita in punto).

---

## Output Excel

Per il dettaglio tabellare con descrizione del calcolo di ogni cella vedi [Tabelle di output: struttura e calcolo valori](#tabelle-di-output-struttura-e-calcolo-valori).

### Foglio `data`

Copia integrale del CSV di input: stesse colonne e stesse righe, senza trasformazioni.

### Foglio `Assignment`

Una riga per ogni combinazione unica di:

- `Resource: Full Name`
- `Assignment: Milestone: Milestone Name`
- `Project: OPA Project Number`

### Foglio `Assignment Status`

Una riga per ogni combinazione unica di:

- `Resource: Full Name`
- `Assignment: Status`
- `Assignment: Forecast Category`

---

## Errori comuni

| Messaggio | Causa | Soluzione |
|-----------|-------|-----------|
| `File non trovato` | Percorso CSV errato | Verificare percorso relativo/assoluto |
| `Il file di input deve avere estensione .csv` | Estensione diversa | Passare un file `.csv` |
| `Impossibile leggere il CSV` | Encoding o separatore non riconosciuto | Esportare di nuovo il CSV o verificare formato |
| `Colonna non trovata` | Intestazioni mancanti o diverse | Controllare i nomi colonna nell'export PSA |
