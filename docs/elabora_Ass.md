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
| `Planned Hours` | — | Somma nel foglio Assignment |
| `Estimated Hours` | — | Somma nel foglio Assignment |
| `Actual Hours` | `actual hours` | Somma in entrambi i fogli di riepilogo |
| `Project: Project Name` | — | Valore descrittivo nel foglio Assignment |
| `Assignment Name` | — | Valore descrittivo nel foglio Assignment |
| `Assignment: Status` | `status` | Chiave di raggruppamento (foglio Assignment Status) |
| `Assignment: Forecast Category` | `Forecast Category` | Chiave di raggruppamento (foglio Assignment Status) |

**Lettura CSV:** encoding provati in ordine `utf-8-sig`, `utf-8`, `latin-1`; separatore rilevato automaticamente oppure `;` / `,`. I valori numerici delle ore accettano la virgola decimale (convertita in punto).

---

## Output Excel

### Foglio `data`

Copia integrale del CSV di input: stesse colonne e stesse righe, senza trasformazioni.

### Foglio `Assignment`

Una riga per ogni combinazione unica di:

- `Resource: Full Name`
- `Assignment: Milestone: Milestone Name`
- `Project: OPA Project Number`

| Colonna | Contenuto |
|---------|-----------|
| `Resource: Full Name` | Nome risorsa |
| `Assignment: Milestone: Milestone Name` | Milestone |
| `Project: OPA Project Number` | Codice progetto OPA |
| `Planned Hours` | Somma delle ore pianificate |
| `Estimated Hours` | Somma delle ore stimate |
| `Actual Hours` | Somma delle ore consuntivate |
| `Project: Project Name` | Nome progetto (primo valore del gruppo) |
| `Assignment Name` | Nome assignment (primo valore del gruppo) |

### Foglio `Assignment Status`

Una riga per ogni combinazione unica di:

- `Resource: Full Name`
- `Assignment: Status`
- `Assignment: Forecast Category`

| Colonna | Contenuto |
|---------|-----------|
| `Resource: Full Name` | Nome risorsa |
| `Assignment: Status` | Stato assignment |
| `Assignment: Forecast Category` | Categoria forecast |
| `Actual Hours` | Somma delle ore consuntivate |

---

## Errori comuni

| Messaggio | Causa | Soluzione |
|-----------|-------|-----------|
| `File non trovato` | Percorso CSV errato | Verificare percorso relativo/assoluto |
| `Il file di input deve avere estensione .csv` | Estensione diversa | Passare un file `.csv` |
| `Impossibile leggere il CSV` | Encoding o separatore non riconosciuto | Esportare di nuovo il CSV o verificare formato |
| `Colonna non trovata` | Intestazioni mancanti o diverse | Controllare i nomi colonna nell'export PSA |
