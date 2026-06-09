# elabora_progetti.py — Documentazione

Script Python che legge un export Excel da PSA/pianificazione e produce un report multi-foglio con pivot settimanali, dettaglio ruoli e riepiloghi per cliente.

---

## Utilizzo

```bash
python elabora_progetti.py [cliente] [input.xlsx] [output.xlsx]
python elabora_progetti.py <input.xlsx>
python elabora_progetti.py --list-kl-combos [cliente] [input.xlsx]
python elabora_progetti.py -h
```

### Parametri posizionali (tutti opzionali)

| Parametro | Default | Descrizione |
|-----------|---------|-------------|
| `cliente` | `Intesa` | Filtro sulla colonna **Cliente** (case-insensitive). Passare `""` da shell per includere tutti i clienti. |
| `input.xlsx` | `input.xlsx` | File Excel di input (export PSA/pianificazione), nella cartella dello script o percorso assoluto. |
| `output.xlsx` | `output_elaborato.xlsx` | File Excel generato. |

**Caso particolare:** se viene passato **un solo argomento** ed è un file `.xlsx`, `.xlsm` o `.xls`, viene interpretato come `input.xlsx`; `cliente` resta `Intesa` e `output.xlsx` il default.

### Opzioni

| Opzione | Descrizione |
|---------|-------------|
| `-h`, `--help` | Mostra l'help da riga di comando (parametri, esempi, configurazione). |
| `--list-kl-combos` | Elenca le coppie uniche colonna K×L (scheduling × commit/exclude) nel sorgente e termina **senza** generare Excel. |

### Esempi

```bash
# Default: cliente Intesa, input.xlsx → output_elaborato.xlsx
python elabora_progetti.py

# Solo file di input (cliente = Intesa, output = default)
python elabora_progetti.py export-psa.xlsx

# Tutti i parametri espliciti
python elabora_progetti.py Intesa input.xlsx output_elaborato.xlsx

# Nessun filtro cliente
python elabora_progetti.py "" input.xlsx

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

---

## Riferimenti nel repository

| File | Ruolo |
|------|-------|
| `elabora_progetti.py` | Script principale |
| `script.config` | Configurazione generale (non versionata) |
| `cust.config` | Configurazione cliente (non versionata) |
| `docs/elabora_progetti.md` | Questa documentazione |
