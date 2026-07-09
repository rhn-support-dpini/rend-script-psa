# goAll — Automazione copia Drive ed elaborazione

Script bash che orchestra la copia incrementale degli export **PSA** e **MIRO** da Google Drive nella cartella dati locale, lanciando gli script Python solo sui file **appena copiati**. Al termine dell'elaborazione rinomina i sorgenti per marcarli come già trattati ed evitare ricopie e rielaborazioni ai giri successivi.

---

## Panoramica

```
Google Drive                          rend-script-psa/venv/Data/
├── DvPFilter-RHCC-All-Assignments-Full/   ──copia──►  DvPFilter-…  ──elabora_progetti.py──►  RHCC-…
│                                                        └──►  YYYY-MM-DD-Result-PSA.xlsx
└── Miro/                                  ──copia──►  YYYY-MM-DD-Intesa-JBOSS.csv
    └── *-Intesa-JBOSS.csv                               ──elabora_JKAN.py──►  x-YYYY-MM-DD-Intesa-JBOSS.csv
                                                         └──►  .xlsx, <input>.html, dbJKAN.csv
```

| Fase | Script ausiliario | Elaborazione Python | Rinomina post-successo |
|------|-------------------|---------------------|------------------------|
| 1–2 PSA | `copia_drive-PSA.sh` | `elabora_progetti.py` | `DvPFilter-` → *(rimosso)* |
| 3–4 MIRO | `copia_drive-Miro.sh` | `elabora_JKAN.py` | prefisso `x-` |

---

## Layout directory

| Percorso | Contenuto |
|----------|-----------|
| `rend-script-psa/` | Script (`goAll`, `elabora_*.py`, `copia_drive-*.sh`), config (`script.config`, `cust.config`) |
| `rend-script-psa/venv/` | Virtualenv Python (opzionale ma consigliato) |
| `rend-script-psa/venv/Data/` | Dati copiati, output Excel/HTML, sorgenti rinominati |

Lo script `goAll` risiede in `rend-script-psa/`; l'esecuzione consigliata è dalla sottocartella `venv`:

```bash
cd rend-script-psa/venv
../goAll
```

Funziona anche con `./goAll` dalla root del progetto: i percorsi sono risolti rispetto alla posizione dello script.

---

## Utilizzo

```bash
cd rend-script-psa/venv
../goAll
```

### Variabili d'ambiente opzionali

| Variabile | Default | Descrizione |
|-----------|---------|-------------|
| `TARGET_DIR` | `<script>/venv/Data` | Cartella destinazione dati e risultati |
| `CLIENTE` | `Intesa` | Filtro cliente passato a `elabora_progetti.py` |
| `DRIVE_DIR` | *(vedi sotto)* | Sovrascrive la sorgente Drive per `copia_drive-PSA.sh` o `copia_drive-Miro.sh` quando invocati singolarmente |
| `LOCAL_DIR` | `<script>/Data` o `TARGET_DIR` | Destinazione copia (impostata automaticamente da `goAll`) |

### Percorsi Google Drive (default macOS)

| Script | Cartella sorgente |
|--------|-------------------|
| `copia_drive-PSA.sh` | `~/Library/CloudStorage/GoogleDrive-…/Il mio Drive/Customers/DvPFilter-RHCC-All-Assignments-Full` |
| `copia_drive-Miro.sh` | `~/Library/CloudStorage/GoogleDrive-…/Il mio Drive/Customers/Miro` |

Le due cartelle sono **parallele** sotto `Customers/`.

---

## Flusso PSA (fasi 1–2)

### Copia

`copia_drive-PSA.sh` copia i file `DvPFilter-RHCC-All-Assignments-Full*` da Drive.

**Salta la copia** se in `Data/` esiste già:
- lo stesso nome (`DvPFilter-RHCC-All-Assignments-Full-2026-07-08-…`), oppure
- la versione elaborata senza prefisso (`RHCC-All-Assignments-Full-2026-07-08-…`).

### Elaborazione

Su ogni file `DvPFilter-*` **appena copiato** (assente prima della copia):

```bash
python elabora_progetti.py "$CLIENTE" "<file>" "<TARGET_DIR>/YYYY-MM-DD-Result-PSA.xlsx"
```

- La data `YYYY-MM-DD` è estratta dal nome file dopo `Full-`.
- In caso di successo il sorgente viene rinominato rimuovendo il prefisso `DvPFilter-`.
- Se esiste già la versione rinominata, eventuali copie duplicate `DvPFilter-*` vengono eliminate.

---

## Flusso MIRO / JKAN (fasi 3–4)

### Copia

`copia_drive-Miro.sh` copia solo i file il cui nome corrisponde al pattern `*-Intesa-JBOSS.csv` (es. `2026-07-08-Intesa-JBOSS.csv`).

**Salta la copia** se in `Data/` esiste già:
- lo stesso nome, oppure
- la versione elaborata con prefisso `x-` (es. `x-2026-07-08-Intesa-JBOSS.csv`).

### Elaborazione

Su ogni file `*-Intesa-JBOSS.csv` **appena copiato** (senza prefisso `x-`, assente prima della copia):

```bash
python elabora_JKAN.py "<file copiato in Data>"
```

- Output: `.xlsx` omonimo, `<input>.html` omonimo dell'Excel, `dbJKAN.csv` (storico, cartella script).
- In caso di successo il CSV sorgente viene rinominato con prefisso `x-`.
- Se esiste già `x-<nome>`, eventuali copie duplicate vengono eliminate.

---

## Logica incrementale

Per entrambi i flussi `goAll` tiene un elenco dei file già presenti in `Data/` **prima** della copia da Drive. Solo i file che compaiono dopo la copia e non erano nell'elenco vengono elaborati.

| Tipo | Nome su Drive | Nome dopo elaborazione | Segnale “già trattato” |
|------|---------------|------------------------|-------------------------|
| PSA | `DvPFilter-RHCC-All-Assignments-Full-…` | `RHCC-All-Assignments-Full-…` | assenza prefisso `DvPFilter-` |
| JKAN | `2026-07-08-Intesa-JBOSS.csv` | `x-2026-07-08-Intesa-JBOSS.csv` | prefisso `x-` |

Al secondo lancio di `goAll` i file già gestiti non vengono ricopiati né rigenerati.

---

## Requisiti

- Bash, `rsync`
- Python 3 con dipendenze di `elabora_progetti.py` e `elabora_JKAN.py` (`pandas`, `openpyxl`)
- Google Drive montato e accessibile ai percorsi configurati
- File di config PSA in `rend-script-psa/` o `rend-script-psa/venv/` (`script.config`, `cust.config`)

---

## Script ausiliari

Gli script di copia possono essere usati anche singolarmente:

```bash
LOCAL_DIR=/path/to/Data ./copia_drive-PSA.sh
LOCAL_DIR=/path/to/Data ./copia_drive-Miro.sh
```

---

## Riferimenti nel repository

| File | Ruolo |
|------|-------|
| `goAll` | Orchestratore principale |
| `copia_drive-PSA.sh` | Copia incrementale export PSA |
| `copia_drive-Miro.sh` | Copia incrementale export MIRO JKAN |
| `elabora_progetti.py` | Report PSA — vedi [elabora_progetti.md](elabora_progetti.md) |
| `elabora_JKAN.py` | Report MIRO Kanban — vedi [elabora_JKAN.md](elabora_JKAN.md) |
| `docs/goAll.md` | Questa documentazione |
