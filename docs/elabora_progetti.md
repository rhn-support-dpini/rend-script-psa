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

## Tabelle di output: struttura e calcolo valori

Le tabelle seguenti replicano il layout dei fogli Excel prodotti. Al posto dei valori numerici o testuali, ogni cella riporta **come** quel dato viene determinato dallo script.

**Convenzioni comuni**

| Elemento | Regola |
|----------|--------|
| Ore → giornate | Tutte le ore (Actual / Estimated) sono convertite in giornate dividendo per **8**. |
| Filtro settimane | Se `WeeksLimit=yes` in `script.config`, i calcoli usano solo righe con `sett_calc` in `[StartWeek, EndWeek)`; le note compaiono in fondo ai fogli `progetti` e `Tabella di Export`. |
| Filtro cliente | Se `cliente` è valorizzato, tutte le tabelle usano solo righe con **Cliente** corrispondente (case-insensitive). |
| `.prjIgnore` | Esclusioni solo in **stampa tabellare** (case-sensitive): **nome progetto PSA** → riga omessa nel tab `progetti` (e tabella HTML Progetti); **codice interno** (col. A, `cust.config`) → riga omessa nel tab `Tabella di Export` e sezione HTML omonima. Calcoli, pivot, Riepilogo, Dettaglio Ruoli, Tentative e foglio `dati` includono sempre tutte le righe sorgente. |
| Duplicazione tabelle | I fogli `progetti` e `Tabella di Export` contengono **due copie** della stessa tabella: la prima con decimali, la seconda con giornate arrotondate all'intero più vicino (colonne G/H nel foglio progetti; colonne I/J/K nell'Export). |

---

### Foglio `dati`

Tabella a colonne: tutte le colonne del file sorgente PSA/CSV più le colonne derivate dall'assegnazione (colonna C del sorgente).

| Colonna sorgente (indice default) | Calcolo / origine |
|-----------------------------------|-------------------|
| Progetto (`ColIdxProgetto` = 0) | Valore grezzo dal file di input. |
| Colonna indice 1 | Seconda colonna del sorgente (es. OPA / identificativo risorsa). |
| Assegnazione (indice 2) | Campo pipe-separated del PSA; non modificato, ma espanso nelle colonne derivate sotto. |
| Ruolo / milestone (`ColIdxRuolo` = 4) | Valore grezzo dal sorgente. |
| Estimated Hours (`ColIdxStimato` = 7) | Normalizzato: virgola → punto; non numerico → 0. |
| Actual Hours (`ColIdxEffettivo` = 8) | Come Estimated Hours. |
| Periodo (`ColIdxPeriodo` = 9) | Valore grezzo (es. `CY2026-W14`). |
| Colonne K, L (indici 10–11) | Stato schedulazione e Forecast Category; usate per colorazione e foglio `Tentative`. |

| Colonna derivata | Calcolo / origine |
|------------------|-------------------|
| `Nome risorsa` | Primo segmento del campo assegnazione (`\|`). |
| `OPA@profilo` | Secondo segmento; usato per distinguere PM/PC (`@pm`, `@pc`) nei calcoli del foglio `progetti`. |
| `Cliente` | Terzo segmento. |
| `Sotto progetto` | Quarto segmento. |
| `Riferimento tabella 1` | Primo codice numerico del quinto segmento (prima di `&`). |
| `Sotto Riferimento tabella 1` | Secondo codice numerico, se presente (`6&4`). |
| `Commento` | Sesto segmento. |

In fondo al foglio: riga **Versione script** con data e hash dell'ultimo commit git.

---

### Foglio `progetti`

Intestazioni fisse (righe 3–10) da `cust.config` (`Intestazione5`, `Intestazione6`, `Intestazione8a`–`10a`). Una riga per ogni **progetto unico** presente nel sorgente (dopo filtri).

| Contract name (A) | OPA Number (B) | Opportunity (C) | End Date (D) | Days redempted PM (E) | Days redempted Consulting (F) | Days Used PM (G) | Days Used Consulting (H) | Days remaining PM (I) | Days remaining Consulting (J) | Riferimento (K) |
|-------------------|----------------|-----------------|--------------|----------------------|------------------------------|------------------|---------------------------|----------------------|------------------------------|-----------------|
| Nome progetto dal sorgente (`ColIdxProgetto`). | Valore della **seconda colonna** del sorgente sulla prima riga di quel progetto. | `Opportunity<N>` in `cust.config`, dove `<N>` è il suffisso del contratto con `ContractName<N>` uguale al nome progetto. | `EndDate<N>` in `cust.config` (stesso abbinamento). | Prima parte di `DaysRedempted<N>` (prima della virgola). | Seconda parte di `DaysRedempted<N>` (dopo la virgola). | Somma **Actual Hours** del progetto con `OPA@profilo` contenente `@pm` o `@pc`, **escluse** righe Non-Billable Labor, ÷ 8. | Somma **Actual Hours** del progetto con profilo diverso da PM/PC, **escluse** righe Non-Billable Labor, ÷ 8. | `E − G` (giorni riscattati PM meno giorni usati PM). | `F − H` (giorni riscattati Consulting meno giorni usati Consulting). | `Riferimento tabella 1` dalla prima riga del progetto nel sorgente arricchito. |

**Colori riga (solo formattazione):** colonne A–D in base alla vicinanza di End Date; I se residuo PM &lt; 5 (rosso) o &lt; 40 (giallo); J se residuo Cons. &lt; 10 (rosso) o &lt; 80 (giallo).

---

### Foglio `Riepilogo Settimanale`

Due sezioni in sequenza verticale, stesso schema colonne. Titoli: `ACTUAL HOURS (GIORNATE) - <data>` e `ESTIMATED HOURS (GIORNATE) - <data>`.

| Project: Project Name | Sotto progetto | Milestone | Commento | Resource: Full Name | Riferimento tabella 1 | *Settimana 1* | *Settimana 2* | … | *Settimana N* | TOTALE RIGA |
|-----------------------|----------------|-----------|----------|---------------------|----------------------|---------------|---------------|---|---------------|-------------|
| Chiave pivot: nome progetto. | Chiave pivot: sotto-progetto dall'assegnazione. | Chiave pivot: ruolo/milestone del sorgente. | Valori distinti di **Commento** per la stessa chiave attività, separati da virgola e a capo. | Nomi risorsa distinti (`Resource: Full Name` o `Nome risorsa`) per la stessa chiave, separati da virgola e a capo. | Chiave pivot: `Riferimento tabella 1`. | Somma ore (Actual o Estimated, a seconda della sezione) per quella attività in quel periodo, ÷ 8. | Idem per la settimana successiva. | … | Idem. | Somma orizzontale di tutte le colonne settimana della riga. |

| Riga finale | Calcolo |
|-------------|---------|
| **TOTALE SETTIMANA** | Per ogni colonna settimana: somma verticale dei valori numerici delle righe dati della sezione. |

**Sezione Actual:** righe con **TOTALE RIGA** = 0 sono omesse. **Settimana corrente:** colonna evidenziata in verde chiaro.

---

### Foglio `Dettaglio Ruoli`

Stessa logica pivot della sezione **Estimated** del Riepilogo, con intestazione `DETTAGLIO RUOLI - ESTIMATED HOURS - <data>`.

| Project: Project Name | Sotto progetto | Milestone | Commento | Resource: Full Name | Riferimento Interno | *Settimana 1* | … | TOTALE RIGA | *(vuota)* | Actual Hours (Giornate) |
|-----------------------|----------------|-----------|----------|---------------------|---------------------|---------------|---|-------------|-----------|-------------------------|
| Come Riepilogo Estimated. | Come Riepilogo. | Come Riepilogo. | Come Riepilogo. | Come Riepilogo. | `Riferimento tabella 1` (rinominato). | Ore stimate per attività/settimana, ÷ 8. | … | Somma orizzontale settimane. | Colonna vuota di separazione. | Somma **Actual Hours** per progetto + sotto progetto + milestone + riferimento (tutto il periodo, non per settimana), ÷ 8. |

**Righe 3–4 (sopra le settimane):** riga 3 = intervallo lun–dom della settimana ISO; riga 4 = nome mese (o `mese1-mese2` se a cavallo).

**Dalla settimana corrente in poi:** ogni cella settimana è colorata in base alla coppia dominante (Status K × Forecast L) tra le righe sorgente che confluiscono in quella cella; in caso di coppie miste vince quella con **maggior somma di ore stimate**.

---

### Foglio `Tentative`

Una riga per ogni record che **non** è `Scheduled` + `Commit`, oppure che appartiene a un bucket (progetto × sotto progetto × milestone × riferimento × settimana) con **più coppie K×L** diverse. Righe con ore stimate = 0 escluse.

| Project: Project Name | Resource: Full Name | Estimated Hours | Time Period: Time Period Name | Assignment: Status | Assignment: Forecast Category | Sotto progetto | Riferimento tabella 1 | Sotto Riferimento tabella 1 | Commento |
|-----------------------|---------------------|-----------------|-------------------------------|--------------------|------------------------------|----------------|----------------------|----------------------------|----------|
| Nome progetto dal sorgente. | `Resource: Full Name` o `Nome risorsa`. | Ore stimate della riga, ÷ 8 (mostrate come giornate). | Periodo (es. `CY2026-W14`). | Colonna K del sorgente (stato assegnazione). | Colonna L del sorgente (forecast). | Da assegnazione. | Da assegnazione. | Da assegnazione (`&` nel campo riferimento). | Da assegnazione. |

**Colore riga:** pastel in base alla coppia normalizzata (K, L) della singola riga (stessa palette del Dettaglio Ruoli).

---

### Foglio `Tabella di Export`

Titolo riga 1: `CodiceInternoTitolo` da `cust.config` + data odierna. Intestazioni riga 2: `CodiceInternoIntestazioni` (13 colonne A–M; colonna A = **CODICE INTERNO**). Righe dati: una per ogni voce `CodiceInterno3`, `CodiceInterno4`, … in `cust.config`.

| CODICE INTERNO (A) | Codice offerta fornitore (B) | Descrizione/Progetto ISP (C) | Tecnologia fornitura (D) | ODA ISP (E) | Ref. ISP (F) | Ref. fornitore (G) | Tecnico fornitore (H) | gg/u acquistati (I) | gg/u consumati (J) | gg/u residui (K) | Totale ordine € (L) | Tariffa media € gg/u (M) |
|--------------------|------------------------------|------------------------------|--------------------------|-------------|--------------|--------------------|-----------------------|---------------------|--------------------|--------------------|---------------------|--------------------------|
| Campo 1 (codice interno, colonna A) della voce `CodiceInternoN`. | Campo 2. | Campo 3. | Campo 4. | Campo 5. | Campo 6. | Campo 7. | Campo 8. | Campo 9 (giornate acquistate, inserite in config). | Somma **Actual Hours** ÷ 8 delle righe il cui `Riferimento tabella 1` **oppure** `Sotto Riferimento tabella 1` coincide con il **ref-key** (colonna N, campo 14) della voce `CodiceInternoN`; include Non-Billable Labor; esclude righe con riferimento vuoto o 0. | `I − consumo billable`, dove il consumo billable usa la stessa regola di J ma **esclude** Non-Billable Labor. | Campo 12 (da config; tipicamente importo ordine). | Campo 13 (tariffa media; colonna M resta vuota in output). Ref-key in colonna N. |

**Riga separatore:** se il codice interno (campo 1 / colonna A) è `-`, la riga è solo visiva (campi da config, senza calcoli J/K).

**Seconda copia della tabella:** colonne I, J, K arrotondate all'intero più vicino; le altre colonne testuali restano invariate.

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

## Tabella codice interno (foglio «Tabella di Export»)

Il foglio **Tabella di Export** (e la sezione omonima nel report HTML) riepiloga le giornate consuntivate per **codice interno** (colonna A), incrociando le righe del sorgente con le voci definite in `cust.config` (`CodiceInterno3`, `CodiceInterno4`, …).

**Chiavi config:** `CodiceInternoTitolo`, `CodiceInternoIntestazioni`, `CodiceInterno3` … (legacy: `Export1`, `Export2`, `Export3` … ancora accettate).

**Regola di esclusione:** le righe con **Riferimento tabella 1** vuoto, mancante o pari a **0** restano nel flusso normale (pivot, foglio `progetti`, riepiloghi, Tentative, grafici) ma **non** entrano nelle somme J/K del tab codice interno — né per il riferimento principale né per il sotto-riferimento di quella riga.

Le righe con riferimento valorizzato e diverso da zero contribuiscono alle colonne J/K (giornate consuntivate e residuo) incrociando sia `Riferimento tabella 1` sia `Sotto Riferimento tabella 1` con il ref-key (colonna N) di ciascuna voce `CodiceInternoN` in config.

- **Colonna J:** somma di tutte le ore consuntivate (in giornate), incluse quelle con milestone **Non-Billable Labor**.
- **Colonna K:** giornate acquisite (config) meno solo le ore **billable**; le ore Non-Billable Labor non riducono il residuo.

---

## Automazione con goAll

Lo script `goAll` copia gli export PSA da Google Drive, li elabora con `elabora_progetti.py` e rinomina i sorgenti rimuovendo il prefisso `DvPFilter-`. Vedi [goAll.md](goAll.md) per il flusso completo (PSA + MIRO).

---

## Riferimenti nel repository

| File | Ruolo |
|------|-------|
| `elabora_progetti.py` | Script principale |
| `goAll` | Automazione copia Drive ed elaborazione incrementale |
| `script.config` | Configurazione generale (non versionata) |
| `cust.config` | Configurazione cliente (non versionata) |
| `docs/elabora_progetti.md` | Questa documentazione |
| `docs/goAll.md` | Documentazione automazione `goAll` |
