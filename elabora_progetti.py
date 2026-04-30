"""
elabora_progetti.py — Report settimanale risorse consulenza Red Hat Italy

Legge un file Excel di input (export da PSA/pianificazione) e produce un file
Excel di output multi-foglio con:
  - dati     : dati sorgente arricchiti con colonne derivate
  - progetti : riepilogo contratti con giorni consuntivati vs. riscattati
  - Riepilogo Settimanale : pivot actual/estimated per progetto e settimana
  - Dettaglio Ruoli       : pivot estimated con breakdown per ruolo/milestone
  - Tabella di Export     : riepilogo giornate per codice ordine

Utilizzo:
    python elabora_progetti.py [input.xlsx] [nome.config] [output.xlsx]

    I tre argomenti sono opzionali; i default sono:
        input.xlsx, nome.config, output_elaborato.xlsx
"""

import pandas as pd
import os
import re
import sys
import logging
import subprocess
import traceback
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference

logging.basicConfig(level=logging.INFO, format='%(message)s')
log = logging.getLogger(__name__)

# --- FUNZIONI DI SUPPORTO ---

def carica_config(nome_file):
    """Legge un file .config chiave=valore e restituisce un dizionario.

    Ignora righe vuote e righe che iniziano con '#'.
    Usa UTF-8 per supportare caratteri accentati nei valori.

    Args:
        nome_file: percorso del file di configurazione.

    Returns:
        dict con le coppie chiave/valore; dizionario vuoto se il file non esiste.
    """
    config = {}
    if not os.path.exists(nome_file):
        return config
    with open(nome_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and "=" in line and not line.startswith('#'):
                key, value = line.split('=', 1)
                config[key.strip()] = value.strip()
    return config

def indicizza_contratti(config):
    """Costruisce un indice inverso {nome_contratto -> suffisso} dalla config.

    Le chiavi ContractName1, ContractName2, … vengono mappate ai rispettivi
    suffissi numerici in modo da recuperare velocemente Opportunity, EndDate e
    DaysRedempted per ogni contratto senza scorrere tutta la config ogni volta.

    Args:
        config: dizionario restituito da carica_config.

    Returns:
        dict {contract_name_string: suffix_string}
    """
    contratti = {}
    for key, value in config.items():
        if key.startswith("ContractName"):
            suffix = key.replace("ContractName", "")
            contratti[value.strip()] = suffix
    return contratti

def trova_valore_config(contract_name, contratti_idx, config, prefisso):
    """Recupera un valore di config per un contratto dato il suo nome.

    Esempio: trova_valore_config("ISP_GPSteam25...", idx, cfg, "Opportunity")
    restituisce il valore di Opportunity1 se ContractName1=="ISP_GPSteam25...".

    Args:
        contract_name:  nome del contratto (valore colonna A del sorgente).
        contratti_idx:  indice restituito da indicizza_contratti.
        config:         dizionario della configurazione.
        prefisso:       prefisso della chiave da cercare (es. "Opportunity").

    Returns:
        Stringa con il valore trovato, oppure stringa vuota se non presente.
    """
    suffix = contratti_idx.get(str(contract_name).strip())
    if suffix is None:
        return ""
    return config.get(f"{prefisso}{suffix}", "")

def estrai_settimana(testo):
    """Estrae il numero di settimana da una stringa nel formato *W<nn>*.

    Usata per ricavare la settimana ISO dal campo periodo (es. "CY2025-W14").

    Args:
        testo: stringa o NaN.

    Returns:
        int con il numero di settimana, o None se non trovato/NaN.
    """
    if pd.isna(testo):
        return None
    match = re.search(r'W(\d+)', str(testo), re.IGNORECASE)
    return int(match.group(1)) if match else None

def get_date_range(year, start_week, end_week):
    """Calcola la data di inizio e fine di un intervallo di settimane ISO.

    Usa il 4 gennaio come ancora ISO per trovare il lunedì della prima settimana
    (conforme alla norma ISO 8601).

    Args:
        year:       anno di riferimento.
        start_week: prima settimana ISO dell'intervallo.
        end_week:   ultima settimana ISO dell'intervallo.

    Returns:
        Tuple (data_inizio, data_fine) in formato "dd/mm/yyyy"; ("N/A", "N/A")
        in caso di errore.
    """
    try:
        d = datetime(year, 1, 4)
        start_date = d + timedelta(weeks=(start_week - 1), days=-d.weekday())
        end_date = d + timedelta(weeks=(end_week - 1), days=-d.weekday() + 6)
        return start_date.strftime("%d/%m/%Y"), end_date.strftime("%d/%m/%Y")
    except Exception:
        return "N/A", "N/A"

def calcola_date_settimana(week_str):
    """Restituisce l'intervallo di date (lun-dom) per una settimana ISO.

    Args:
        week_str: stringa nel formato "YYYY-Www" (es. "2025-W14").

    Returns:
        Stringa "dd/mm - dd/mm" con inizio e fine settimana, o None se il
        formato non è riconosciuto.
    """
    match = re.search(r'(\d{4})-W(\d+)', week_str, re.IGNORECASE)
    if not match:
        return None
    year, week = int(match.group(1)), int(match.group(2))
    try:
        lunedi = datetime.fromisocalendar(year, week, 1)
        domenica = datetime.fromisocalendar(year, week, 7)
        return f"{lunedi.strftime('%d/%m')} - {domenica.strftime('%d/%m')}"
    except Exception:
        return None

def splitta_assegnazione(val):
    """Divide il campo assegnazione (colonna C del sorgente) nei suoi 6 componenti.

    Il formato atteso è: "Nome|OPA@profilo|Cliente|SottoProgetto|Riferimento|Commento"
    Il separatore è il carattere pipe '|'. Se i campi sono meno di 6, li riempie
    con stringhe vuote; se sono più di 6, tronca al sesto.

    Args:
        val: stringa grezza della cella, o NaN.

    Returns:
        Lista di esattamente 6 stringhe.
    """
    parti = [p.strip() for p in str(val).split('|')] if pd.notna(val) else []
    while len(parti) < 6:
        parti.append("")
    return parti[:6]

# --- CARICAMENTO E PREPARAZIONE DATI ---

def carica_dati(file_excel_input, config):
    """Carica il file Excel sorgente e prepara il DataFrame arricchito.

    Le colonne di interesse (ruolo, ore stimate, ore effettive, progetto, periodo)
    vengono lette per indice: i valori di default sono compatibili con il formato
    PSA standard ma possono essere sovrascritta tramite config (ColIdxRuolo, ecc.).

    Espande anche la colonna assegnazione (colonna C) nelle 6 colonne derivate:
    "Nome risorsa", "OPA@profilo", "Cliente", "Sotto progetto",
    "Riferimento tabella 1", "Commento".

    Args:
        file_excel_input: percorso del file Excel di input.
        config:           dizionario della configurazione.

    Returns:
        Tuple (df_src, df_dati_comp, col_proj, col_period,
               col_role_name, col_estimated, col_actual)
        dove df_src è il DataFrame originale e df_dati_comp è quello arricchito.
    """
    df_src = pd.read_excel(file_excel_input)

    idx_ruolo = int(config.get('ColIdxRuolo', 4))
    idx_stimato = int(config.get('ColIdxStimato', 7))
    idx_effettivo = int(config.get('ColIdxEffettivo', 8))
    idx_proj = int(config.get('ColIdxProgetto', 0))
    idx_period = int(config.get('ColIdxPeriodo', 9))

    col_role_name = df_src.columns[idx_ruolo]
    col_estimated = df_src.columns[idx_stimato]
    col_actual = df_src.columns[idx_effettivo]

    # Normalizza le ore: virgola decimale → punto, valori non numerici → 0
    for c in [col_actual, col_estimated]:
        df_src[c] = pd.to_numeric(df_src[c].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)

    nuove_col = ["Nome risorsa", "OPA@profilo", "Cliente", "Sotto progetto", "Riferimento tabella 1", "Commento"]
    df_split = pd.DataFrame(df_src.iloc[:, 2].apply(splitta_assegnazione).tolist(), columns=nuove_col)
    df_dati_comp = pd.concat([df_src, df_split], axis=1)

    col_proj = df_dati_comp.columns[idx_proj]
    col_period = df_dati_comp.columns[idx_period]

    return df_src, df_dati_comp, col_proj, col_period, col_role_name, col_estimated, col_actual

# --- CALCOLO PIVOT ---

def calcola_pivot(df_dati_comp, col_period, col_actual, col_estimated,
                  col_proj, col_role_name, weeks_limit_active, start_w, end_w):
    """Calcola le tre pivot table (actual, estimated, estimated per ruolo).

    Se weeks_limit_active è True, filtra le righe in base alla colonna
    'sett_calc' (settimana estratta dal periodo) nell'intervallo [start_w, end_w).

    Le ore vengono convertite in giornate dividendo per 8.

    Args:
        df_dati_comp:        DataFrame arricchito.
        col_period:          nome della colonna periodo (es. "CY2025-W14").
        col_actual:          nome della colonna ore consuntivate.
        col_estimated:       nome della colonna ore stimate.
        col_proj:            nome della colonna progetto.
        col_role_name:       nome della colonna ruolo.
        weeks_limit_active:  True se il filtro settimane è attivo.
        start_w:             prima settimana del filtro (inclusa).
        end_w:               ultima settimana del filtro (esclusa).

    Returns:
        Tuple (df_per_calc, pivot_actual, pivot_estimated, pivot_role_est)
        dove df_per_calc è il DataFrame filtrato usato per i calcoli.
    """
    col_sottoproj = "Sotto progetto"
    col_rif = "Riferimento tabella 1"

    df_per_calc = df_dati_comp.copy()
    if weeks_limit_active:
        mask = (df_dati_comp['sett_calc'].notna()) & \
               (df_dati_comp['sett_calc'] >= start_w) & \
               (df_dati_comp['sett_calc'] < end_w)
        df_per_calc = df_dati_comp[mask].copy()

    def create_pivot(df_input, value_col, index_cols):
        """Pivot con colonne = settimane, valori in giornate, totale riga."""
        temp_df = df_input.copy()
        temp_df[value_col] = temp_df[value_col] / 8.0
        pivot = temp_df.pivot_table(index=index_cols, columns=col_period, values=value_col, aggfunc='sum').fillna(0)
        pivot['TOTALE RIGA'] = pivot.sum(axis=1)
        return pivot.reset_index()

    index_base = [col_proj, col_sottoproj, col_rif]
    pivot_actual = create_pivot(df_per_calc, col_actual, index_base)
    pivot_estimated = create_pivot(df_per_calc, col_estimated, index_base)
    pivot_role_est = create_pivot(df_per_calc, col_estimated, [col_proj, col_sottoproj, col_role_name, col_rif])
    pivot_role_est = pivot_role_est.rename(columns={col_role_name: 'Milestone'})

    return df_per_calc, pivot_actual, pivot_estimated, pivot_role_est

# --- PREPARAZIONE RIGHE PROGETTI ---

def prepara_righe_progetti(df_src, df_dati_comp, df_per_calc, col_proj, col_actual, config, contratti_idx):
    """Costruisce la lista di dizionari da scrivere nel foglio 'progetti'.

    Per ogni progetto unico nel sorgente calcola:
    - giorni PM (ore con profilo @pm o @pc, diviso 8)
    - giorni Consulting (ore rimanenti, diviso 8)
    - giorni riscattati (da DaysRedempted<N> nel config, formato "pm, consulting")
    - opportunity, end date e riferimento da config

    Args:
        df_src:         DataFrame originale.
        df_dati_comp:   DataFrame arricchito (per leggere i metadati di riga).
        df_per_calc:    DataFrame filtrato usato per i calcoli.
        col_proj:       nome della colonna progetto.
        col_actual:     nome della colonna ore consuntivate.
        config:         dizionario della configurazione.
        contratti_idx:  indice {nome_contratto -> suffisso}.

    Returns:
        Lista di dict con chiavi A..K corrispondenti alle colonne del foglio.
    """
    rows_progetti = []
    for proj in df_src[col_proj].unique():
        df_p_full = df_dati_comp[df_dati_comp[col_proj] == proj]
        df_p_calc = df_per_calc[df_per_calc[col_proj] == proj]

        # @pm e @pc sono i profili di Project Manager / Project Coordinator
        pm_mask = df_p_calc['OPA@profilo'].str.contains('@pm|@pc', case=False, na=False)
        giorni_pm = float(df_p_calc.loc[pm_mask, col_actual].sum()) / 8.0
        giorni_cons = float(df_p_calc.loc[~pm_mask, col_actual].sum()) / 8.0

        rif_val_raw = df_p_full.iloc[0]['Riferimento tabella 1']
        rif_val = str(rif_val_raw) if pd.notna(rif_val_raw) else ""

        red_s = trova_valore_config(proj, contratti_idx, config, "DaysRedempted")
        if red_s and "," in red_s:
            try: giorni_pm_red   = float(red_s.split(",")[0].strip())
            except ValueError:   giorni_pm_red   = 0.0
            try: giorni_cons_red = float(red_s.split(",")[1].strip())
            except ValueError:   giorni_cons_red = 0.0
        else:
            giorni_pm_red, giorni_cons_red = 0.0, 0.0

        rows_progetti.append({
            'A': proj,
            'B': df_p_full.iloc[0, 1],
            'C': trova_valore_config(proj, contratti_idx, config, "Opportunity"),
            'D': trova_valore_config(proj, contratti_idx, config, "EndDate"),
            'E': giorni_pm_red,
            'F': giorni_cons_red,
            'G': giorni_pm,
            'H': giorni_cons,
            'I': "",
            'J': "",
            'K': rif_val
        })
    return rows_progetti

# --- SCRITTURA FOGLI BASE ---

def scrivi_fogli_base(file_output, df_dati_comp, rows_progetti):
    """Scrive i cinque fogli con dati grezzi nel file di output.

    I fogli Riepilogo Settimanale, Dettaglio Ruoli e Tabella di Export vengono
    creati vuoti qui; la formattazione e i dati vengono aggiunti nelle funzioni
    successive tramite openpyxl diretto.

    Il foglio 'dati' parte dalla riga 4 (righe 1-3 riservate a metadati).
    Il foglio 'progetti' parte dalla riga 11 (righe 1-10 per intestazioni).

    Args:
        file_output:    percorso del file Excel da creare/sovrascrivere.
        df_dati_comp:   DataFrame arricchito (senza la colonna temporanea sett_calc).
        rows_progetti:  lista di dict prodotta da prepara_righe_progetti.
    """
    with pd.ExcelWriter(file_output, engine='openpyxl') as writer:
        df_dati_comp.drop(columns=['sett_calc']).to_excel(writer, sheet_name='dati', index=False, startrow=0)
        pd.DataFrame(rows_progetti).to_excel(writer, sheet_name='progetti', index=False, startrow=10, header=False)
        pd.DataFrame().to_excel(writer, sheet_name='Riepilogo Settimanale', index=False)
        pd.DataFrame().to_excel(writer, sheet_name='Dettaglio Ruoli', index=False)
        pd.DataFrame().to_excel(writer, sheet_name='Tabella di Export', index=False)

# --- FORMATTAZIONE TAB PROGETTI ---

def formatta_tab_progetti(ws_p, config, rows_progetti, weeks_limit_active, bold, center):
    """Applica intestazioni, formattazione e tabella dati al foglio 'progetti'.

    La struttura del foglio è:
      - Righe 3-4  : intestazioni titolo (merge A:K)
      - Righe 5-7  : informazioni cliente (da config Intestazione8a/9a/10a)
      - Righe 9-10 : header tabella (con merge per gruppi di colonne)
      - Righe 11+  : dati progetti
      - Tabella duplicata a distanza fissa (per layout di stampa)

    Le date di scadenza entro 2 settimane vengono evidenziate in giallo.

    Args:
        ws_p:               worksheet 'progetti' openpyxl.
        config:             dizionario della configurazione.
        rows_progetti:      lista di dict prodotta da prepara_righe_progetti.
        weeks_limit_active: True se il filtro settimane è attivo.
        bold:               Font(bold=True) precostruito.
        center:             Alignment(horizontal='center') precostruito.
    """
    ws_p.merge_cells("A3:K3")
    ws_p['A3'] = config.get('Intestazione5', '')
    ws_p['A3'].font = Font(bold=True, sz=14)
    ws_p['A3'].alignment = center

    ws_p.merge_cells("A4:K4")
    ws_p['A4'] = config.get('Intestazione6', '')
    ws_p['A4'].font = Font(bold=True, sz=12)
    ws_p['A4'].alignment = center

    ws_p['A5'] = config.get('Intestazione8a', '')
    ws_p['A6'] = config.get('Intestazione9a', '')
    ws_p['A7'] = config.get('Intestazione10a', '')

    yellow_fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    oggi = datetime.now().date()
    limite_due_sett = oggi + timedelta(weeks=2)

    def scrivi_intestazione(base_row):
        """Scrive l'intestazione a due righe della tabella contratti."""
        for col_lett, titolo in [('A', 'Contract name'), ('B', 'OPA Number'),
                                  ('C', 'Opportunity'), ('D', 'End Date'), ('K', 'Riferimento')]:
            ws_p[f'{col_lett}{base_row}'] = titolo
            ws_p.merge_cells(f'{col_lett}{base_row}:{col_lett}{base_row + 1}')
        ws_p.merge_cells(f'E{base_row}:F{base_row}')
        ws_p[f'E{base_row}'] = 'Days redempted'
        ws_p.merge_cells(f'G{base_row}:H{base_row}')
        ws_p[f'G{base_row}'] = 'Days Used'
        ws_p.merge_cells(f'I{base_row}:J{base_row}')
        ws_p[f'I{base_row}'] = 'Days remaining'
        for i, testo in enumerate(["Project Manager", "Consulting"] * 3):
            ws_p.cell(row=base_row + 1, column=5 + i).value = testo
        for riga in [base_row, base_row + 1]:
            for cella in ws_p[riga]:
                cella.font, cella.alignment = bold, center

    def scrivi_dati(data_start_row, arrotonda_gh=False):
        """Scrive le righe dati e applica highlight per scadenze imminenti."""
        for i, row in enumerate(rows_progetti):
            r = data_start_row + i
            ws_p.cell(row=r, column=1).value = row['A']
            ws_p.cell(row=r, column=2).value = row['B']
            ws_p.cell(row=r, column=3).value = row['C']
            ws_p.cell(row=r, column=4).value = row['D']
            ws_p.cell(row=r, column=5).value = row['E']
            ws_p.cell(row=r, column=6).value = row['F']
            g_val, h_val = row['G'], row['H']
            if arrotonda_gh:
                try: g_val = round(float(g_val))
                except (ValueError, TypeError): pass
                try: h_val = round(float(h_val))
                except (ValueError, TypeError): pass
            ws_p.cell(row=r, column=7).value = g_val
            ws_p.cell(row=r, column=8).value = h_val
            ws_p.cell(row=r, column=11).value = row['K']
            if row['D']:
                try:
                    end_date = datetime.strptime(str(row['D']), "%d/%m/%Y").date()
                    if end_date <= limite_due_sett:
                        ws_p.cell(row=r, column=4).fill = yellow_fill
                except ValueError:
                    pass
            try:
                ws_p.cell(row=r, column=9).value  = float(row['E']) - float(g_val)
            except (ValueError, TypeError):
                ws_p.cell(row=r, column=9).value  = 0.0
            try:
                ws_p.cell(row=r, column=10).value = float(row['F']) - float(h_val)
            except (ValueError, TypeError):
                ws_p.cell(row=r, column=10).value = 0.0
            ws_p.cell(row=r, column=2).alignment = center
            for col in range(6, 11):
                ws_p.cell(row=r, column=col).alignment = center
            ws_p.cell(row=r, column=11).alignment = center

    scrivi_intestazione(9)
    scrivi_dati(11, arrotonda_gh=True)

    # Tabella duplicata: distanza fissa dalla prima per il layout di stampa
    n = len(rows_progetti)
    dup_header_row = 10 + n + 6
    scrivi_intestazione(dup_header_row)
    scrivi_dati(dup_header_row + 2)

# --- FORMATTAZIONE TAB RIEPILOGO SETTIMANALE ---

def formatta_riepilogo(ws_rs, ws_dr, pivot_actual, pivot_estimated, pivot_role_est,
                       current_week_str, bold, center, green_fill, red_thick,
                       df_dati_comp, col_proj, col_role_name, col_period, col_status_k, col_status_l,
                       df_per_calc, col_actual):
    """Riempie i fogli 'Riepilogo Settimanale' e 'Dettaglio Ruoli'.

    Riepilogo Settimanale: due pivot (actual e estimated) in sequenza verticale.

    Dettaglio Ruoli: pivot estimated per ruolo con:
    - Riga 3: intervallo date di ogni settimana (lun-dom)
    - Riga 4: nome del mese (o "mese1-mese2" se la settimana è a cavallo di mese)
    - Colorazione pastello dalla settimana corrente in poi basata su
      stato schedulazione (colonne K e L del sorgente):
        verde  = Scheduled/Commit
        rosso  = Tentative/Exclude
        giallo = Tentative/Upside
        viola  = altro
    - Bordo blu spesso sulla settimana corrente
    - Colonna extra "Actual Hours (Giornate)" a destra della tabella
    - Legenda colori in fondo alla tabella

    Args:
        ws_rs:            worksheet 'Riepilogo Settimanale'.
        ws_dr:            worksheet 'Dettaglio Ruoli'.
        pivot_actual:     DataFrame pivot ore consuntivate.
        pivot_estimated:  DataFrame pivot ore stimate.
        pivot_role_est:   DataFrame pivot ore stimate per ruolo.
        current_week_str: stringa settimana corrente es. "CY2025-W14".
        bold:             Font(bold=True).
        center:           Alignment(horizontal='center').
        green_fill:       PatternFill verde chiaro per evidenziare la settimana corrente.
        red_thick:        Side(style='thick', color='FF0000') per bordi gruppi.
        df_dati_comp:     DataFrame arricchito (per lookup stato schedulazione).
        col_proj:         nome colonna progetto.
        col_role_name:    nome colonna ruolo.
        col_period:       nome colonna periodo.
        col_status_k:     nome colonna stato schedulazione (colonna K sorgente).
        col_status_l:     nome colonna stato commit/exclude (colonna L sorgente).
        df_per_calc:      DataFrame filtrato per i calcoli.
        col_actual:       nome colonna ore consuntivate.
    """
    def write_pivot(ws, start_row, df, title, border_group=False, extra_center_cols=None):
        """Scrive una pivot su un worksheet a partire da start_row.

        Aggiunge:
        - riga titolo in grassetto
        - riga header colonne in grassetto/centrato
        - righe dati con allineamento numerico centrato
        - evidenziazione verde sulla colonna della settimana corrente
        - bordi rossi spessi per raggruppare righe con stesso progetto
          (solo se border_group=True)
        - riga TOTALE SETTIMANA in fondo

        Returns:
            int: prima riga disponibile dopo la sezione (start_row + altezza + 5).
        """
        if extra_center_cols is None:
            extra_center_cols = set()
        ws.cell(row=start_row, column=1).value = title
        ws.cell(row=start_row, column=1).font = Font(bold=True, sz=12)

        for c_idx, col in enumerate(df.columns):
            cella = ws.cell(row=start_row + 1, column=c_idx + 1)
            cella.value = str(col)
            cella.font = bold
            cella.alignment = center

        num_cols_idx = len(df.columns) - len(df.select_dtypes(include=['number']).columns)
        for r_idx, row in enumerate(df.values):
            for c_idx, val in enumerate(row):
                cella = ws.cell(row=start_row + 2 + r_idx, column=c_idx + 1)
                cella.value = val
                if c_idx == 2 or c_idx >= num_cols_idx or c_idx in extra_center_cols:
                    cella.alignment = center

        ultima_riga = start_row + 2 + len(df) - 1
        ultima_col = len(df.columns)

        for r in range(start_row + 1, ultima_riga + 1):
            ws.cell(row=r, column=ultima_col).font = bold

        for c in range(1, ultima_col):
            if current_week_str in str(ws.cell(row=start_row + 1, column=c).value):
                for r in range(start_row + 1, ultima_riga + 2):
                    ws.cell(row=r, column=c).fill = green_fill

        if border_group:
            grp_start = start_row + 2
            for r in range(grp_start, ultima_riga + 1):
                if r == ultima_riga or ws.cell(row=r, column=1).value != ws.cell(row=r + 1, column=1).value:
                    for riga in range(grp_start, r + 1):
                        ws.cell(row=riga, column=1).border = Border(
                            left=red_thick,
                            top=ws.cell(row=riga, column=1).border.top,
                            bottom=ws.cell(row=riga, column=1).border.bottom
                        )
                        ws.cell(row=riga, column=ultima_col).border = Border(
                            right=red_thick,
                            top=ws.cell(row=riga, column=ultima_col).border.top,
                            bottom=ws.cell(row=riga, column=ultima_col).border.bottom
                        )
                    for col in range(1, ultima_col + 1):
                        ws.cell(row=grp_start, column=col).border = Border(
                            top=red_thick,
                            left=ws.cell(row=grp_start, column=col).border.left,
                            right=ws.cell(row=grp_start, column=col).border.right
                        )
                        ws.cell(row=r, column=col).border = Border(
                            bottom=red_thick,
                            left=ws.cell(row=r, column=col).border.left,
                            right=ws.cell(row=r, column=col).border.right
                        )
                    grp_start = r + 1

        ws.cell(row=ultima_riga + 1, column=1).value = "TOTALE SETTIMANA"
        ws.cell(row=ultima_riga + 1, column=1).font = bold
        for c in range(num_cols_idx + 1, ultima_col + 1):
            s = sum(
                ws.cell(row=r, column=c).value
                for r in range(start_row + 2, ultima_riga + 1)
                if isinstance(ws.cell(row=r, column=c).value, (int, float))
            )
            cella_tot = ws.cell(row=ultima_riga + 1, column=c)
            cella_tot.value = s
            cella_tot.font = bold
            cella_tot.alignment = center

        return ultima_riga + 5

    data_prod = datetime.now().strftime('%d/%m/%Y')
    r_next = write_pivot(ws_rs, 1, pivot_actual, f"ACTUAL HOURS (GIORNATE) - {data_prod}")
    r_next = write_pivot(ws_rs, r_next, pivot_estimated, f"ESTIMATED HOURS (GIORNATE) - {data_prod}")

    pivot_role_display = pivot_role_est.rename(columns={'Riferimento tabella 1': 'Riferimento Interno'})
    dr_next = write_pivot(ws_dr, 1, pivot_role_display,
                          f"DETTAGLIO RUOLI - ESTIMATED HOURS - {data_prod}",
                          border_group=True, extra_center_cols={1, 3})
    ultima_riga_dr = dr_next - 5

    # Inserisce riga 3 (date settimane) e riga 4 (mesi)
    ws_dr.insert_rows(3)
    ws_dr.insert_rows(4)
    num_idx = len(pivot_role_display.columns) - len(pivot_role_display.select_dtypes(include=['number']).columns)
    ultima_col = len(pivot_role_display.columns)

    # Merge colonne indice e TOTALE RIGA su righe 2-4
    for col in range(1, num_idx + 1):
        col_letter = ws_dr.cell(row=2, column=col).column_letter
        ws_dr.merge_cells(f'{col_letter}2:{col_letter}4')
        ws_dr.cell(row=2, column=col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    last_col_letter = ws_dr.cell(row=2, column=ultima_col).column_letter
    ws_dr.merge_cells(f'{last_col_letter}2:{last_col_letter}4')
    ws_dr.cell(row=2, column=ultima_col).alignment = Alignment(horizontal='center', vertical='center')

    mesi_it = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno',
               'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre']

    current_week_col_dr = None
    for col in range(num_idx + 1, ultima_col):
        header_val = ws_dr.cell(row=2, column=col).value
        if not header_val:
            continue
        if current_week_str in str(header_val) and current_week_col_dr is None:
            current_week_col_dr = col
        # Riga 3: date settimana
        date_str = calcola_date_settimana(str(header_val))
        if date_str:
            c3 = ws_dr.cell(row=3, column=col)
            c3.value = date_str
            c3.alignment = center
            c3.font = bold
        # Riga 4: mesi
        match = re.search(r'(\d{4})-W(\d+)', str(header_val), re.IGNORECASE)
        if match:
            yr, wk = int(match.group(1)), int(match.group(2))
            try:
                lun = datetime.fromisocalendar(yr, wk, 1)
                dom = datetime.fromisocalendar(yr, wk, 7)
                ml, md = mesi_it[lun.month - 1], mesi_it[dom.month - 1]
                c4 = ws_dr.cell(row=4, column=col)
                c4.value = ml if ml == md else f"{ml}-{md}"
                c4.alignment = center
                c4.font = bold
            except Exception:
                pass

    fill_verde_past  = PatternFill(fill_type="solid", fgColor="C8E6C9")
    fill_rosso_past  = PatternFill(fill_type="solid", fgColor="FFCDD2")
    fill_giallo_past = PatternFill(fill_type="solid", fgColor="FFF9C4")
    fill_viola_past  = PatternFill(fill_type="solid", fgColor="E1BEE7")
    no_fill          = PatternFill(fill_type=None)

    data_start_row = 5
    ultima_riga_ws = ultima_riga_dr + 2
    col_rif_src    = 'Riferimento tabella 1'

    # Rimuove green_fill dalla colonna settimana corrente e aggiunge bordi blu spessi
    if current_week_col_dr is not None:
        blue_thick = Side(style='thick', color='0070C0')
        first_week_col = num_idx + 1
        for r in range(2, ultima_riga_ws + 2):
            # Bordo sinistro e destro sulla settimana corrente
            cw = ws_dr.cell(row=r, column=current_week_col_dr)
            cw.fill = no_fill
            eb = cw.border
            cw.border = Border(top=eb.top, bottom=eb.bottom, left=blue_thick, right=blue_thick)
            # Bordo destro sulla colonna prima della settimana corrente
            if current_week_col_dr > 1:
                cp = ws_dr.cell(row=r, column=current_week_col_dr - 1)
                ep = cp.border
                cp.border = Border(top=ep.top, bottom=ep.bottom, left=ep.left, right=blue_thick)
            # Bordo sinistro sulla colonna dopo la settimana corrente
            cn = ws_dr.cell(row=r, column=current_week_col_dr + 1)
            en = cn.border
            cn.border = Border(top=en.top, bottom=en.bottom, left=blue_thick, right=en.right)
            # Bordo prima della prima colonna settimana (se diversa dalla corrente)
            if first_week_col != current_week_col_dr:
                cf = ws_dr.cell(row=r, column=first_week_col)
                ef = cf.border
                cf.border = Border(top=ef.top, bottom=ef.bottom, left=blue_thick, right=ef.right)
                ci = ws_dr.cell(row=r, column=first_week_col - 1)
                ei = ci.border
                ci.border = Border(top=ei.top, bottom=ei.bottom, left=ei.left, right=blue_thick)

    # Colorazione pastello dalla settimana corrente in poi (inclusa)
    if current_week_col_dr is not None:
        for c in range(current_week_col_dr, ultima_col):
            week_val = str(ws_dr.cell(row=2, column=c).value or '')
            if not week_val:
                continue
            for r in range(data_start_row, ultima_riga_ws + 1):
                proj      = ws_dr.cell(row=r, column=1).value
                if not proj:
                    continue
                sottoprog = ws_dr.cell(row=r, column=2).value
                milestone = ws_dr.cell(row=r, column=3).value
                rif       = ws_dr.cell(row=r, column=4).value

                mask = (
                    (df_dati_comp[col_proj].astype(str)        == str(proj))      &
                    (df_dati_comp['Sotto progetto'].astype(str) == str(sottoprog)) &
                    (df_dati_comp[col_role_name].astype(str)   == str(milestone)) &
                    (df_dati_comp[col_rif_src].astype(str)     == str(rif))       &
                    (df_dati_comp[col_period].astype(str)      == week_val)
                )
                subset = df_dati_comp[mask]
                if subset.empty:
                    continue

                k_vals = set(subset[col_status_k].astype(str).str.strip().str.lower().unique())
                l_vals = set(subset[col_status_l].astype(str).str.strip().str.lower().unique())

                cella = ws_dr.cell(row=r, column=c)
                if k_vals == {'scheduled'} and l_vals == {'commit'}:
                    cella.fill = fill_verde_past
                elif k_vals == {'tentative'} and l_vals == {'exclude'}:
                    cella.fill = fill_rosso_past
                elif k_vals == {'tentative'} and l_vals == {'upside'}:
                    cella.fill = fill_giallo_past
                else:
                    cella.fill = fill_viola_past

    # Colonna Actual Hours (Giornate) — ultima_col + 2 (lascia una colonna vuota)
    act_col = ultima_col + 2
    actual_grp = (
        df_per_calc.groupby([col_proj, 'Sotto progetto', col_role_name, col_rif_src])[col_actual]
        .sum() / 8.0
    ).to_dict()

    act_hdr = ws_dr.cell(row=2, column=act_col)
    act_hdr.value = "Actual Hours (Giornate)"
    act_hdr.font = bold
    act_hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_dr.merge_cells(
        start_row=2, start_column=act_col,
        end_row=4,   end_column=act_col
    )

    totale_actual = 0.0
    for r in range(data_start_row, ultima_riga_ws + 1):
        proj      = ws_dr.cell(row=r, column=1).value
        if not proj:
            continue
        sottoprog = ws_dr.cell(row=r, column=2).value
        milestone = ws_dr.cell(row=r, column=3).value
        rif       = ws_dr.cell(row=r, column=4).value
        val = actual_grp.get((proj, sottoprog, milestone, rif), 0.0)
        c_act = ws_dr.cell(row=r, column=act_col)
        c_act.value = val
        c_act.alignment = center
        totale_actual += val

    tot_act = ws_dr.cell(row=ultima_riga_ws + 1, column=act_col)
    tot_act.value = totale_actual
    tot_act.font = bold
    tot_act.alignment = center

    # Legenda colori sotto la tabella a partire dalla colonna settimana corrente
    if current_week_col_dr is not None:
        green_med  = Side(style='medium', color='00B050')
        green_thin = Side(style='thin',   color='00B050')
        leggenda = [
            ("Scheduled", "Commit",  fill_verde_past),
            ("Tentative",  "Exclude", fill_rosso_past),
            ("Tentative",  "Upside",  fill_giallo_past),
            ("Altro",      "",        fill_viola_past),
        ]
        leg_start = ultima_riga_ws + 3
        lc1, lc2  = current_week_col_dr, current_week_col_dr + 1
        n_leg     = len(leggenda)

        for i, (t1, t2, fill) in enumerate(leggenda):
            r        = leg_start + i
            is_top   = (i == 0)
            is_bot   = (i == n_leg - 1)
            for off, txt in [(0, t1), (1, t2)]:
                c   = lc1 + off
                cel = ws_dr.cell(row=r, column=c)
                cel.value     = txt
                cel.fill      = fill
                cel.alignment = center
                cel.font      = bold
                cel.border    = Border(
                    top    = green_med  if is_top else green_thin,
                    bottom = green_med  if is_bot else green_thin,
                    left   = green_med  if off == 0 else green_thin,
                    right  = green_med  if off == 1 else green_thin,
                )
        # Merge cella "Altro" sulle due colonne
        ws_dr.merge_cells(
            start_row=leg_start + n_leg - 1, start_column=lc1,
            end_row=leg_start + n_leg - 1,   end_column=lc2
        )

# --- FORMATTAZIONE TAB EXPORT ---

def formatta_tab_export(ws_e, config, df_per_calc, col_rif, col_actual,
                        fill_verde, fill_nero, font_bianco_bold, center, right_align):
    """Riempie il foglio 'Tabella di Export' con il riepilogo giornate per codice ordine.

    Legge le righe di esportazione dalla config (Export3, Export4, …): ogni riga
    è una lista di valori separati da virgola che rappresentano i campi della
    tabella (codice interno, codice offerta, descrizione, …).

    Per ogni riga cerca se uno dei valori del campo "codice" è presente nella
    colonna "Riferimento tabella 1" del sorgente e, se sì, riporta la somma
    delle ore consuntivate (in giornate) nella colonna J.

    Come per il foglio progetti, la tabella è duplicata: la prima copia ha
    valori esatti, la seconda arrotonda le giornate all'intero più vicino.

    Args:
        ws_e:             worksheet 'Tabella di Export'.
        config:           dizionario della configurazione.
        df_per_calc:      DataFrame filtrato per i calcoli.
        col_rif:          nome della colonna riferimento (es. "Riferimento tabella 1").
        col_actual:       nome della colonna ore consuntivate.
        fill_verde:       PatternFill verde scuro per la riga titolo.
        fill_nero:        PatternFill nero per le righe header.
        font_bianco_bold: Font bianco grassetto per testo su sfondo scuro.
        center:           Alignment(horizontal='center').
        right_align:      Alignment(horizontal='right').

    Returns:
        int: numero di riga dell'ultima riga scritta nella tabella duplicata.
    """
    somme_rif = (df_per_calc.groupby(col_rif)[col_actual].sum() / 8.0).to_dict()
    somme_rif = {str(k).strip(): v for k, v in somme_rif.items()}

    righe_export = []
    idx = 3
    while f"Export{idx}" in config:
        vals = [v.strip() for v in config[f"Export{idx}"].split(',')]
        valore_match = next((somme_rif[v] for v in vals if v in somme_rif), 0.0)
        righe_export.append((vals, valore_match))
        idx += 1

    def scrivi_titolo(start_row):
        ws_e.merge_cells(f"B{start_row}:M{start_row}")
        ws_e[f'B{start_row}'] = f"{config.get('Export1', '')} {datetime.now().strftime('%d/%m/%Y')}"
        ws_e[f'B{start_row}'].fill = fill_verde
        ws_e[f'B{start_row}'].font = font_bianco_bold
        ws_e[f'B{start_row}'].alignment = center

    def scrivi_header(start_row):
        header_vals = config.get('Export2', '').split(',')
        for i in range(1, 14):
            cella = ws_e.cell(row=start_row, column=i)
            if i <= len(header_vals):
                cella.value = header_vals[i - 1].strip()
            cella.fill = fill_nero
            cella.font = font_bianco_bold
            cella.alignment = center

    def scrivi_dati(start_row, arrotonda=False):
        for i, (vals, valore_match) in enumerate(righe_export):
            r = start_row + i
            for j, v in enumerate(vals):
                ws_e.cell(row=r, column=1 + j).value = v
            if arrotonda:
                try:
                    i_raw = ws_e.cell(row=r, column=9).value
                    ws_e.cell(row=r, column=9).value = round(float(str(i_raw).replace(',', '.')))
                except (ValueError, TypeError):
                    pass
            ws_e.cell(row=r, column=9).alignment = right_align
            j_val = round(valore_match) if arrotonda else valore_match
            ws_e.cell(row=r, column=10).value = j_val
            try:
                i_num = float(str(ws_e.cell(row=r, column=9).value or 0).replace(',', '.'))
                k_val = round(i_num - j_val) if arrotonda else round(i_num - j_val, 2)
                ws_e.cell(row=r, column=11).value = k_val
            except (ValueError, TypeError):
                ws_e.cell(row=r, column=11).value = ''
        return start_row + len(righe_export)

    scrivi_titolo(1)
    scrivi_header(2)
    riga_fine = scrivi_dati(3)

    dup_start = riga_fine + 5
    scrivi_titolo(dup_start)
    scrivi_header(dup_start + 1)
    scrivi_dati(dup_start + 2, arrotonda=True)

    return dup_start + 1 + len(righe_export)

# --- AGGIUNGI NOTE ---

def aggiungi_note(ws_p, ws_e, anno_corrente, start_w, end_w, rows_progetti, riga_export, bold):
    """Aggiunge una nota sul filtro settimane attivo in fondo ai fogli 'progetti' ed 'Export'.

    La nota riporta l'intervallo di settimane e le date corrispondenti.

    Args:
        ws_p:           worksheet 'progetti'.
        ws_e:           worksheet 'Tabella di Export'.
        anno_corrente:  anno usato per calcolare le date.
        start_w:        prima settimana del filtro.
        end_w:          ultima settimana del filtro.
        rows_progetti:  lista righe progetto (per calcolare la posizione).
        riga_export:    ultima riga scritta nel foglio Export.
        bold:           Font(bold=True).
    """
    d_s, d_e = get_date_range(anno_corrente, start_w, end_w)
    nota1 = "NOTA: Limiti settimane ATTIVATI:"
    nota2 = f"    Range: {start_w} - {end_w} [{d_s} - {d_e}]"
    n = len(rows_progetti)
    lr_p = 17 + 2 * n  # dopo la tabella duplicata (10+n+6+2+n-1)
    for ws_t, lr_t in [(ws_p, lr_p), (ws_e, riga_export)]:
        ws_t[f'A{lr_t + 2}'] = nota1
        ws_t[f'A{lr_t + 2}'].font = Font(italic=True)
        ws_t[f'A{lr_t + 3}'] = nota2
        ws_t[f'A{lr_t + 3}'].font = bold

# --- GENERAZIONE HTML ---

def genera_html(df_dati_comp, rows_progetti, pivot_actual, pivot_estimated,
                pivot_role_est, df_per_calc, config, col_rif, col_actual, file_output):
    """Genera un file HTML navigabile con tutte le tabelle del report.

    Il file ha lo stesso nome del file Excel con estensione .html e viene
    scritto nella stessa directory. Include una barra laterale fissa con
    indice e link ancorati a ciascuna sezione.
    """
    html_path = os.path.splitext(file_output)[0] + '.html'
    now_str   = datetime.now().strftime('%d/%m/%Y %H:%M')
    titolo    = config.get('Intestazione5', 'Report Settimanale Risorse')

    df_dati = df_dati_comp.drop(columns=['sett_calc'], errors='ignore')

    # ── DataFrame progetti ───────────────────────────────────────────────────
    def _f(v):
        try:    return round(float(v), 2)
        except: return v or ''

    df_proj = pd.DataFrame([{
        'Contract Name': r['A'],
        'OPA Number':    r['B'],
        'Opportunity':   r['C'],
        'End Date':      r['D'],
        'Risc. PM':      _f(r['E']),
        'Risc. Cons.':   _f(r['F']),
        'Usati PM':      _f(r['G']),
        'Usati Cons.':   _f(r['H']),
        'Riferimento':   r['K'],
    } for r in rows_progetti])

    # ── DataFrame export ─────────────────────────────────────────────────────
    somme_rif = {str(k).strip(): round(v / 8.0, 2)
                 for k, v in df_per_calc.groupby(col_rif)[col_actual].sum().items()}
    hdr_exp   = [h.strip() for h in config.get('Export2', '').split(',')]
    righe_exp = []
    i_e = 3
    while f"Export{i_e}" in config:
        vals = [v.strip() for v in config[f"Export{i_e}"].split(',')]
        gg   = next((somme_rif[v] for v in vals if v in somme_rif), 0.0)
        while len(vals) < 10:
            vals.append('')
        righe_exp.append(vals[:9] + [round(gg, 2)] + vals[10:])
        i_e += 1

    if righe_exp:
        nc      = max(len(r) for r in righe_exp)
        df_exp  = pd.DataFrame(
            [r + [''] * (nc - len(r)) for r in righe_exp],
            columns=(hdr_exp + [''] * nc)[:nc]
        )
    else:
        df_exp = pd.DataFrame()

    pv_ruoli = pivot_role_est.rename(columns={'Riferimento tabella 1': 'Riferimento'})

    # ── CSS ─────────────────────────────────────────────────────────────────
    css = """
:root{--pri:#1a56db;--pri-l:#dbeafe;--pri-d:#1e40af;--sid:#0f172a;
  --bg:#f1f5f9;--srf:#fff;--brd:#e2e8f0;--txt:#1e293b;--mut:#64748b}
*,::before,::after{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  background:var(--bg);color:var(--txt);display:flex;min-height:100vh;font-size:14px}
.sidebar{width:215px;min-width:215px;background:var(--sid);padding:1.25rem 1rem;
  position:sticky;top:0;height:100vh;overflow-y:auto;
  display:flex;flex-direction:column;gap:1.25rem}
.sb-title{color:#fff;font-size:.9rem;font-weight:700;padding-bottom:.875rem;
  border-bottom:1px solid rgba(255,255,255,.1)}
.sb-meta{color:#475569;font-size:.73rem;line-height:1.65}
.nav-list{list-style:none;display:flex;flex-direction:column;gap:1px}
.nav-sep{color:#374151;font-size:.65rem;font-weight:700;letter-spacing:.09em;
  text-transform:uppercase;padding:.75rem .5rem .2rem;margin-top:.25rem}
.nav-list a{color:#94a3b8;text-decoration:none;display:block;padding:.37rem .65rem;
  border-radius:4px;font-size:.83rem;transition:background .15s,color .15s}
.nav-list a:hover{background:rgba(255,255,255,.08);color:#e2e8f0}
.nav-list a.sub{padding-left:1.15rem;font-size:.79rem;color:#64748b}
.nav-list a.sub:hover{color:#94a3b8}
main{flex:1;padding:1.75rem 2rem;overflow-x:hidden;min-width:0}
.pg-hdr{margin-bottom:1.75rem}
.pg-hdr h1{font-size:1.45rem;font-weight:700;color:var(--pri-d);margin-bottom:.3rem}
.pg-hdr p{color:var(--mut);font-size:.82rem}
.sec{background:var(--srf);border-radius:8px;
  box-shadow:0 1px 4px rgba(0,0,0,.09);margin-bottom:1.75rem;overflow:hidden}
.sec-hdr{background:var(--pri);color:#fff;padding:.7rem 1.2rem;
  display:flex;align-items:center;justify-content:space-between}
.sec-hdr h2{font-size:.95rem;font-weight:600;letter-spacing:.01em}
.badge{background:rgba(255,255,255,.22);border-radius:20px;
  padding:.1rem .55rem;font-size:.72rem}
.sub-sec{padding:.875rem 1.2rem;border-bottom:1px solid var(--brd)}
.sub-sec:last-child{border-bottom:none}
.sub-sec h3{font-size:.85rem;font-weight:600;color:var(--pri-d);margin-bottom:.65rem;
  display:flex;align-items:center;gap:.45rem}
.sub-sec h3::before{content:'';display:inline-block;width:3px;height:.9em;
  background:var(--pri);border-radius:2px}
.tbl-wrap{overflow:auto;max-height:520px;border-top:1px solid var(--brd)}
table{border-collapse:collapse;width:100%;font-size:.8rem}
thead th{background:var(--pri-l);color:var(--pri-d);font-weight:600;padding:.5rem .72rem;
  text-align:left;white-space:nowrap;position:sticky;top:0;
  border-bottom:2px solid var(--pri);z-index:1}
tbody tr:nth-child(even){background:#f8fafc}
tbody tr:hover{background:#eff6ff}
td{padding:.38rem .72rem;border-bottom:1px solid var(--brd);white-space:nowrap}
td.num{text-align:right;font-variant-numeric:tabular-nums;color:#1e3a8a}
.row-count{padding:.35rem 1.2rem;font-size:.72rem;color:var(--mut);
  border-top:1px solid var(--brd);background:#f8fafc}
.empty{padding:1.25rem;color:var(--mut);font-style:italic}
.btt{position:fixed;bottom:1.25rem;right:1.25rem;background:var(--pri);color:#fff;
  border:none;border-radius:50%;width:2.25rem;height:2.25rem;font-size:1.1rem;
  cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.25);transition:background .15s;
  display:flex;align-items:center;justify-content:center}
.btt:hover{background:var(--pri-d)}
"""

    # ── Helper: DataFrame → tabella HTML ────────────────────────────────────
    def _cell(val):
        if pd.isna(val):
            return '<td></td>'
        if isinstance(val, float):
            return f'<td class="num">{val:,.2f}</td>'
        return f'<td>{val}</td>'

    def _tbl(df):
        if df is None or df.empty:
            return '<p class="empty">Nessun dato disponibile.</p>'
        out = ['<div class="tbl-wrap"><table><thead><tr>']
        for col in df.columns:
            out.append(f'<th>{col}</th>')
        out.append('</tr></thead><tbody>')
        for _, row in df.iterrows():
            out.append('<tr>')
            out.extend(_cell(v) for v in row)
            out.append('</tr>')
        out.append('</tbody></table></div>')
        out.append(f'<p class="row-count">{len(df):,} record</p>')
        return ''.join(out)

    # ── Sezioni ─────────────────────────────────────────────────────────────
    sec_dati = f"""
<section class="sec" id="dati">
  <div class="sec-hdr"><h2>Dati</h2><span class="badge">{len(df_dati):,} righe</span></div>
  <div class="sub-sec">{_tbl(df_dati)}</div>
</section>"""

    sec_proj = f"""
<section class="sec" id="progetti">
  <div class="sec-hdr"><h2>Progetti</h2><span class="badge">{len(df_proj)} contratti</span></div>
  <div class="sub-sec">{_tbl(df_proj)}</div>
</section>"""

    sec_riep = f"""
<section class="sec" id="riepilogo">
  <div class="sec-hdr"><h2>Riepilogo Settimanale</h2></div>
  <div class="sub-sec" id="actual">
    <h3>Actual Hours (Giornate)</h3>
    {_tbl(pivot_actual)}
  </div>
  <div class="sub-sec" id="estimated">
    <h3>Estimated Hours (Giornate)</h3>
    {_tbl(pivot_estimated)}
  </div>
</section>"""

    sec_ruoli = f"""
<section class="sec" id="dettaglio-ruoli">
  <div class="sec-hdr"><h2>Dettaglio Ruoli</h2></div>
  <div class="sub-sec">{_tbl(pv_ruoli)}</div>
</section>"""

    exp_title = config.get('Export1', 'Tabella di Export')
    sec_exp = f"""
<section class="sec" id="export">
  <div class="sec-hdr"><h2>Tabella di Export</h2></div>
  <div class="sub-sec">
    <h3>{exp_title}</h3>
    {_tbl(df_exp)}
  </div>
</section>"""

    # ── Navigazione laterale ─────────────────────────────────────────────────
    nav = f"""<nav class="sidebar">
  <div class="sb-title">Report Risorse</div>
  <div class="sb-meta">Generato il<br>{now_str}</div>
  <ul class="nav-list">
    <li class="nav-sep">Indice</li>
    <li><a href="#dati">Dati</a></li>
    <li><a href="#progetti">Progetti</a></li>
    <li><a href="#riepilogo">Riepilogo Settimanale</a></li>
    <li><a href="#actual" class="sub">&#x2937; Actual</a></li>
    <li><a href="#estimated" class="sub">&#x2937; Estimated</a></li>
    <li><a href="#dettaglio-ruoli">Dettaglio Ruoli</a></li>
    <li><a href="#export">Tabella Export</a></li>
  </ul>
</nav>"""

    # ── Assemblaggio finale ──────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="it">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{titolo}</title>
  <style>{css}</style>
</head>
<body>
{nav}
<main>
  <div class="pg-hdr">
    <h1>{titolo}</h1>
    <p>Generato il {now_str}</p>
  </div>
{sec_dati}
{sec_proj}
{sec_riep}
{sec_ruoli}
{sec_exp}
</main>
<button class="btt" onclick="window.scrollTo({{top:0,behavior:'smooth'}})" title="Torna in cima">&#8679;</button>
</body>
</html>"""

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    log.info("HTML generato: %s", html_path)


# --- TAB GRAFICI ---

def crea_tab_grafici(wb, df_per_calc, rows_progetti, col_proj, col_period, col_actual, col_estimated):
    """Crea il foglio 'grafici' con 5 grafici su dati aggregati.

    Grafici prodotti:
      1. Actual vs Estimated per progetto (barre raggruppate)
      2. Trend settimanale giornate actual (linea con marcatori)
      3. Ripartizione per profilo PM/Consulting (ciambella)
      4. Top risorse per giornate actual (barre orizzontali)
      5. Stato contratti: giorni riscattati vs usati (barre sovrapposte)

    Le tabelle dati di supporto vengono scritte in colonna A del foglio;
    i grafici sono disposti in una griglia 2×2 + 1 a partire da colonna H.
    """
    ws = wb.create_sheet('grafici')
    hdr_font = Font(bold=True)

    # ── aggregazioni ──────────────────────────────────────────────────────────

    actual_pp     = df_per_calc.groupby(col_proj)[col_actual].sum() / 8.0
    estimated_pp  = df_per_calc.groupby(col_proj)[col_estimated].sum() / 8.0
    progetti      = list(actual_pp.index)

    def _sort_key(s):
        m = re.search(r'(\d{4}).*W(\d+)', str(s), re.IGNORECASE)
        return (int(m.group(1)), int(m.group(2))) if m else (0, 0)

    trend = df_per_calc.groupby(col_period)[col_actual].sum() / 8.0
    trend = trend.loc[sorted(trend.index, key=_sort_key)]

    risorse_s = (
        df_per_calc[df_per_calc['Nome risorsa'].astype(str).str.strip().ne('') &
                    df_per_calc['Nome risorsa'].notna()]
        .groupby('Nome risorsa')[col_actual].sum() / 8.0
    ).sort_values(ascending=False).head(10)

    # ── scrittura tabelle dati (col A–E) ──────────────────────────────────────

    def hdr(row, *labels):
        for i, lbl in enumerate(labels):
            c = ws.cell(row=row, column=1 + i)
            c.value = lbl
            c.font  = hdr_font

    # Tabella 1: actual vs estimated per progetto
    T1 = 2
    hdr(T1, 'Progetto', 'Actual (gg)', 'Estimated (gg)')
    for i, proj in enumerate(progetti):
        ws.cell(row=T1+1+i, column=1).value = proj
        ws.cell(row=T1+1+i, column=2).value = round(float(actual_pp[proj]),    2)
        ws.cell(row=T1+1+i, column=3).value = round(float(estimated_pp[proj]), 2)
    T1_END = T1 + len(progetti)

    # Tabella 2: trend settimanale
    T2 = T1_END + 2
    hdr(T2, 'Settimana', 'Giornate Actual')
    for i, (week, val) in enumerate(trend.items()):
        ws.cell(row=T2+1+i, column=1).value = str(week)
        ws.cell(row=T2+1+i, column=2).value = round(float(val), 2)
    T2_END = T2 + len(trend)

    # Tabella 4: top risorse
    T4 = T2_END + 2
    hdr(T4, 'Risorsa', 'Giornate Actual')
    for i, (risorsa, val) in enumerate(risorse_s.items()):
        ws.cell(row=T4+1+i, column=1).value = risorsa
        ws.cell(row=T4+1+i, column=2).value = round(float(val), 2)
    T4_END = T4 + len(risorse_s)

    # Tabella 5: stato contratti
    T5 = T4_END + 2
    hdr(T5, 'Contratto', 'Risc. PM', 'Risc. Cons.', 'Usati PM', 'Usati Cons.')
    for i, r in enumerate(rows_progetti):
        row = T5 + 1 + i
        ws.cell(row=row, column=1).value = r['A']
        for j, key in enumerate(['E', 'F', 'G', 'H'], start=2):
            try:
                ws.cell(row=row, column=j).value = round(float(r[key]), 2)
            except (ValueError, TypeError):
                ws.cell(row=row, column=j).value = 0.0
    T5_END = T5 + len(rows_progetti)

    # ── helper grafici ────────────────────────────────────────────────────────

    CW, CH = 18, 12  # larghezza e altezza grafici in cm

    def _bar(title, horiz=False, stacked=False):
        c = BarChart()
        c.type   = "bar" if horiz else "col"
        c.title  = title
        c.width  = CW
        c.height = CH
        if stacked:
            c.grouping = "stacked"
        return c

    # ── grafico 1: actual vs estimated (col H, riga 2) ───────────────────────
    c1   = _bar("Actual vs Estimated per Progetto (giornate)")
    ref1 = Reference(ws, min_col=2, min_row=T1,     max_col=3, max_row=T1_END)
    cat1 = Reference(ws, min_col=1, min_row=T1+1,              max_row=T1_END)
    c1.add_data(ref1, titles_from_data=True)
    c1.set_categories(cat1)
    ws.add_chart(c1, "H2")

    # ── grafico 4: top risorse orizzontale (col T, riga 2) ───────────────────
    c4   = _bar("Top Risorse per Giornate Actual", horiz=True)
    ref4 = Reference(ws, min_col=2, min_row=T4,     max_row=T4_END)
    cat4 = Reference(ws, min_col=1, min_row=T4+1,   max_row=T4_END)
    c4.add_data(ref4, titles_from_data=True)
    c4.set_categories(cat4)
    ws.add_chart(c4, "T2")

    # ── grafico 2: trend settimanale (col H, riga 30) ────────────────────────
    c2 = LineChart()
    c2.title  = "Trend Settimanale Giornate Actual"
    c2.width  = CW
    c2.height = CH
    ref2 = Reference(ws, min_col=2, min_row=T2,   max_row=T2_END)
    cat2 = Reference(ws, min_col=1, min_row=T2+1, max_row=T2_END)
    c2.add_data(ref2, titles_from_data=True)
    c2.set_categories(cat2)
    ws.add_chart(c2, "H30")

    # ── grafico 5: stato contratti impilato (col T, riga 30) ─────────────────
    c5   = _bar("Stato Contratti: Giorni Riscattati vs Utilizzati", stacked=True)
    ref5 = Reference(ws, min_col=2, min_row=T5,   max_col=5, max_row=T5_END)
    cat5 = Reference(ws, min_col=1, min_row=T5+1, max_row=T5_END)
    c5.add_data(ref5, titles_from_data=True)
    c5.set_categories(cat5)
    ws.add_chart(c5, "T30")



# --- FUNZIONE PRINCIPALE ---

def elabora_dati(file_excel_input, file_config, file_output):
    """Orchestratore principale: legge input, calcola, scrive e formatta l'output.

    Flusso:
    1. Carica config e indice contratti
    2. Legge il file Excel e prepara i DataFrame
    3. Calcola le pivot table
    4. Prepara le righe del foglio progetti
    5. Scrive i fogli base con pandas ExcelWriter
    6. Riapre il file con openpyxl e applica formattazione
    7. Salva il file finale

    Args:
        file_excel_input: percorso del file Excel sorgente.
        file_config:      percorso del file di configurazione .config.
        file_output:      percorso del file Excel di output da generare.
    """
    try:
        config = carica_config(file_config)
        contratti_idx = indicizza_contratti(config)

        start_w = int(config.get('StartWeek', 0))
        end_w = int(config.get('EndWeek', 99))
        weeks_limit_active = config.get('WeeksLimit', 'no').lower() == 'yes'

        oggi = datetime.now()
        anno_corrente = oggi.year
        settimana_corrente = oggi.isocalendar()[1]
        current_week_str = f"CY{anno_corrente}-W{settimana_corrente:02d}"

        if not os.path.exists(file_excel_input):
            log.error("ERRORE: %s non trovato.", file_excel_input)
            return

        log.info("1. Caricamento e preparazione dati...")
        df_src, df_dati_comp, col_proj, col_period, col_role_name, col_estimated, col_actual = carica_dati(file_excel_input, config)

        df_dati_comp['sett_calc'] = df_dati_comp[col_period].apply(estrai_settimana)

        df_per_calc, pivot_actual, pivot_estimated, pivot_role_est = calcola_pivot(
            df_dati_comp, col_period, col_actual, col_estimated,
            col_proj, col_role_name, weeks_limit_active, start_w, end_w
        )

        col_rif = "Riferimento tabella 1"
        rows_progetti = prepara_righe_progetti(df_src, df_dati_comp, df_per_calc, col_proj, col_actual, config, contratti_idx)

        log.info("2. Scrittura fogli base...")
        scrivi_fogli_base(file_output, df_dati_comp, rows_progetti)

        log.info("3. Applicazione formattazione e dati mancanti...")
        wb = load_workbook(file_output)

        ws_d = wb['dati']

        try:
            git_date = subprocess.check_output(
                ['git', 'log', '--format=%cd %h', '--date=format:%Y-%m-%d', '-1'],
                cwd=os.path.dirname(os.path.abspath(__file__)),
                stderr=subprocess.DEVNULL
            ).decode().strip()
        except Exception:
            git_date = ""
        # startrow=0 → header a riga 1, dati da riga 2 fino a riga 1+n
        ultima_riga_dati = 1 + len(df_dati_comp)
        versione_row = ultima_riga_dati + 2
        ws_d.cell(row=versione_row, column=1).value = f"Versione script: {git_date}"

        bold = Font(bold=True)
        center = Alignment(horizontal='center', vertical='center')
        right_align = Alignment(horizontal='right', vertical='center')
        red_thick = Side(style='thick', color='FF0000')
        green_fill = PatternFill(fill_type="solid", fgColor="C6EFCE")
        fill_verde = PatternFill(fill_type="solid", fgColor="006400")
        fill_nero = PatternFill(fill_type="solid", fgColor="000000")
        font_bianco_bold = Font(color="FFFFFF", bold=True)

        formatta_tab_progetti(wb['progetti'], config, rows_progetti, weeks_limit_active, bold, center)
        # Colonne K e L del sorgente contengono lo stato di schedulazione e commit/exclude
        col_status_k = df_dati_comp.columns[10]
        col_status_l = df_dati_comp.columns[11]
        formatta_riepilogo(wb['Riepilogo Settimanale'], wb['Dettaglio Ruoli'],
                           pivot_actual, pivot_estimated, pivot_role_est,
                           current_week_str, bold, center, green_fill, red_thick,
                           df_dati_comp, col_proj, col_role_name, col_period, col_status_k, col_status_l,
                           df_per_calc, col_actual)

        riga_export = formatta_tab_export(
            wb['Tabella di Export'], config, df_per_calc, col_rif, col_actual,
            fill_verde, fill_nero, font_bianco_bold, center, right_align
        )

        if weeks_limit_active:
            aggiungi_note(wb['progetti'], wb['Tabella di Export'], anno_corrente,
                          start_w, end_w, rows_progetti, riga_export, bold)

        log.info("4. Creazione tab grafici...")
        crea_tab_grafici(wb, df_per_calc, rows_progetti, col_proj, col_period, col_actual, col_estimated)

        log.info("5. Generazione file HTML...")
        genera_html(df_dati_comp, rows_progetti, pivot_actual, pivot_estimated,
                    pivot_role_est, df_per_calc, config, col_rif, col_actual, file_output)

        wb.save(file_output)
        log.info("SUCCESSO: File generato con tutte le intestazioni e dati Export.")

    except Exception as e:
        log.error("ERRORE: %s", str(e))
        traceback.print_exc()


if __name__ == "__main__":
    input_file = sys.argv[1] if len(sys.argv) > 1 else 'input.xlsx'
    config_file = sys.argv[2] if len(sys.argv) > 2 else 'nome.config'
    output_file = sys.argv[3] if len(sys.argv) > 3 else 'output_elaborato.xlsx'
    elabora_dati(input_file, config_file, output_file)
