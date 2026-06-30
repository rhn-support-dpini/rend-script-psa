"""
elabora_JKAN.py — Estrae le card dalle Kanban in un export MIRO

Legge un export CSV di una board MIRO, individua tutte le sezioni Kanban
e produce un file Excel con una riga per card (espansa per tag temporali in Description).

Utilizzo:
    python elabora_JKAN.py [input.csv]
    python elabora_JKAN.py -h

Parametri:
    input.csv   (opzionale) Export CSV di una board MIRO da elaborare.
                Default: 2026-06-06-WIP.csv nella cartella dello script.
                Accetta percorso relativo o assoluto.

Output:
    File .xlsx con stesso nome e percorso del CSV di input.
    Fogli: data-all (tutte le card), data-check (card attive da verificare), stat, graph.
    dbJKAN.csv — storico snapshot colonne Kanban (cartella dello script).
    dbJKAN.html — grafici burnup, WIP, velocità (stessa cartella del CSV input).
    Le righe Description con prefisso "#" generano sotto-righe da colonna N;
    A–M sono merge verticali per Title, con bordo rosso pastello per card.
    Colonna J (InizioLavorazione(GG)), K (Waiting #), L (Totale Lavorazione),
    M (Period SUM), N (Giorni), O (TAG Temporali).
    Waiting #: giorni dall'ultimo tag "# Waiting -" a oggi, con sfondo giallo pastello.
    Totale Lavorazione (L): percentuale su Estimate (0-50 verde, 51-80 giallo, 81-100 rosso pastello, >100 rosso acceso).
    Period SUM: somma Giorni con tag in Progress.
    Colonna G (Estimate): valore numerico dal tag "# Estimate" in Description, non dal CSV.
    Giorni: se manca la 2ª data si usa oggi, eccetto tag Done (solo chiusura).
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import Counter
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, PatternFill, Side

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_JKAN_CSV = os.path.join(SCRIPT_DIR, "dbJKAN.csv")
DB_JKAN_HTML_NAME = "dbJKAN.html"
DATA_IN_FILENAME_RE = re.compile(r"(\d{4})[-/](\d{2})[-/](\d{2})")
SCOPE_TOTALE_CARD = 30
KANBAN_SNAPSHOT_COLS = [
    "Under analysis",
    "backlog",
    "in progress",
    "waiting for fab.",
    "Fab test in progress",
    "Acronimi done",
]
DB_JKAN_HEADERS = ["data"] + KANBAN_SNAPSHOT_COLS
WIP_SNAPSHOT_COLS = [
    "in progress",
    "waiting for fab.",
    "Fab test in progress",
]
KANBAN_SNAPSHOT_COLS_LEGACY = {
    "Backlog": "backlog",
    "Waiting": "Under analysis",
    "In Progress": "in progress",
    "Waiting Test": "waiting for fab.",
    "Test in progress": "Fab test in progress",
    "Done": "Acronimi done",
}

KANBAN_COLUMNS = [
    "Title",
    "Description",
    "Status",
    "Assignee",
    "Start Date",
    "End Date",
    "Estimate",
    "Priority",
    "Tags",
]
GIORNI_COL = "Giorni"
TAG_TEMPORALI_COL = "TAG Temporali"
INIZIO_LAVORAZIONE_COL = "InizioLavorazione(GG)"
WAITING_COL = "Waiting #"
TOTALE_LAVORAZIONE_COL = "Totale Lavorazione"
TAGS_SUM_COL = "Tags_sum"
PERIOD_SUM_HEADER = "Period SUM"
OUTPUT_SHEET = "data-all"
DATA_CHECK_SHEET = "data-check"
STATUS_ESCLUSI_DATA_CHECK = frozenset(
    {
        "Complete",
        "Abandoned",
        "Probably dismissed / delayed to 2027",
        "new - to be verified",
    }
)
STAT_SHEET = "stat"
GRAPH_SHEET = "graph"
CENTER_COLS = {3, 4, 7, 10, 11, 12, 13}  # C, D, G, J, K, L, M
COL_ESTIMATE = 7  # G
COL_TAGS_ORIG = 9  # I
COL_INIZIO_LAVORAZIONE = 10  # J
COL_WAITING = 11  # K
COL_TOTALE_LAVORAZIONE = 12  # L
COL_TAGS_SUM = 13  # M: Period SUM
COL_CARD_END = 13  # A–M: dati card (merge verticali per Title)
COL_GIORNI = 14  # N
COL_TAG = 15  # O
COL_LAST = 15
LEGENDA_COLONNE = [
    ("A–I", "", "dati card"),
    (
        "J",
        INIZIO_LAVORAZIONE_COL,
        "giorni dal tag '# Inizio Attivita' - <data>' a oggi",
    ),
    (
        "K",
        WAITING_COL,
        "giorni dall'ultimo tag '# Waiting -' a oggi (sfondo giallo)",
    ),
    (
        "L",
        TOTALE_LAVORAZIONE_COL,
        "giorni Working / Estimate: 0-50% verde, 51-80% giallo, 81-100% rosso pastello, >100% rosso acceso",
    ),
    ("M", PERIOD_SUM_HEADER, "tempo trascorso dalla prima attivita'"),
    ("N", GIORNI_COL, "per riga tag"),
    ("O", TAG_TEMPORALI_COL, "per riga tag"),
]
LEGENDA_COMMENTO_COL = 3
LEGENDA_COMMENTO_COL_FIN = 4
PASTEL_RED_BORDER = Side(style="medium", color="E8A0A0")
PASTEL_GREEN_FILL = PatternFill(fill_type="solid", fgColor="D9EAD3")
PASTEL_RED_FILL = PatternFill(fill_type="solid", fgColor="FFEBEE")
PASTEL_YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFFDE7")
BRIGHT_RED_FILL = PatternFill(fill_type="solid", fgColor="E53935")
DATE_IN_TAG_RE = re.compile(r"(\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2})")
ESTIMATE_TAG_RE = re.compile(
    r"^#\s*Estimate\s*[:-]?\s*([\d.,]+)",
    re.IGNORECASE,
)
INIZIO_ATTIVITA_TAG_RE = re.compile(
    r"^#\s*Inizio\s+Attivit[aà]'?\s*-\s*",
    re.IGNORECASE,
)
WAITING_TAG_RE = re.compile(
    r"^#\s*Waiting\s*-\s*",
    re.IGNORECASE,
)


def risolvi_percorso(nome_o_path):
    if os.path.isabs(nome_o_path):
        return nome_o_path
    return os.path.normpath(os.path.join(SCRIPT_DIR, nome_o_path))


def percorso_output_da_csv(input_path):
    base, _ = os.path.splitext(input_path)
    return base + ".xlsx"


def percorso_html_jkan_da_csv(input_path):
    """dbJKAN.html nella stessa cartella del CSV di input."""
    return os.path.join(os.path.dirname(input_path), DB_JKAN_HTML_NAME)


def normalizza_testo(valore, compatta_spazi=False):
    if valore is None:
        return ""
    testo = str(valore).replace("\xa0", " ").strip()
    if compatta_spazi:
        return re.sub(r"\s+", " ", testo)
    return testo


def is_kanban_header(row):
    if len(row) < len(KANBAN_COLUMNS):
        return False
    return [
        normalizza_testo(c, compatta_spazi=True)
        for c in row[: len(KANBAN_COLUMNS)]
    ] == KANBAN_COLUMNS


def is_card_row(row):
    return len(row) >= len(KANBAN_COLUMNS)


def leggi_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def estrai_card_kanban(rows):
    """Estrae le card da tutte le sezioni Kanban riconosciute nel CSV."""
    card = []
    sezioni = 0

    for i, row in enumerate(rows):
        if not is_kanban_header(row):
            continue

        sezioni += 1
        for j in range(i + 1, len(rows)):
            data_row = rows[j]
            if is_kanban_header(data_row):
                break
            if not is_card_row(data_row):
                break
            record = {
                col: normalizza_testo(data_row[idx]) if idx < len(data_row) else ""
                for idx, col in enumerate(KANBAN_COLUMNS)
            }
            if any(record.values()):
                card.append(record)

    return sezioni, card


def estrai_tag_temporali(description):
    """Restituisce le righe di Description che iniziano con '#'."""
    if not description:
        return []
    tag = []
    for line in str(description).splitlines():
        if line.lstrip().startswith("#"):
            tag.append(line.strip())
    return tag


def is_tag_estimate(tag):
    if not tag:
        return False
    return bool(ESTIMATE_TAG_RE.match(str(tag).strip()))


def is_tag_inizio_attivita(tag):
    if not tag:
        return False
    return bool(INIZIO_ATTIVITA_TAG_RE.match(str(tag).strip()))


def giorni_da_inizio_attivita(description, data_oggi=None):
    """Giorni tra la data nel tag '# Inizio Attivita' - <data>' e oggi."""
    if data_oggi is None:
        data_oggi = datetime.now().date()
    for tag in estrai_tag_temporali(description):
        if not is_tag_inizio_attivita(tag):
            continue
        date = estrai_date_da_tag(tag)
        if date:
            return (data_oggi - date[0]).days
    return None


def is_tag_waiting(tag):
    if not tag:
        return False
    return bool(WAITING_TAG_RE.match(str(tag).strip()))


def giorni_da_ultimo_tag_waiting(description, data_oggi=None):
    """Giorni tra la data nell'ultimo tag '# Waiting -' e oggi."""
    if data_oggi is None:
        data_oggi = datetime.now().date()
    tag_list = [
        tag
        for tag in estrai_tag_temporali(description)
        if not is_tag_estimate(tag)
    ]
    if not tag_list:
        return None
    ultimo = tag_list[-1]
    if not is_tag_waiting(ultimo):
        return None
    date = estrai_date_da_tag(ultimo)
    if not date:
        return None
    return (data_oggi - date[0]).days


def estrai_estimate_da_description(description):
    """Estrae il valore numerico dal tag '# Estimate' nella Description."""
    for tag in estrai_tag_temporali(description):
        match = ESTIMATE_TAG_RE.match(tag.strip())
        if match:
            return parse_numero(match.group(1))
    return None


def parse_data(valore):
    if valore is None:
        return None
    testo = normalizza_testo(valore)
    if not testo:
        return None

    if isinstance(valore, datetime):
        return valore.date()

    iso = re.match(
        r"^(\d{4})-(\d{2})-(\d{2})",
        testo,
    )
    if iso:
        return datetime(int(iso.group(1)), int(iso.group(2)), int(iso.group(3))).date()

    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(testo, fmt).date()
        except ValueError:
            continue
    return None


def estrai_date_da_tag(tag):
    """Estrae le date presenti nel testo di un TAG Temporale, in ordine di apparizione."""
    if not tag:
        return []
    date = []
    for match in DATE_IN_TAG_RE.finditer(str(tag)):
        parsed = parse_data(match.group(1))
        if parsed is not None:
            date.append(parsed)
    return date


def is_tag_done(tag):
    if not tag:
        return False
    return bool(re.search(r"\bdone\b", str(tag), re.IGNORECASE))


def tag_contiene_in_progress(tag):
    if not tag:
        return False
    return bool(re.search(r"\bin progress\b", str(tag), re.IGNORECASE))


def tag_contiene_lavorazione(tag):
    if not tag:
        return False
    return bool(re.search(r"\blavorazione\b", str(tag), re.IGNORECASE))


def tag_contiene_lavorazione_o_in_progress(tag):
    return tag_contiene_lavorazione(tag) or tag_contiene_in_progress(tag)


def giorni_da_tag_temporale(tag, tag_successivo=None, data_oggi=None):
    """
    Giorni tra la prima e la seconda data nel TAG Temporale.
    Con una sola data usa il tag successivo; se assente usa la data odierna,
    tranne per i tag Done che hanno solo la data di chiusura.
    """
    if data_oggi is None:
        data_oggi = datetime.now().date()

    date = estrai_date_da_tag(tag)
    if len(date) >= 2:
        return (date[1] - date[0]).days
    if tag_successivo:
        date_succ = estrai_date_da_tag(tag_successivo)
        if date and date_succ:
            return (date_succ[0] - date[0]).days
    if date and not is_tag_done(tag):
        return (data_oggi - date[0]).days
    return None


def parse_numero(valore):
    if valore is None or valore == "":
        return None
    if isinstance(valore, (int, float)):
        return float(valore)
    testo = normalizza_testo(valore).replace(",", ".")
    try:
        return float(testo)
    except ValueError:
        return None


def espandi_card_con_tag(card):
    """Replica ogni card per tag temporale; aggiunge colonne Giorni e TAG Temporali."""
    righe = []
    for record in card:
        nuova_base = dict(record)
        estimate = estrai_estimate_da_description(record.get("Description", ""))
        nuova_base["Estimate"] = estimate if estimate is not None else ""
        inizio_lavorazione = giorni_da_inizio_attivita(
            record.get("Description", "")
        )
        nuova_base[INIZIO_LAVORAZIONE_COL] = inizio_lavorazione
        waiting_gg = giorni_da_ultimo_tag_waiting(record.get("Description", ""))
        nuova_base[WAITING_COL] = waiting_gg

        tag_list = [
            tag
            for tag in estrai_tag_temporali(record.get("Description", ""))
            if not is_tag_estimate(tag)
        ]
        if not tag_list:
            nuova = dict(nuova_base)
            nuova[TOTALE_LAVORAZIONE_COL] = None
            nuova[TAGS_SUM_COL] = None
            nuova[GIORNI_COL] = None
            nuova[TAG_TEMPORALI_COL] = ""
            righe.append(nuova)
            continue
        for i, tag in enumerate(tag_list):
            tag_next = tag_list[i + 1] if i + 1 < len(tag_list) else None
            nuova = dict(nuova_base)
            nuova[TOTALE_LAVORAZIONE_COL] = None
            nuova[TAGS_SUM_COL] = None
            nuova[GIORNI_COL] = giorni_da_tag_temporale(tag, tag_next)
            nuova[TAG_TEMPORALI_COL] = tag
            righe.append(nuova)
    return righe


def colonne_output():
    return KANBAN_COLUMNS + [
        INIZIO_LAVORAZIONE_COL,
        WAITING_COL,
        TOTALE_LAVORAZIONE_COL,
        TAGS_SUM_COL,
        GIORNI_COL,
        TAG_TEMPORALI_COL,
    ]


def somma_giorni_gruppo(ws, start, end, filtro_tag=None):
    totale = 0.0
    ha_valori = False
    for row in range(start, end + 1):
        tag = ws.cell(row=row, column=COL_TAG).value
        if filtro_tag is not None and not filtro_tag(tag):
            continue
        val = parse_numero(ws.cell(row=row, column=COL_GIORNI).value)
        if val is not None:
            totale += val
            ha_valori = True
    return totale if ha_valori else None


def percentuale_working_su_estimate(estimate, tot_lavorazione):
    if estimate is None or tot_lavorazione is None or estimate <= 0:
        return None
    return (tot_lavorazione / estimate) * 100


def applica_colore_confronto_estimate(cella, estimate, valore):
    if estimate is None or valore is None:
        return
    if estimate >= valore:
        cella.fill = PASTEL_GREEN_FILL
    else:
        cella.fill = PASTEL_RED_FILL


def applica_colore_totale_lavorazione(cella, estimate, tot_lavorazione):
    percentuale = percentuale_working_su_estimate(estimate, tot_lavorazione)
    if percentuale is None:
        return
    if percentuale <= 50:
        cella.fill = PASTEL_GREEN_FILL
    elif percentuale <= 80:
        cella.fill = PASTEL_YELLOW_FILL
    elif percentuale <= 100:
        cella.fill = PASTEL_RED_FILL
    else:
        cella.fill = BRIGHT_RED_FILL


def applica_colore_waiting(ws, start):
    cella = ws.cell(row=start, column=COL_WAITING)
    if parse_numero(cella.value) is not None:
        cella.fill = PASTEL_YELLOW_FILL


def applica_totali_gruppo(ws, start, end):
    center = Alignment(horizontal="center", vertical="center")

    tot_lavorazione = somma_giorni_gruppo(
        ws, start, end, tag_contiene_lavorazione_o_in_progress
    )
    cella_lav = ws.cell(row=start, column=COL_TOTALE_LAVORAZIONE)
    cella_lav.value = tot_lavorazione
    cella_lav.alignment = center

    somma = somma_giorni_gruppo(ws, start, end, tag_contiene_in_progress)
    cella = ws.cell(row=start, column=COL_TAGS_SUM)
    cella.value = somma
    cella.alignment = center

    estimate = parse_numero(ws.cell(row=start, column=COL_ESTIMATE).value)
    applica_colore_totale_lavorazione(cella_lav, estimate, tot_lavorazione)
    applica_colore_confronto_estimate(cella, estimate, somma)

    cella_waiting = ws.cell(row=start, column=COL_WAITING)
    cella_waiting.alignment = center
    applica_colore_waiting(ws, start)


def gruppi_righe_per_title(ws):
    """Raggruppa righe dati (dalla 2) consecutivi con stesso Title in colonna A."""
    gruppi = []
    max_row = ws.max_row
    if max_row < 2:
        return gruppi

    row = 2
    while row <= max_row:
        title = ws.cell(row=row, column=1).value
        start = row
        while row + 1 <= max_row and ws.cell(row=row + 1, column=1).value == title:
            row += 1
        gruppi.append((start, row))
        row += 1
    return gruppi


def applica_bordo_gruppo(ws, min_row, max_row, min_col, max_col, side):
    """Applica un bordo perimetrale al rettangolo di celle indicato."""
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            border = Border(
                left=side if col == min_col else None,
                right=side if col == max_col else None,
                top=side if row == min_row else None,
                bottom=side if row == max_row else None,
            )
            cell.border = border


def formatta_foglio_card(ws):
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    middle = Alignment(vertical="center", wrap_text=True)
    gruppi = gruppi_righe_per_title(ws)

    for row in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=COL_LAST,
    ):
        for cell in row:
            if cell.row == 1 or cell.column in CENTER_COLS:
                cell.alignment = center
            else:
                cell.alignment = middle

    for start, end in gruppi:
        if end > start:
            for col in range(1, COL_CARD_END + 1):
                ws.merge_cells(
                    start_row=start,
                    start_column=col,
                    end_row=end,
                    end_column=col,
                )
                merged = ws.cell(row=start, column=col)
                if col in CENTER_COLS:
                    merged.alignment = center
                else:
                    merged.alignment = middle
        applica_totali_gruppo(ws, start, end)
        applica_bordo_gruppo(ws, start, end, 1, COL_LAST, PASTEL_RED_BORDER)

    return gruppi


def riepilogo_da_gruppi(ws, gruppi):
    riepilogo = []
    for start, _end in gruppi:
        riepilogo.append(
            {
                "title": normalizza_testo(ws.cell(row=start, column=1).value),
                "status": normalizza_testo(ws.cell(row=start, column=3).value),
                "estimate": parse_numero(ws.cell(row=start, column=COL_ESTIMATE).value),
                "period_sum": parse_numero(
                    ws.cell(row=start, column=COL_TAGS_SUM).value
                ),
            }
        )
    return riepilogo


def tempo_mancante(estimate, period_sum):
    if estimate is None or period_sum is None:
        return None
    return max(0.0, estimate - period_sum)


def aggiungi_footer_data(ws, gruppi):
    center = Alignment(horizontal="center", vertical="center")
    data_last = ws.max_row
    totals_row = data_last + 2
    ts_row = totals_row + 1
    legend_start = ts_row + 2

    tot_estimate = 0.0
    tot_lavorazione = 0.0
    tot_period = 0.0
    ha_estimate = False
    ha_lavorazione = False
    ha_period = False
    for start, _end in gruppi:
        estimate = parse_numero(ws.cell(row=start, column=COL_ESTIMATE).value)
        lavorazione = parse_numero(
            ws.cell(row=start, column=COL_TOTALE_LAVORAZIONE).value
        )
        period_sum = parse_numero(ws.cell(row=start, column=COL_TAGS_SUM).value)
        if estimate is not None:
            tot_estimate += estimate
            ha_estimate = True
        if lavorazione is not None:
            tot_lavorazione += lavorazione
            ha_lavorazione = True
        if period_sum is not None:
            tot_period += period_sum
            ha_period = True

    ws.cell(row=totals_row, column=1).value = len(gruppi)
    ws.cell(row=totals_row, column=1).alignment = center
    if ha_estimate:
        ws.cell(row=totals_row, column=COL_ESTIMATE).value = tot_estimate
        ws.cell(row=totals_row, column=COL_ESTIMATE).alignment = center
    if ha_lavorazione:
        ws.cell(row=totals_row, column=COL_TOTALE_LAVORAZIONE).value = tot_lavorazione
        ws.cell(row=totals_row, column=COL_TOTALE_LAVORAZIONE).alignment = center
    if ha_period:
        ws.cell(row=totals_row, column=COL_TAGS_SUM).value = tot_period
        ws.cell(row=totals_row, column=COL_TAGS_SUM).alignment = center

    ws.cell(row=ts_row, column=1).value = datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )
    aggiungi_legenda_colonne(ws, legend_start)


def aggiungi_legenda_colonne(ws, start_row):
    """Legenda colonne: lettera, titolo, commento (C–D senza a capo in C)."""
    left = Alignment(vertical="center", wrap_text=False)
    center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    commento = Alignment(vertical="center", wrap_text=False)
    row = start_row
    for lettera, titolo, testo_commento in LEGENDA_COLONNE:
        cell_lettera = ws.cell(row=row, column=1)
        cell_lettera.value = lettera
        cell_lettera.alignment = center

        cell_titolo = ws.cell(row=row, column=2)
        cell_titolo.value = titolo
        cell_titolo.alignment = left

        ws.merge_cells(
            start_row=row,
            start_column=LEGENDA_COMMENTO_COL,
            end_row=row,
            end_column=LEGENDA_COMMENTO_COL_FIN,
        )
        cell_commento = ws.cell(row=row, column=LEGENDA_COMMENTO_COL)
        cell_commento.value = testo_commento
        cell_commento.alignment = commento
        row += 1


def crea_foglio_stat(wb, riepilogo):
    ws = wb.create_sheet(STAT_SHEET)
    center = Alignment(horizontal="center", vertical="center")
    ws.cell(row=1, column=1).value = "Status"
    ws.cell(row=1, column=2).value = "Conteggio"
    ws.cell(row=1, column=1).alignment = center
    ws.cell(row=1, column=2).alignment = center

    conteggi = Counter(
        r["status"] if r["status"] else "(vuoto)" for r in riepilogo
    )
    for idx, (status, count) in enumerate(sorted(conteggi.items()), start=2):
        ws.cell(row=idx, column=1).value = status
        ws.cell(row=idx, column=2).value = count
        ws.cell(row=idx, column=1).alignment = center
        ws.cell(row=idx, column=2).alignment = center


def crea_foglio_graph(wb, riepilogo):
    ws = wb.create_sheet(GRAPH_SHEET)
    center = Alignment(horizontal="center", vertical="center")
    ws.cell(row=1, column=1).value = "Acronimo"
    ws.cell(row=1, column=2).value = "Tempo mancante"
    ws.cell(row=1, column=1).alignment = center
    ws.cell(row=1, column=2).alignment = center

    dati = []
    for r in riepilogo:
        if not r["title"]:
            continue
        mancante = tempo_mancante(r["estimate"], r["period_sum"])
        dati.append((r["title"], mancante if mancante is not None else 0))

    if not dati:
        return

    for idx, (title, mancante) in enumerate(dati, start=2):
        ws.cell(row=idx, column=1).value = title
        ws.cell(row=idx, column=2).value = mancante
        ws.cell(row=idx, column=1).alignment = center
        ws.cell(row=idx, column=2).alignment = center

    last_row = len(dati) + 1
    chart = BarChart()
    chart.type = "col"
    chart.title = "Tempo mancante per acronimo"
    chart.y_axis.title = "Giorni"
    chart.x_axis.title = "Acronimo"
    chart.height = 12
    chart.width = 20

    data_ref = Reference(ws, min_col=2, min_row=1, max_row=last_row)
    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=last_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "D2")


def status_escluso_da_data_check(status):
    return normalizza_testo(status, compatta_spazi=True) in STATUS_ESCLUSI_DATA_CHECK


def filtra_card_per_data_check(card):
    return [
        record
        for record in card
        if not status_escluso_da_data_check(record.get("Status", ""))
    ]


def formatta_foglio_dati(ws):
    ws.cell(row=1, column=COL_TOTALE_LAVORAZIONE).value = TOTALE_LAVORAZIONE_COL
    ws.cell(row=1, column=COL_TAGS_SUM).value = PERIOD_SUM_HEADER
    gruppi = formatta_foglio_card(ws)
    aggiungi_footer_data(ws, gruppi)
    return gruppi


def scrivi_excel(card, output_path):
    righe_all = espandi_card_con_tag(card)
    df_all = pd.DataFrame(righe_all, columns=colonne_output())
    card_check = filtra_card_per_data_check(card)
    righe_check = espandi_card_con_tag(card_check)
    df_check = pd.DataFrame(righe_check, columns=colonne_output())

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
        df_check.to_excel(writer, sheet_name=DATA_CHECK_SHEET, index=False)

    wb = load_workbook(output_path)
    gruppi = formatta_foglio_dati(wb[OUTPUT_SHEET])
    formatta_foglio_dati(wb[DATA_CHECK_SHEET])
    riepilogo = riepilogo_da_gruppi(wb[OUTPUT_SHEET], gruppi)
    crea_foglio_stat(wb, riepilogo)
    crea_foglio_graph(wb, riepilogo)
    wb.save(output_path)


def estrai_data_da_nome_file(path):
    """Estrae una data ISO (yyyy-mm-dd) dal nome del file di input."""
    nome = os.path.basename(path)
    match = DATA_IN_FILENAME_RE.search(nome)
    if not match:
        raise ValueError(
            f"Impossibile estrarre la data dal nome file: {nome!r}. "
            "Atteso un pattern yyyy-mm-dd o yyyy/mm/dd nel nome."
        )
    anno, mese, giorno = (int(match.group(i)) for i in range(1, 4))
    return datetime(anno, mese, giorno).date()


def is_attesa_test_o_fab(status, description=""):
    """True se la card è in attesa test/fabbrica (colonna waiting for fab.)."""
    testo = normalizza_testo(status, compatta_spazi=True).lower()
    if re.search(r"waiting\s*(for\s*fab\.?|test)|attesa\s*(fab\.?|test)", testo):
        return True
    desc = normalizza_testo(description).lower()
    if re.search(r"attesa\s*(test|fab\.?)", desc):
        return True
    tag_list = [
        tag
        for tag in estrai_tag_temporali(description)
        if not is_tag_estimate(tag)
    ]
    if tag_list and re.search(
        r"attesa\s*(test|fab\.?)", tag_list[-1], re.IGNORECASE
    ):
        return True
    return False


def classifica_colonna_kanban(status, description=""):
    """Mappa Status (e Description) a una colonna Kanban dello snapshot."""
    testo = normalizza_testo(status, compatta_spazi=True).lower()
    if not testo:
        return None
    if re.search(r"\bcomplete\b", testo):
        return None
    if re.search(r"\bacronimi\s*done\b", testo) or testo == "done":
        return "Acronimi done"
    if re.search(r"fab\s*test\s*in\s*progress", testo) or re.search(
        r"test\s*in\s*progress", testo
    ) or testo == "test":
        return "Fab test in progress"
    if re.search(r"in\s*progress", testo) or re.search(r"lavorazione", testo):
        return "in progress"
    if is_attesa_test_o_fab(status, description):
        return "waiting for fab."
    if re.search(r"\bbacklog\b", testo):
        return "backlog"
    if re.search(r"under\s*analysis", testo):
        return "Under analysis"
    if re.search(r"\bwaiting\b", testo) or re.search(
        r"attesa\s*(lavorazione|intesa)?", testo
    ):
        return "Under analysis"
    return None


def conteggio_per_colonna_kanban(card):
    """Conta le card per colonna Kanban (Under analysis, backlog, …)."""
    conteggi = Counter({col: 0 for col in KANBAN_SNAPSHOT_COLS})
    non_mappate = 0
    for record in card:
        colonna = classifica_colonna_kanban(
            record.get("Status", ""),
            record.get("Description", ""),
        )
        if colonna is None:
            non_mappate += 1
            continue
        conteggi[colonna] += 1
    return conteggi, non_mappate


def _migra_colonne_db_jkan(df):
    """Converte colonne legacy (Backlog, Waiting, …) al nuovo schema Kanban."""
    for old, new in KANBAN_SNAPSHOT_COLS_LEGACY.items():
        if old not in df.columns:
            continue
        if new in df.columns:
            df[new] = pd.to_numeric(df[new], errors="coerce").fillna(0).astype(int)
            df[new] += pd.to_numeric(df[old], errors="coerce").fillna(0).astype(int)
        else:
            df[new] = pd.to_numeric(df[old], errors="coerce").fillna(0).astype(int)
        df = df.drop(columns=[old])
    return df


def carica_db_jkan(csv_path=DB_JKAN_CSV):
    """Carica lo storico snapshot; restituisce DataFrame ordinato per data."""
    if not os.path.exists(csv_path):
        return pd.DataFrame(columns=DB_JKAN_HEADERS)
    df = pd.read_csv(csv_path, dtype={"data": str})
    df = _migra_colonne_db_jkan(df)
    for col in DB_JKAN_HEADERS:
        if col not in df.columns:
            df[col] = 0 if col != "data" else ""
    df = df[DB_JKAN_HEADERS].copy()
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    df = df.dropna(subset=["data"]).sort_values("data").reset_index(drop=True)
    for col in KANBAN_SNAPSHOT_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    return df


def somma_colonne_scope(df):
    """Somma Under analysis … Acronimi done (scope tracciato per snapshot)."""
    return df[KANBAN_SNAPSHOT_COLS].sum(axis=1)


def _etichette_asse_tempo(df):
    """Restituisce date (dd/mm/yyyy) e numeri settimana ISO (Wnn) per l'asse X."""
    dates = []
    weeks = []
    for ts in df["data"]:
        giorno = ts.date() if hasattr(ts, "date") else ts
        dates.append(giorno.strftime("%d/%m/%Y"))
        weeks.append(f"W{giorno.isocalendar()[1]:02d}")
    return dates, weeks


def _serie_totali(df):
    return somma_colonne_scope(df).tolist()


def _serie_wip(df):
    return df[WIP_SNAPSHOT_COLS].sum(axis=1).tolist()


def aggiorna_db_jkan(data_snapshot, conteggi, csv_path=DB_JKAN_CSV):
    """Inserisce o sovrascrive la riga per la data indicata."""
    df = carica_db_jkan(csv_path)
    riga = {"data": pd.Timestamp(data_snapshot)}
    for col in KANBAN_SNAPSHOT_COLS:
        riga[col] = int(conteggi.get(col, 0))

    data_key = pd.Timestamp(data_snapshot).normalize()
    if df.empty:
        df = pd.DataFrame([riga])
    else:
        mask = df["data"].dt.normalize() == data_key
        if mask.any():
            for col in DB_JKAN_HEADERS:
                df.loc[mask, col] = riga[col]
        else:
            df = pd.concat([df, pd.DataFrame([riga])], ignore_index=True)

    df = df.sort_values("data").reset_index(drop=True)
    export = df.copy()
    export["data"] = export["data"].dt.strftime("%Y-%m-%d")
    export.to_csv(csv_path, index=False)
    return df


def _serie_velocita(df):
    done = df["Acronimi done"].astype(int).tolist()
    if len(done) <= 1:
        return [0] * len(done)
    return [0] + [done[i] - done[i - 1] for i in range(1, len(done))]


def _dati_grafici_jkan(df):
    dates, weeks = _etichette_asse_tempo(df)
    n = len(df)
    return {
        "dates": dates,
        "weeks": weeks,
        "burnup": {
            "done": df["Acronimi done"].astype(int).tolist(),
            "scope": _serie_totali(df),
            "target": [SCOPE_TOTALE_CARD] * n,
        },
        "wip": _serie_wip(df),
        "velocity": _serie_velocita(df),
        "stacked": {
            col: df[col].astype(int).tolist() for col in KANBAN_SNAPSHOT_COLS
        },
    }


def genera_html_jkan(df, html_path, data_ultimo_snapshot=None):
    """Genera report HTML con grafici burnup e metriche di avanzamento."""
    if df.empty:
        return

    dati = _dati_grafici_jkan(df)
    chart_json = json.dumps(dati, ensure_ascii=False)
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    ultima = (
        data_ultimo_snapshot.strftime("%d/%m/%Y")
        if data_ultimo_snapshot
        else df["data"].iloc[-1].strftime("%d/%m/%Y")
    )
    tot_ultimo = int(somma_colonne_scope(df).iloc[-1])
    done_ultimo = int(df["Acronimi done"].iloc[-1])
    wip_ultimo = int(df[WIP_SNAPSHOT_COLS].iloc[-1].sum())

    tabella_rows = []
    for _, row in df.iterrows():
        cells = [row["data"].strftime("%d/%m/%Y")]
        cells += [str(int(row[col])) for col in KANBAN_SNAPSHOT_COLS]
        cells.append(str(int(somma_colonne_scope(df).loc[row.name])))
        tabella_rows.append(
            "<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>"
        )
    tabella_html = "\n".join(tabella_rows)

    html = f"""<!DOCTYPE html>
<html lang="it">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>JKAN — Burnup e metriche</title>
  <style>
    :root {{
      --bg: #f8fafc; --card: #fff; --text: #1e293b; --muted: #64748b;
      --accent: #1a56db; --border: #e2e8f0;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0; font-family: system-ui, -apple-system, sans-serif;
      background: var(--bg); color: var(--text); line-height: 1.5;
    }}
    main {{ max-width: 1100px; margin: 0 auto; padding: 1.5rem 1rem 3rem; }}
    h1 {{ margin: 0 0 .25rem; font-size: 1.6rem; }}
    .sub {{ color: var(--muted); margin: 0 0 1.5rem; }}
    .kpis {{
      display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
      gap: .75rem; margin-bottom: 1.5rem;
    }}
    .kpi {{
      background: var(--card); border: 1px solid var(--border);
      border-radius: 8px; padding: .85rem 1rem;
    }}
    .kpi b {{ display: block; font-size: 1.5rem; color: var(--accent); }}
    .kpi span {{ font-size: .85rem; color: var(--muted); }}
    section {{
      background: var(--card); border: 1px solid var(--border);
      border-radius: 8px; padding: 1rem 1.25rem; margin-bottom: 1.25rem;
    }}
    section h2 {{ margin: 0 0 .75rem; font-size: 1.1rem; }}
    .chart-wrap {{ position: relative; height: 320px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: .9rem; }}
    th, td {{ border: 1px solid var(--border); padding: .45rem .6rem; text-align: center; }}
    th {{ background: #f1f5f9; }}
    td:first-child, th:first-child {{ text-align: left; }}
  </style>
</head>
<body>
<main>
  <h1>JKAN — Burnup e metriche</h1>
  <p class="sub">Generato il {now_str} · ultimo snapshot: {ultima}</p>

  <div class="kpis">
    <div class="kpi"><b>{tot_ultimo}</b><span>Scope tracciato (somma colonne)</span></div>
    <div class="kpi"><b>{SCOPE_TOTALE_CARD}</b><span>Target card totali</span></div>
    <div class="kpi"><b>{done_ultimo}</b><span>Acronimi done</span></div>
    <div class="kpi"><b>{wip_ultimo}</b><span>WIP (in progress + waiting for fab. + Fab test)</span></div>
    <div class="kpi"><b>{len(df)}</b><span>Snapshot registrati</span></div>
  </div>

  <section id="burnup">
    <h2>Burnup — Acronimi done vs scope</h2>
    <p class="sub">Scope tracciato = somma di Under analysis, backlog, in progress, waiting for fab., Fab test in progress e Acronimi done. Linea tratteggiata: target {SCOPE_TOTALE_CARD} card.</p>
    <div class="chart-wrap"><canvas id="chart-burnup"></canvas></div>
  </section>

  <section id="stacked">
    <h2>Distribuzione card per colonna Kanban</h2>
    <div class="chart-wrap"><canvas id="chart-stacked"></canvas></div>
  </section>

  <section id="wip">
    <h2>WIP — lavoro in corso</h2>
    <p class="sub">Somma di in progress, waiting for fab. e Fab test in progress (esclusi Under analysis, backlog e Acronimi done).</p>
    <div class="chart-wrap"><canvas id="chart-wip"></canvas></div>
  </section>

  <section id="velocity">
    <h2>Velocità — acronimi completati per snapshot</h2>
    <p class="sub">Incremento di Acronimi done rispetto allo snapshot precedente.</p>
    <div class="chart-wrap"><canvas id="chart-velocity"></canvas></div>
  </section>

  <section id="tabella">
    <h2>Storico dbJKAN.csv</h2>
    <table>
      <thead>
        <tr>
          <th>Data</th>
          <th>Under analysis</th>
          <th>backlog</th>
          <th>in progress</th>
          <th>waiting for fab.</th>
          <th>Fab test in progress</th>
          <th>Acronimi done</th>
          <th>Totale</th>
        </tr>
      </thead>
      <tbody>
        {tabella_html}
      </tbody>
    </table>
  </section>
</main>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js"></script>
<script>
(function() {{
  const D = {chart_json};
  const axisLabels = D.dates.map(function(d, i) {{ return [d, D.weeks[i]]; }});
  const STACK_COLORS = {{
    "Under analysis": "#6366f1",
    "backlog": "#94a3b8",
    "in progress": "#1a56db",
    "waiting for fab.": "#fb923c",
    "Fab test in progress": "#8b5cf6",
    "Acronimi done": "#10b981"
  }};
  const baseOpts = {{
    responsive: true,
    maintainAspectRatio: false,
    plugins: {{ legend: {{ position: "top" }} }},
    scales: {{
      x: {{ ticks: {{ maxRotation: 0, autoSkip: true }} }},
      y: {{ beginAtZero: true, ticks: {{ precision: 0 }} }}
    }}
  }};

  new Chart(document.getElementById("chart-burnup"), {{
    type: "line",
    data: {{
      labels: axisLabels,
      datasets: [
        {{
          label: "Acronimi done",
          data: D.burnup.done,
          borderColor: "#10b981",
          backgroundColor: "rgba(16,185,129,0.15)",
          fill: true, tension: 0.25, pointRadius: 4
        }},
        {{
          label: "Scope tracciato",
          data: D.burnup.scope,
          borderColor: "#64748b",
          borderDash: [6, 4],
          backgroundColor: "transparent",
          fill: false, tension: 0.25, pointRadius: 3
        }},
        {{
          label: "Target card ({SCOPE_TOTALE_CARD})",
          data: D.burnup.target,
          borderColor: "#cbd5e1",
          borderDash: [2, 4],
          backgroundColor: "transparent",
          fill: false, tension: 0, pointRadius: 0
        }}
      ]
    }},
    options: baseOpts
  }});

  new Chart(document.getElementById("chart-stacked"), {{
    type: "bar",
    data: {{
      labels: axisLabels,
      datasets: Object.keys(D.stacked).map(function(col) {{
        return {{
          label: col,
          data: D.stacked[col],
          backgroundColor: STACK_COLORS[col] || "#cbd5e1",
          stack: "kanban"
        }};
      }})
    }},
    options: Object.assign({{}}, baseOpts, {{
      scales: {{
        x: {{ stacked: true, ticks: {{ maxRotation: 0, autoSkip: true }} }},
        y: {{ stacked: true, beginAtZero: true, ticks: {{ precision: 0 }} }}
      }}
    }})
  }});

  new Chart(document.getElementById("chart-wip"), {{
    type: "line",
    data: {{
      labels: axisLabels,
      datasets: [{{
        label: "WIP",
        data: D.wip,
        borderColor: "#1a56db",
        backgroundColor: "rgba(26,86,219,0.12)",
        fill: true, tension: 0.25, pointRadius: 4
      }}]
    }},
    options: baseOpts
  }});

  new Chart(document.getElementById("chart-velocity"), {{
    type: "bar",
    data: {{
      labels: axisLabels,
      datasets: [{{
        label: "Acronimi completati",
        data: D.velocity,
        backgroundColor: "#10b981"
      }}]
    }},
    options: baseOpts
  }});
}})();
</script>
</body>
</html>"""

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)


def elabora(input_csv):
    input_path = risolvi_percorso(input_csv)
    output_path = percorso_output_da_csv(input_path)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File di input non trovato: {input_path}")

    rows = leggi_csv(input_path)
    sezioni, card = estrai_card_kanban(rows)

    if not card:
        raise ValueError(
            f"Nessuna card Kanban trovata nel file {input_path}"
        )

    scrivi_excel(card, output_path)

    data_snapshot = estrai_data_da_nome_file(input_path)
    conteggi, non_mappate = conteggio_per_colonna_kanban(card)
    df_storico = aggiorna_db_jkan(data_snapshot, conteggi)
    html_path = percorso_html_jkan_da_csv(input_path)
    genera_html_jkan(df_storico, html_path, data_ultimo_snapshot=data_snapshot)

    righe_output = len(espandi_card_con_tag(card))
    print(f"Sezioni kanban: {sezioni}")
    print(f"Card estratte: {len(card)}")
    print(f"Righe output: {righe_output}")
    print(f"Output: {output_path}")
    print(f"Snapshot {data_snapshot.isoformat()}: {dict(conteggi)}")
    if non_mappate:
        print(f"Card con Status non mappato: {non_mappate}")
    print(f"Database: {DB_JKAN_CSV}")
    print(f"Grafici: {html_path}")


def crea_parser():
    epilog = """\
Esempi:
  python elabora_JKAN.py
  python elabora_JKAN.py 2026-06-06-WIP.csv
  python elabora_JKAN.py /percorso/export-miro.csv

Parametri:
  input.csv   Export CSV MIRO (opzionale).
              Default: 2026-06-06-WIP.csv nella cartella dello script.

Output:
  <input>.xlsx — stesso percorso del CSV, estensione .xlsx.
  dbJKAN.csv   — storico snapshot (cartella dello script).
  dbJKAN.html  — grafici burnup (stessa cartella del CSV di input).
"""
    parser = argparse.ArgumentParser(
        description=(
            "Estrae le card da tutte le Kanban in un export MIRO "
            "e produce un file Excel con metriche temporali."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=epilog,
    )
    parser.add_argument(
        "input_csv",
        nargs="?",
        default="2026-06-06-WIP.csv",
        metavar="input.csv",
        help="export CSV MIRO (default: %(default)s)",
    )
    return parser


def main(argv=None):
    args = crea_parser().parse_args(argv)
    elabora(args.input_csv)


if __name__ == "__main__":
    main()
