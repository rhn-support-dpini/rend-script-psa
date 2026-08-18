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
    Fogli: data-all (tutte le card), data-export (export ridotto), stat.
    dbJKAN.csv — storico snapshot colonne Kanban (cartella dello script).
    <input>.html — report Scrum/Kanban (stesso percorso del .xlsx prodotto).
    Le righe Description con prefisso "#" generano sotto-righe da colonna P (TAG Temporali);
    A–N sono merge verticali per Title, con bordo rosso pastello per card.
    J (InizioLavorazione(GG)): giorni dal tag "# Inizio Attivita'" a oggi, se presente.
    K (Totale Waiting): somma giornate tag "# Waiting -" in col. P.
    L (Totale Lavorazione): somma tag "# Working -"; sfondo per % su Estimate (col. G).
    M (Rework time): somma col. O (Giorni) per tag "# Rework" in col. P.
    N (Fix time): somma giornate tag "# Fix" in col. P.
    O (Giorni): giornate per riga tag. P (TAG Temporali): testo del tag per riga.
    Colonna G (Estimate): valore numerico dal tag "# Estimate" in Description, non dal CSV.
    Giornate lavorative (lun-ven); nei tag a due date inizio incluso, fine esclusa;
    se manca la 2ª data si usa oggi (incluso), eccetto tag Done.
"""

import argparse
import csv
import html
import json
import os
import re
import sys
from collections import Counter
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, PatternFill, Side

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_JKAN_CSV = os.path.join(SCRIPT_DIR, "dbJKAN.csv")
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
KANBAN_DISPLAY_LABELS = {
    "Under analysis": "Under analysis",
    "backlog": "Backlog",
    "in progress": "In progress",
    "waiting for fab.": "Waiting for fabric",
    "Fab test in progress": "Fab. Test in progress",
    "Acronimi done": "Acronimi Done",
}
KANBAN_CHART_COLORS = {
    "Under analysis": "#6366f1",
    "backlog": "#94a3b8",
    "in progress": "#1a56db",
    "waiting for fab.": "#fb923c",
    "Fab test in progress": "#8b5cf6",
    "Acronimi done": "#10b981",
}
KANBAN_CHART_FILL = {
    col: f"rgba({int(KANBAN_CHART_COLORS[col][1:3], 16)},{int(KANBAN_CHART_COLORS[col][3:5], 16)},{int(KANBAN_CHART_COLORS[col][5:7], 16)},0.15)"
    for col in KANBAN_SNAPSHOT_COLS
}
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
FIX_TIME_COL = "Fix time"
GIORNI_COL = "Giorni"
TAG_TEMPORALI_COL = "TAG Temporali"
INIZIO_LAVORAZIONE_COL = "InizioLavorazione(GG)"
TOTALE_WAITING_COL = "Totale Waiting"
TOTALE_LAVORAZIONE_COL = "Totale Lavorazione"
REWORK_TIME_COL = "Rework time"
OUTPUT_SHEET = "data-all"
DATA_EXPORT_SHEET = "data-export"
COLONNE_EXPORT = ["Title", "Status", "Start Date", "End Date", "Tags"]
EXPORT_COL_LAST = 5  # E: Tags
STATUS_ESCLUSI_EXPORT = frozenset(
    {
        "Attività nuove",
        "Attività in Backlog",
        "Attività in Progress",
        "Attività Complete",
        "Attività Abbandonate",
    }
)
STAT_SHEET = "stat"
CENTER_COLS = {3, 4, 7, 10, 11, 12, 13, 14, 15, 16}  # C, D, G, J–O
COL_ESTIMATE = 7  # G — Estimate (confronto % con col. L)
COL_TAGS_ORIG = 9  # I
COL_INIZIO_LAVORAZIONE = 10  # J
COL_TOTALE_WAITING = 11  # K
COL_TOTALE_LAVORAZIONE = 12  # L
COL_REWORK_TIME = 13  # M
COL_FIX_TIME = 14  # N
COL_CARD_END = 14  # A–N: dati card (merge verticali per Title)
COL_GIORNI = 15  # O
COL_TAG = 16  # P
COL_LAST = 16
LEGENDA_COLONNE = [
    ("A–I", "", "dati card"),
    (
        "J",
        INIZIO_LAVORAZIONE_COL,
        "giornate lavorative (lun-ven) dal tag '# Inizio Attivita' - <data>' a oggi",
    ),
    (
        "K",
        TOTALE_WAITING_COL,
        "somma giornate lavorative (lun-ven) dei tag '# Waiting -' in colonna P",
    ),
    (
        "L",
        TOTALE_LAVORAZIONE_COL,
        "somma giornate lavorative (lun-ven) tag '# Working -' in colonna P; "
        "sfondo L: (L/Estimate)×100 — ≤50% verde, 51–80% giallo, 81–100% rosso pastello, >100% rosso acceso",
    ),
    (
        "M",
        REWORK_TIME_COL,
        "somma colonna O (Giorni) per righe con tag '# Rework' in colonna P",
    ),
    (
        "N",
        FIX_TIME_COL,
        "somma giornate lavorative (lun-ven) dei tag '# Fix' in colonna P",
    ),
    (
        "O",
        GIORNI_COL,
        "giornate lavorative (lun-ven) per riga tag; due date: inizio incluso, fine esclusa",
    ),
    ("P", TAG_TEMPORALI_COL, "per riga tag"),
]
LEGENDA_COMMENTO_COL = 3
LEGENDA_COMMENTO_COL_FIN = 4
PASTEL_RED_BORDER = Side(style="medium", color="E8A0A0")
PASTEL_GREEN_FILL = PatternFill(fill_type="solid", fgColor="D9EAD3")
PASTEL_RED_FILL = PatternFill(fill_type="solid", fgColor="FFEBEE")
PASTEL_YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFFDE7")
PASTEL_GRAY_FILL = PatternFill(fill_type="solid", fgColor="E8E8E8")
PASTEL_BLUE_FILL = PatternFill(fill_type="solid", fgColor="D6EAF8")
EXPORT_BRIGHT_GREEN_FILL = PatternFill(fill_type="solid", fgColor="81C784")
BRIGHT_RED_FILL = PatternFill(fill_type="solid", fgColor="E53935")
EXPORT_BLUE_BORDER = Side(style="medium", color="2563EB")
EXPORT_COL_DIVIDER = Side(style="thin", color="93C5FD")
EXPORT_STATUS_FILLS = {
    "abandoned": PASTEL_GRAY_FILL,
    "fab. test in progress": PASTEL_YELLOW_FILL,
    "under analysis": PASTEL_RED_FILL,
    "in progress": PASTEL_GREEN_FILL,
    "backlog": PASTEL_BLUE_FILL,
    "acronimi done": EXPORT_BRIGHT_GREEN_FILL,
}
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
WORKING_TAG_RE = re.compile(
    r"^#\s*Working\s*-\s*",
    re.IGNORECASE,
)
FIX_TAG_RE = re.compile(
    r"^#\s*Fix\b",
    re.IGNORECASE,
)
REWORK_TAG_RE = re.compile(
    r"^#\s*Rework\b",
    re.IGNORECASE,
)
TAG_DUE_DATE_PREFIXES = ("Working", "Waiting", "Assignee")
TAG_DUE_DATE_PREFIX_RE = re.compile(
    r"^#\s*(" + "|".join(TAG_DUE_DATE_PREFIXES) + r")\s*-\s*",
    re.IGNORECASE,
)
TAG_DUE_DATE_STRUCTURE_RE = re.compile(
    r"^#\s*(?P<tag>" + "|".join(TAG_DUE_DATE_PREFIXES) + r")\s*-\s*"
    r"(?P<data1>\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2})\s*-\s*"
    r"(?:(?P<data2>\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2})\s*-\s*"
    r"|-\s*)"
    r"(?P<commento>.*)$",
    re.IGNORECASE,
)


def risolvi_percorso(nome_o_path):
    if os.path.isabs(nome_o_path):
        return nome_o_path
    return os.path.normpath(os.path.join(SCRIPT_DIR, nome_o_path))


def percorso_output_da_csv(input_path):
    base, _ = os.path.splitext(input_path)
    return base + ".xlsx"


def percorso_html_da_xlsx(xlsx_path):
    """Report HTML omonimo del file Excel prodotto (.xlsx → .html)."""
    base, _ = os.path.splitext(xlsx_path)
    return base + ".html"


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
    """Giornate lavorative tra la data nel tag '# Inizio Attivita' - <data>' e oggi."""
    if data_oggi is None:
        data_oggi = datetime.now().date()
    for tag in estrai_tag_temporali(description):
        if not is_tag_inizio_attivita(tag):
            continue
        date = estrai_date_da_tag(tag)
        if date:
            return giorni_lavorativi_tra(date[0], data_oggi)
    return None


def is_tag_waiting(tag):
    if not tag:
        return False
    return bool(WAITING_TAG_RE.match(str(tag).strip()))


def is_tag_working(tag):
    if not tag:
        return False
    return bool(WORKING_TAG_RE.match(str(tag).strip()))


def is_tag_fix(tag):
    if not tag:
        return False
    return bool(FIX_TAG_RE.match(str(tag).strip()))


def is_tag_rework(tag):
    if not tag:
        return False
    return bool(REWORK_TAG_RE.match(str(tag).strip()))


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


def richiede_struttura_due_date(tag):
    """True per tag con struttura '# <nome> - <data> - <data> -' (es. Working, Waiting)."""
    if not tag:
        return False
    return bool(TAG_DUE_DATE_PREFIX_RE.match(str(tag).strip()))


def parse_tag_due_date(tag):
    """
    Analizza '# <tag> - <data1> - <data2> - <commento>'.
    data2 puo' mancare (attivita' in corso, indicata come '- -').
    Restituisce dict con chiavi data1, data2 (opzionale), commento, oppure None.
    """
    if not tag:
        return None
    match = TAG_DUE_DATE_STRUCTURE_RE.match(str(tag).strip())
    if not match:
        return None
    data1 = parse_data(match.group("data1"))
    if data1 is None:
        return None
    data2_raw = match.group("data2")
    data2 = parse_data(data2_raw) if data2_raw else None
    return {
        "data1": data1,
        "data2": data2,
        "commento": (match.group("commento") or "").strip(),
    }


def estrai_date_campi_due_date(tag):
    """Estrae le date dai campi '# <tag> - <data1> - <data2> -' (data2 opzionale)."""
    parsed = parse_tag_due_date(tag)
    if not parsed:
        return []
    date = [parsed["data1"]]
    if parsed["data2"] is not None:
        date.append(parsed["data2"])
    return date


def valida_struttura_tag_due_date(tag, title=None):
    """
    Verifica la struttura '# <tag> - <data> - <data> -'.
    La seconda data puo' mancare (attivita' in corso). Stampa Warning se non conforme.
    """
    if not richiede_struttura_due_date(tag):
        return True

    contesto = f" [card: {title}]" if title else ""
    testo = str(tag).strip()
    parsed = parse_tag_due_date(tag)

    if parsed is None:
        print(
            "Warning: struttura tag non conforme "
            f"(atteso '# <tag> - <data> - <data> -'): {testo!r}{contesto}"
        )
        return False

    return True


def is_tag_done(tag):
    if not tag:
        return False
    return bool(re.search(r"\bdone\b", str(tag), re.IGNORECASE))


def tag_contiene_in_progress(tag):
    if not tag:
        return False
    return bool(re.search(r"\bin progress\b", str(tag), re.IGNORECASE))


def giorni_lavorativi_tra(inizio, fine, fine_inclusa=True):
    """
    Giornate lavorative italiane (lun-ven).
    fine_inclusa=True: periodo [inizio, fine] inclusi (default).
    fine_inclusa=False: periodo [inizio, fine) con data finale esclusa.
    """
    if fine_inclusa:
        if inizio > fine:
            return 0
    elif inizio >= fine:
        return 0

    giorni = 0
    corrente = inizio
    while (corrente <= fine) if fine_inclusa else (corrente < fine):
        if corrente.weekday() < 5:
            giorni += 1
        corrente += timedelta(days=1)
    return giorni


def giorni_da_tag_working(tag, data_oggi=None):
    """
    Giornate lavorative nel tag '# Working - <start> - <end>' (solo lun-ven).
    Con due date: [start, end) — inizio incluso, fine esclusa.
    Con sola start date: giornate lavorative da start a oggi incluso (in corso).
    """
    if data_oggi is None:
        data_oggi = datetime.now().date()
    if not is_tag_working(tag):
        return None
    date = estrai_date_campi_due_date(tag)
    if len(date) >= 2:
        return giorni_lavorativi_tra(date[0], date[1], fine_inclusa=False)
    if date:
        return giorni_lavorativi_tra(date[0], data_oggi)
    return None


def giorni_da_tag_temporale(tag, tag_successivo=None, data_oggi=None):
    """
    Giornate lavorative (lun-ven) nel TAG Temporale.
    Tag a due date (Working, Waiting, …): [inizio, fine) con fine esclusa;
    senza fine esplicita usa oggi incluso.
    Tag generici: stessa logica; con una data e tag successivo la data del
    successivo e' fine esclusa. I tag Done hanno solo la data di chiusura.
    """
    if data_oggi is None:
        data_oggi = datetime.now().date()

    if richiede_struttura_due_date(tag):
        date = estrai_date_campi_due_date(tag)
        if len(date) >= 2:
            return giorni_lavorativi_tra(date[0], date[1], fine_inclusa=False)
        if date:
            return giorni_lavorativi_tra(date[0], data_oggi)
        return None

    date = estrai_date_da_tag(tag)
    if len(date) >= 2:
        return giorni_lavorativi_tra(date[0], date[1], fine_inclusa=False)
    if tag_successivo:
        date_succ = estrai_date_da_tag(tag_successivo)
        if date and date_succ:
            return giorni_lavorativi_tra(
                date[0], date_succ[0], fine_inclusa=False
            )
    if date and not is_tag_done(tag):
        return giorni_lavorativi_tra(date[0], data_oggi)
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
        estimate_tag = estrai_estimate_da_description(record.get("Description", ""))
        estimate_csv = parse_numero(record.get("Estimate", ""))
        if estimate_tag is not None:
            nuova_base["Estimate"] = estimate_tag
        elif estimate_csv is not None:
            nuova_base["Estimate"] = estimate_csv
        else:
            nuova_base["Estimate"] = ""
        nuova_base[INIZIO_LAVORAZIONE_COL] = giorni_da_inizio_attivita(
            record.get("Description", "")
        )

        tag_list = [
            tag
            for tag in estrai_tag_temporali(record.get("Description", ""))
            if not is_tag_estimate(tag)
        ]
        if not tag_list:
            nuova = dict(nuova_base)
            nuova[TOTALE_WAITING_COL] = None
            nuova[TOTALE_LAVORAZIONE_COL] = None
            nuova[REWORK_TIME_COL] = None
            nuova[FIX_TIME_COL] = None
            nuova[GIORNI_COL] = None
            nuova[TAG_TEMPORALI_COL] = ""
            righe.append(nuova)
            continue
        for i, tag in enumerate(tag_list):
            valida_struttura_tag_due_date(tag, title=record.get("Title"))
            tag_next = tag_list[i + 1] if i + 1 < len(tag_list) else None
            nuova = dict(nuova_base)
            nuova[TOTALE_WAITING_COL] = None
            nuova[TOTALE_LAVORAZIONE_COL] = None
            nuova[REWORK_TIME_COL] = None
            nuova[FIX_TIME_COL] = None
            nuova[GIORNI_COL] = giorni_da_tag_temporale(tag, tag_next)
            nuova[TAG_TEMPORALI_COL] = tag
            righe.append(nuova)
    return righe


def colonne_output():
    return KANBAN_COLUMNS + [
        INIZIO_LAVORAZIONE_COL,
        TOTALE_WAITING_COL,
        TOTALE_LAVORAZIONE_COL,
        REWORK_TIME_COL,
        FIX_TIME_COL,
        GIORNI_COL,
        TAG_TEMPORALI_COL,
    ]


def somma_giorni_tag_gruppo(ws, start, end, matcher, data_oggi=None):
    """Somma giornate lavorative per righe tag (col. P) che passano matcher."""
    if data_oggi is None:
        data_oggi = datetime.now().date()
    totale = 0.0
    ha_valori = False
    for row in range(start, end + 1):
        tag = ws.cell(row=row, column=COL_TAG).value
        if not matcher(tag):
            continue
        tag_next = None
        if row + 1 <= end:
            tag_next = ws.cell(row=row + 1, column=COL_TAG).value
        giorni = giorni_da_tag_temporale(tag, tag_next, data_oggi=data_oggi)
        if giorni is not None:
            totale += giorni
            ha_valori = True
    return totale if ha_valori else None


def somma_colonna_giorni_filtrata(ws, start, end, matcher):
    """Somma colonna O (Giorni) per righe il cui tag (col. P) passa matcher."""
    totale = 0.0
    ha_valori = False
    for row in range(start, end + 1):
        tag = ws.cell(row=row, column=COL_TAG).value
        if not matcher(tag):
            continue
        val = parse_numero(ws.cell(row=row, column=COL_GIORNI).value)
        if val is not None:
            totale += val
            ha_valori = True
    return totale if ha_valori else None


def percentuale_su_estimate(estimate, totale):
    """Percentuale (totale / estimate) × 100 per confronto col. M vs Estimate (G)."""
    if estimate is None or totale is None or estimate <= 0:
        return None
    return (totale / estimate) * 100


def applica_colore_colonna_m(cella, estimate, totale_lavorazione):
    """
    Sfondo col. M (Totale Lavorazione) in base a (M / Estimate) × 100.
    Estimate in col. G. Fasce: ≤50% verde; 51–80% giallo; 81–100% rosso pastello; >100% rosso acceso.
    """
    percentuale = percentuale_su_estimate(estimate, totale_lavorazione)
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


def applica_totali_gruppo(ws, start, end):
    center = Alignment(horizontal="center", vertical="center")

    tot_waiting = somma_giorni_tag_gruppo(ws, start, end, is_tag_waiting)
    cella_wait = ws.cell(row=start, column=COL_TOTALE_WAITING)
    cella_wait.value = tot_waiting
    cella_wait.alignment = center

    tot_lavorazione = somma_giorni_tag_gruppo(ws, start, end, is_tag_working)
    cella_lav = ws.cell(row=start, column=COL_TOTALE_LAVORAZIONE)
    cella_lav.value = tot_lavorazione
    cella_lav.alignment = center

    tot_rework = somma_colonna_giorni_filtrata(ws, start, end, is_tag_rework)
    cella_rework = ws.cell(row=start, column=COL_REWORK_TIME)
    cella_rework.value = tot_rework
    cella_rework.alignment = center

    tot_fix = somma_giorni_tag_gruppo(ws, start, end, is_tag_fix)
    cella_fix = ws.cell(row=start, column=COL_FIX_TIME)
    cella_fix.value = tot_fix
    cella_fix.alignment = center

    estimate = parse_numero(ws.cell(row=start, column=COL_ESTIMATE).value)
    applica_colore_colonna_m(cella_lav, estimate, tot_lavorazione)

    cella_inizio = ws.cell(row=start, column=COL_INIZIO_LAVORAZIONE)
    cella_inizio.alignment = center


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
                    ws.cell(row=start, column=COL_TOTALE_LAVORAZIONE).value
                ),
            }
        )
    return riepilogo


def status_escluso_da_export(status):
    return normalizza_testo(status, compatta_spazi=True) in STATUS_ESCLUSI_EXPORT


def filtra_card_per_export(card):
    return [
        record
        for record in card
        if not status_escluso_da_export(record.get("Status", ""))
    ]


def righe_export(card):
    """Una riga per card: Title, Status, Start Date, End Date, Tags (col. E)."""
    return [
        {col: record.get(col, "") for col in COLONNE_EXPORT}
        for record in filtra_card_per_export(card)
    ]


def fill_export_per_status(status):
    chiave = normalizza_testo(status, compatta_spazi=True).lower()
    return EXPORT_STATUS_FILLS.get(chiave)


def applica_bordo_riga_export(ws, row, min_col, max_col):
    """Bordo blu perimetrale alla riga e divisori blu chiaro tra le colonne."""
    for col in range(min_col, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.border = Border(
            left=EXPORT_BLUE_BORDER if col == min_col else None,
            right=EXPORT_COL_DIVIDER if col < max_col else EXPORT_BLUE_BORDER,
            top=EXPORT_BLUE_BORDER,
            bottom=EXPORT_BLUE_BORDER,
        )


def formatta_foglio_export(ws):
    """Colora A–E per Status (col. B), centra celle, bordi riga e colonne."""
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    max_col = min(ws.max_column, EXPORT_COL_LAST)
    for row in range(1, ws.max_row + 1):
        fill = None
        if row >= 2:
            fill = fill_export_per_status(ws.cell(row=row, column=2).value)
        applica_bordo_riga_export(ws, row, 1, max_col)
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = center
            if fill is not None:
                cell.fill = fill


def aggiungi_footer_data(ws, gruppi):
    center = Alignment(horizontal="center", vertical="center")
    data_last = ws.max_row
    totals_row = data_last + 2
    ts_row = totals_row + 1
    legend_start = ts_row + 2

    tot_estimate = 0.0
    tot_waiting = 0.0
    tot_lavorazione = 0.0
    ha_estimate = False
    ha_waiting = False
    ha_lavorazione = False
    for start, _end in gruppi:
        estimate = parse_numero(ws.cell(row=start, column=COL_ESTIMATE).value)
        waiting = parse_numero(ws.cell(row=start, column=COL_TOTALE_WAITING).value)
        lavorazione = parse_numero(
            ws.cell(row=start, column=COL_TOTALE_LAVORAZIONE).value
        )
        if estimate is not None:
            tot_estimate += estimate
            ha_estimate = True
        if waiting is not None:
            tot_waiting += waiting
            ha_waiting = True
        if lavorazione is not None:
            tot_lavorazione += lavorazione
            ha_lavorazione = True

    ws.cell(row=totals_row, column=1).value = len(gruppi)
    ws.cell(row=totals_row, column=1).alignment = center
    if ha_estimate:
        ws.cell(row=totals_row, column=COL_ESTIMATE).value = tot_estimate
        ws.cell(row=totals_row, column=COL_ESTIMATE).alignment = center
    if ha_waiting:
        ws.cell(row=totals_row, column=COL_TOTALE_WAITING).value = tot_waiting
        ws.cell(row=totals_row, column=COL_TOTALE_WAITING).alignment = center
    if ha_lavorazione:
        ws.cell(row=totals_row, column=COL_TOTALE_LAVORAZIONE).value = tot_lavorazione
        ws.cell(row=totals_row, column=COL_TOTALE_LAVORAZIONE).alignment = center

    n_card_waiting = conteggio_card_totale_waiting(ws, gruppi)
    ws.cell(row=ts_row, column=1).value = datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )
    aggiungi_legenda_colonne(ws, legend_start, n_card_totale_waiting=n_card_waiting)


def aggiungi_legenda_colonne(ws, start_row, n_card_totale_waiting=None):
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
        if (
            titolo == TOTALE_WAITING_COL
            and n_card_totale_waiting is not None
        ):
            cell_titolo.value = f"{titolo} ({n_card_totale_waiting} card)"
        else:
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


def formatta_foglio_dati(ws):
    ws.cell(row=1, column=COL_INIZIO_LAVORAZIONE).value = INIZIO_LAVORAZIONE_COL
    ws.cell(row=1, column=COL_TOTALE_WAITING).value = TOTALE_WAITING_COL
    ws.cell(row=1, column=COL_TOTALE_LAVORAZIONE).value = TOTALE_LAVORAZIONE_COL
    ws.cell(row=1, column=COL_REWORK_TIME).value = REWORK_TIME_COL
    ws.cell(row=1, column=COL_FIX_TIME).value = FIX_TIME_COL
    gruppi = formatta_foglio_card(ws)
    aggiungi_footer_data(ws, gruppi)
    return gruppi


def scrivi_excel(card, output_path):
    righe_all = espandi_card_con_tag(card)
    df_all = pd.DataFrame(righe_all, columns=colonne_output())
    df_export = pd.DataFrame(righe_export(card), columns=COLONNE_EXPORT)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
        df_export.to_excel(writer, sheet_name=DATA_EXPORT_SHEET, index=False)

    wb = load_workbook(output_path)
    gruppi = formatta_foglio_dati(wb[OUTPUT_SHEET])
    formatta_foglio_export(wb[DATA_EXPORT_SHEET])
    riepilogo = riepilogo_da_gruppi(wb[OUTPUT_SHEET], gruppi)
    crea_foglio_stat(wb, riepilogo)
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


def is_attesa_test_o_fab(status):
    """True se lo Status MIRO indica attesa test/fabbrica (colonna waiting for fab.)."""
    testo = normalizza_testo(status, compatta_spazi=True).lower()
    return bool(
        re.search(r"waiting\s*(for\s*fab\.?|test)|attesa\s*(fab\.?|test)", testo)
    )


def classifica_colonna_kanban(status):
    """Mappa il campo Status MIRO a una colonna Kanban dello snapshot."""
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
    if is_attesa_test_o_fab(status):
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
        colonna = classifica_colonna_kanban(record.get("Status", ""))
        if colonna is None:
            non_mappate += 1
            continue
        conteggi[colonna] += 1
    return conteggi, non_mappate


def is_status_new_to_be_verified(status):
    return (
        normalizza_testo(status, compatta_spazi=True).lower()
        == "new - to be verified"
    )


def conteggio_card_new_to_be_verified(card):
    return sum(
        1 for record in card
        if is_status_new_to_be_verified(record.get("Status", ""))
    )


def riepilogo_acronimi_scope(card):
    """Target, concordati (scope Kanban), in verifica, gap e liste acronimi per verifica."""
    conteggi, _ = conteggio_per_colonna_kanban(card)
    concordati_list = []
    in_verifica_list = []
    non_mappati_list = []
    tutti_board = []
    per_colonna = {col: [] for col in KANBAN_SNAPSHOT_COLS}
    status_per_colonna = {col: Counter() for col in KANBAN_SNAPSHOT_COLS}

    for record in card:
        title = normalizza_testo(record.get("Title", "")) or "(senza title)"
        tutti_board.append(title)
        status = record.get("Status", "")
        status_label = normalizza_testo(status) or "(vuoto)"
        if is_status_new_to_be_verified(status):
            in_verifica_list.append(title)
        colonna = classifica_colonna_kanban(status)
        if colonna in KANBAN_SNAPSHOT_COLS:
            concordati_list.append(title)
            per_colonna[colonna].append(title)
            status_per_colonna[colonna][status_label] += 1
        elif colonna is None:
            non_mappati_list.append(title)

    concordati = int(sum(conteggi[col] for col in KANBAN_SNAPSHOT_COLS))
    in_verifica = len(set(in_verifica_list))
    target = SCOPE_TOTALE_CARD
    return {
        "target": target,
        "concordati": concordati,
        "in_verifica": in_verifica,
        "da_aggiungere": target - concordati,
        "acronimi_tutti_board": sorted(set(tutti_board)),
        "acronimi_concordati": sorted(set(concordati_list)),
        "acronimi_in_verifica": sorted(set(in_verifica_list)),
        "acronimi_non_mappati": sorted(set(non_mappati_list)),
        "acronimi_per_colonna": {
            col: sorted(set(per_colonna[col])) for col in KANBAN_SNAPSHOT_COLS
        },
        "status_per_colonna": {
            col: dict(status_per_colonna[col]) for col in KANBAN_SNAPSHOT_COLS
        },
    }


def conteggio_card_totale_waiting(ws, gruppi):
    """Card con Totale Waiting (col. L) valorizzato."""
    return sum(
        1
        for start, _end in gruppi
        if parse_numero(ws.cell(row=start, column=COL_TOTALE_WAITING).value)
        is not None
    )


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
    not_started = (
        df["Under analysis"].astype(int) + df["backlog"].astype(int)
    ).tolist()
    in_delivery = _serie_wip(df)
    done = df["Acronimi done"].astype(int).tolist()
    return {
        "dates": dates,
        "weeks": weeks,
        "column_labels": KANBAN_DISPLAY_LABELS,
        "colors": KANBAN_CHART_COLORS,
        "burnup": {
            "done": done,
            "scope": _serie_totali(df),
            "target": [SCOPE_TOTALE_CARD] * n,
        },
        "wip": in_delivery,
        "velocity": _serie_velocita(df),
        "stacked": {
            col: df[col].astype(int).tolist() for col in KANBAN_SNAPSHOT_COLS
        },
        "columns": {
            col: df[col].astype(int).tolist() for col in KANBAN_SNAPSHOT_COLS
        },
        "scrum": {
            "not_started": not_started,
            "in_delivery": in_delivery,
            "done": done,
        },
    }


def _html_griglia_colonne_kanban():
    """Sezione HTML: un mini-grafico per ogni colonna Kanban."""
    cards = []
    for i, col in enumerate(KANBAN_SNAPSHOT_COLS):
        label = KANBAN_DISPLAY_LABELS[col]
        cards.append(
            f'    <div class="chart-card">\n'
            f'      <h3>{label}</h3>\n'
            f'      <div class="chart-wrap-sm"><canvas id="chart-col-{i}"></canvas></div>\n'
            f'    </div>'
        )
    return "\n".join(cards)


def _nome_acronimo_da_title(title):
    """Prima riga del Title MIRO (nome acronimo)."""
    if not title:
        return "(senza title)"
    prima = title.splitlines()[0].strip()
    return prima or "(senza title)"


def _html_acronimi_virgola(acronimi):
    """Nomi acronimo (prima riga Title) separati da virgola."""
    if not acronimi:
        return "<em>nessuno</em>"
    nomi = [_nome_acronimo_da_title(t) for t in acronimi]
    return html.escape(", ".join(nomi))


def _html_status_breakdown(status_counter):
    """Elenco Status con conteggio card, ordinato per numero decrescente."""
    if not status_counter:
        return "<em>nessuno</em>"
    items = sorted(status_counter.items(), key=lambda x: (-x[1], x[0]))
    parti = [f"{html.escape(st)} ({n})" for st, n in items]
    return ", ".join(parti)


def _merge_status_per_colonne(cols, status_per_colonna):
    merged = Counter()
    for col in cols:
        merged.update(status_per_colonna.get(col, {}))
    return dict(merged)


def _merge_acronimi_per_colonne(cols, acronimi_per_colonna):
    tutti = []
    for col in cols:
        tutti.extend(acronimi_per_colonna.get(col, []))
    return sorted(set(tutti))


def _html_kpi_box_dettaglio(status_counter, acronimi, extra_lines=None):
    """Dettaglio KPI: Status con conteggio e lista acronimi."""
    lines = []
    if status_counter:
        n_status = sum(status_counter.values())
        lines.append(
            f"Status ({n_status}): {_html_status_breakdown(status_counter)}"
        )
    if acronimi is not None:
        lines.append(
            f"Acronimi ({len(acronimi)}): {_html_acronimi_virgola(acronimi)}"
        )
    if extra_lines:
        lines.extend(extra_lines)
    return "<br>".join(lines) if lines else ""


def _html_kpi_top_row_dettaglio(
    riepilogo, df, tot_ultimo, done_ultimo, wip_ultimo, not_started_ultimo
):
    """Seconda riga di 6 KPI sotto il riassunto: Status + acronimi per ciascun riquadro."""
    if not riepilogo:
        return ""
    status_pc = riepilogo.get("status_per_colonna", {})
    acronimi_pc = riepilogo.get("acronimi_per_colonna", {})
    all_cols = KANBAN_SNAPSHOT_COLS
    not_started_cols = ["Under analysis", "backlog"]

    snapshot_dates = [
        row["data"].strftime("%d/%m/%Y") for _, row in df.iterrows()
    ]
    snapshot_extra = [
        f"Date ({len(snapshot_dates)}): "
        f"{html.escape(', '.join(snapshot_dates)) if snapshot_dates else 'nessuna'}"
    ]

    boxes = [
        (
            "Scope tracciato",
            tot_ultimo,
            _html_kpi_box_dettaglio(
                _merge_status_per_colonne(all_cols, status_pc),
                _merge_acronimi_per_colonne(all_cols, acronimi_pc),
            ),
        ),
        (
            "Target card",
            SCOPE_TOTALE_CARD,
            _html_kpi_box_dettaglio(
                None,
                None,
                extra_lines=["Target di progetto: 30 card"],
            ),
        ),
        (
            "Non avviato (Under analysis + Backlog)",
            not_started_ultimo,
            _html_kpi_box_dettaglio(
                _merge_status_per_colonne(not_started_cols, status_pc),
                _merge_acronimi_per_colonne(not_started_cols, acronimi_pc),
            ),
        ),
        (
            "In delivery (WIP)",
            wip_ultimo,
            _html_kpi_box_dettaglio(
                _merge_status_per_colonne(WIP_SNAPSHOT_COLS, status_pc),
                _merge_acronimi_per_colonne(WIP_SNAPSHOT_COLS, acronimi_pc),
            ),
        ),
        (
            "Acronimi Done",
            done_ultimo,
            _html_kpi_box_dettaglio(
                status_pc.get("Acronimi done", {}),
                acronimi_pc.get("Acronimi done", []),
            ),
        ),
        (
            "Snapshot",
            len(df),
            _html_kpi_box_dettaglio(None, None, extra_lines=snapshot_extra),
        ),
    ]

    return "\n".join(
        f'    <div class="kpi"><b>{val}</b><span>'
        f'{html.escape(label)}'
        + (f'<br><small>{dettaglio}</small>' if dettaglio else "")
        + "</span></div>"
        for label, val, dettaglio in boxes
    )


def _html_dettaglio_per_colonna(per_colonna):
    """Breakdown acronimi concordati per colonna Kanban."""
    parti = []
    for col in KANBAN_SNAPSHOT_COLS:
        acronimi = per_colonna.get(col, [])
        label = html.escape(KANBAN_DISPLAY_LABELS[col])
        n = len(acronimi)
        lista = _html_acronimi_virgola(acronimi)
        parti.append(f"<strong>{label} ({n})</strong> {lista}")
    return "<br>".join(parti)


def _testo_colonne_kanban_concordati():
    """Etichette display delle colonne Kanban sommate in Acronimi concordati."""
    return ", ".join(KANBAN_DISPLAY_LABELS[col] for col in KANBAN_SNAPSHOT_COLS)


def _html_kpi_acronimi(riepilogo):
    """Scatole KPI HTML: target, concordati, in verifica."""
    if not riepilogo:
        return ""
    colonne = html.escape(_testo_colonne_kanban_concordati())
    concordati = riepilogo.get("acronimi_concordati", [])
    in_verifica = riepilogo.get("acronimi_in_verifica", [])
    per_colonna = riepilogo.get("acronimi_per_colonna", {})

    dettagli = [
        ("Acronimi target", riepilogo["target"], ""),
        (
            "Acronimi concordati",
            riepilogo["concordati"],
            (
                f"Somma colonne Kanban: {colonne}<br>"
                f"Acronimi ({len(concordati)}): "
                f"{_html_dettaglio_per_colonna(per_colonna)}"
            ),
        ),
        (
            "Acronimi in verifica",
            riepilogo["in_verifica"],
            (
                "Status: new - to be verified<br>"
                f"Acronimi ({len(in_verifica)}): "
                f"{_html_acronimi_virgola(in_verifica)}"
            ),
        ),
    ]
    return "\n".join(
        f'    <div class="kpi"><b>{val}</b><span>'
        f'{html.escape(label)}'
        + (f'<br><small>{dettaglio}</small>' if dettaglio else "")
        + "</span></div>"
        for label, val, dettaglio in dettagli
    )


def _js_griglia_colonne_kanban():
    """JavaScript: grafici linea per singola colonna Kanban."""
    blocks = []
    for i, col in enumerate(KANBAN_SNAPSHOT_COLS):
        label = json.dumps(KANBAN_DISPLAY_LABELS[col], ensure_ascii=False)
        color = json.dumps(KANBAN_CHART_COLORS[col])
        fill = json.dumps(KANBAN_CHART_FILL[col])
        key = json.dumps(col, ensure_ascii=False)
        blocks.append(f"""
  new Chart(document.getElementById("chart-col-{i}"), {{
    type: "line",
    data: {{
      labels: axisLabels,
      datasets: [{{
        label: {label},
        data: D.columns[{key}],
        borderColor: {color},
        backgroundColor: {fill},
        fill: true, tension: 0.3, pointRadius: 3
      }}]
    }},
    options: baseOptsSm
  }});""")
    return "".join(blocks)


def genera_html_jkan(df, html_path, data_ultimo_snapshot=None, riepilogo_acronimi=None):
    """Genera report HTML Scrum/Kanban con burnup, CFD ed evoluzione colonne."""
    if df.empty:
        return

    dati = _dati_grafici_jkan(df)
    chart_json = json.dumps(dati, ensure_ascii=False)
    griglia_html = _html_griglia_colonne_kanban()
    js_colonne = _js_griglia_colonne_kanban()
    kpi_acronimi_html = _html_kpi_acronimi(riepilogo_acronimi)
    thead_cols = "".join(
        f"<th>{KANBAN_DISPLAY_LABELS[col]}</th>" for col in KANBAN_SNAPSHOT_COLS
    )
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    ultima = (
        data_ultimo_snapshot.strftime("%d/%m/%Y")
        if data_ultimo_snapshot
        else df["data"].iloc[-1].strftime("%d/%m/%Y")
    )
    tot_ultimo = int(somma_colonne_scope(df).iloc[-1])
    done_ultimo = int(df["Acronimi done"].iloc[-1])
    wip_ultimo = int(df[WIP_SNAPSHOT_COLS].iloc[-1].sum())
    not_started_ultimo = int(
        df["Under analysis"].iloc[-1] + df["backlog"].iloc[-1]
    )
    kpi_top_dettaglio_html = _html_kpi_top_row_dettaglio(
        riepilogo_acronimi,
        df,
        tot_ultimo,
        done_ultimo,
        wip_ultimo,
        not_started_ultimo,
    )

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
  <title>JKAN — Report Scrum Kanban</title>
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
    main {{ max-width: 1200px; margin: 0 auto; padding: 1.5rem 1rem 3rem; }}
    h1 {{ margin: 0 0 .25rem; font-size: 1.6rem; }}
    .sub {{ color: var(--muted); margin: 0 0 1rem; font-size: .92rem; }}
    .kpis {{
      display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
      gap: .75rem; margin-bottom: 1.5rem;
    }}
    .kpi {{
      background: var(--card); border: 1px solid var(--border);
      border-radius: 8px; padding: .85rem 1rem;
    }}
    .kpi b {{ display: block; font-size: 1.5rem; color: var(--accent); }}
    .kpi span {{ font-size: .85rem; color: var(--muted); }}
    .kpi span small {{ display: block; margin-top: .35rem; font-size: .78rem; line-height: 1.35; color: var(--muted); }}
    .kpis-acronimi {{
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    }}
    .kpis-acronimi .kpi span small {{
      max-height: 14rem; overflow-y: auto;
    }}
    section {{
      background: var(--card); border: 1px solid var(--border);
      border-radius: 8px; padding: 1rem 1.25rem; margin-bottom: 1.25rem;
    }}
    section h2 {{ margin: 0 0 .75rem; font-size: 1.1rem; }}
    .chart-wrap {{ position: relative; height: 340px; }}
    .chart-wrap-lg {{ position: relative; height: 400px; }}
    .chart-grid {{
      display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
      gap: 1rem;
    }}
    .chart-card {{
      border: 1px solid var(--border); border-radius: 6px; padding: .75rem;
      background: #fafbfc;
    }}
    .chart-card h3 {{ margin: 0 0 .5rem; font-size: .95rem; font-weight: 600; }}
    .chart-wrap-sm {{ position: relative; height: 200px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: .9rem; }}
    th, td {{ border: 1px solid var(--border); padding: .45rem .6rem; text-align: center; }}
    th {{ background: #f1f5f9; }}
    td:first-child, th:first-child {{ text-align: left; }}
  </style>
</head>
<body>
<main>
  <h1>JKAN — Report Scrum Kanban</h1>
  <p class="sub">Generato il {now_str} · ultimo snapshot: {ultima}</p>

  <div class="kpis">
    <div class="kpi"><b>{tot_ultimo}</b><span>Scope tracciato</span></div>
    <div class="kpi"><b>{SCOPE_TOTALE_CARD}</b><span>Target card</span></div>
    <div class="kpi"><b>{not_started_ultimo}</b><span>Non avviato (Under analysis + Backlog)</span></div>
    <div class="kpi"><b>{wip_ultimo}</b><span>In delivery (WIP)</span></div>
    <div class="kpi"><b>{done_ultimo}</b><span>Acronimi Done</span></div>
    <div class="kpi"><b>{len(df)}</b><span>Snapshot</span></div>
  </div>

  <div class="kpis kpis-acronimi">
{kpi_top_dettaglio_html}
  </div>

  <section id="acronimi-scope">
    <h2>Riepilogo acronimi</h2>
    <p class="sub">Snapshot corrente: target di progetto, scope concordato sul board Kanban, card in verifica e gap residuo. Generato il {now_str}.</p>
    <div class="kpis kpis-acronimi">
{kpi_acronimi_html}
    </div>
  </section>

  <section id="cfd">
    <h2>Cumulative Flow Diagram (CFD)</h2>
    <p class="sub">Evoluzione cumulativa delle colonne Kanban nel tempo — visualizzazione tipica Scrum/Agile per individuare colli di bottiglia e squilibri nel flusso.</p>
    <div class="chart-wrap-lg"><canvas id="chart-cfd"></canvas></div>
  </section>

  <section id="columns-all">
    <h2>Evoluzione colonne Kanban — vista comparata</h2>
    <p class="sub">Andamento di tutte le colonne nello stesso grafico.</p>
    <div class="chart-wrap-lg"><canvas id="chart-columns-all"></canvas></div>
  </section>

  <section id="columns-grid">
    <h2>Evoluzione per colonna Kanban</h2>
    <p class="sub">Dettaglio snapshot-by-snapshot di ciascuna colonna del board.</p>
    <div class="chart-grid">
{griglia_html}
    </div>
  </section>

  <section id="scrum-pipeline">
    <h2>Pipeline Scrum — non avviato / in delivery / completato</h2>
    <p class="sub">Aggregazione per fase: analisi e backlog (upstream), lavorazione e test (delivery), acronimi completati.</p>
    <div class="chart-wrap"><canvas id="chart-scrum"></canvas></div>
  </section>

  <section id="burnup">
    <h2>Burnup — Acronimi Done vs scope</h2>
    <p class="sub">Scope tracciato = somma di tutte le colonne. Linea tratteggiata: target {SCOPE_TOTALE_CARD} card.</p>
    <div class="chart-wrap"><canvas id="chart-burnup"></canvas></div>
  </section>

  <section id="stacked">
    <h2>Distribuzione card per colonna (barre impilate)</h2>
    <div class="chart-wrap"><canvas id="chart-stacked"></canvas></div>
  </section>

  <section id="wip">
    <h2>WIP — lavoro in corso</h2>
    <p class="sub">In progress + Waiting for fabric + Fab. Test in progress.</p>
    <div class="chart-wrap"><canvas id="chart-wip"></canvas></div>
  </section>

  <section id="velocity">
    <h2>Velocità — throughput per snapshot</h2>
    <p class="sub">Incremento di Acronimi Done rispetto allo snapshot precedente.</p>
    <div class="chart-wrap"><canvas id="chart-velocity"></canvas></div>
  </section>

  <section id="tabella">
    <h2>Storico dbJKAN.csv</h2>
    <table>
      <thead>
        <tr>
          <th>Data</th>
          {thead_cols}
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
  const COLS = {json.dumps(KANBAN_SNAPSHOT_COLS, ensure_ascii=False)};
  const axisLabels = D.dates.map(function(d, i) {{ return [d, D.weeks[i]]; }});
  const baseOpts = {{
    responsive: true,
    maintainAspectRatio: false,
    plugins: {{ legend: {{ position: "top" }} }},
    scales: {{
      x: {{ ticks: {{ maxRotation: 0, autoSkip: true }} }},
      y: {{ beginAtZero: true, ticks: {{ precision: 0 }} }}
    }}
  }};
  const baseOptsSm = Object.assign({{}}, baseOpts, {{
    plugins: {{ legend: {{ display: false }} }}
  }});

  function colDataset(col, stackId) {{
    return {{
      label: D.column_labels[col] || col,
      data: D.stacked[col],
      borderColor: D.colors[col],
      backgroundColor: D.colors[col],
      stack: stackId,
      fill: true,
      tension: 0.25,
      pointRadius: 2
    }};
  }}

  new Chart(document.getElementById("chart-cfd"), {{
    type: "line",
    data: {{
      labels: axisLabels,
      datasets: COLS.map(function(col) {{ return colDataset(col, "cfd"); }})
    }},
    options: Object.assign({{}}, baseOpts, {{
      scales: {{
        x: {{ stacked: true, ticks: {{ maxRotation: 0, autoSkip: true }} }},
        y: {{ stacked: true, beginAtZero: true, ticks: {{ precision: 0 }} }}
      }},
      plugins: {{ legend: {{ position: "bottom" }} }}
    }})
  }});

  new Chart(document.getElementById("chart-columns-all"), {{
    type: "line",
    data: {{
      labels: axisLabels,
      datasets: COLS.map(function(col) {{
        return {{
          label: D.column_labels[col] || col,
          data: D.columns[col],
          borderColor: D.colors[col],
          backgroundColor: "transparent",
          fill: false,
          tension: 0.3,
          pointRadius: 3
        }};
      }})
    }},
    options: Object.assign({{}}, baseOpts, {{
      plugins: {{ legend: {{ position: "bottom" }} }}
    }})
  }});

  new Chart(document.getElementById("chart-scrum"), {{
    type: "line",
    data: {{
      labels: axisLabels,
      datasets: [
        {{
          label: "Non avviato (Under analysis + Backlog)",
          data: D.scrum.not_started,
          borderColor: "#6366f1",
          backgroundColor: "rgba(99,102,241,0.12)",
          fill: true, tension: 0.25, pointRadius: 3
        }},
        {{
          label: "In delivery (WIP)",
          data: D.scrum.in_delivery,
          borderColor: "#1a56db",
          backgroundColor: "rgba(26,86,219,0.12)",
          fill: true, tension: 0.25, pointRadius: 3
        }},
        {{
          label: "Acronimi Done",
          data: D.scrum.done,
          borderColor: "#10b981",
          backgroundColor: "rgba(16,185,129,0.12)",
          fill: true, tension: 0.25, pointRadius: 3
        }}
      ]
    }},
    options: baseOpts
  }});

  new Chart(document.getElementById("chart-burnup"), {{
    type: "line",
    data: {{
      labels: axisLabels,
      datasets: [
        {{
          label: "Acronimi Done",
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
      datasets: COLS.map(function(col) {{
        return {{
          label: D.column_labels[col] || col,
          data: D.stacked[col],
          backgroundColor: D.colors[col],
          stack: "kanban"
        }};
      }})
    }},
    options: Object.assign({{}}, baseOpts, {{
      scales: {{
        x: {{ stacked: true, ticks: {{ maxRotation: 0, autoSkip: true }} }},
        y: {{ stacked: true, beginAtZero: true, ticks: {{ precision: 0 }} }}
      }},
      plugins: {{ legend: {{ position: "bottom" }} }}
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
{js_colonne}
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
    html_path = percorso_html_da_xlsx(output_path)
    riepilogo_acronimi = riepilogo_acronimi_scope(card)
    genera_html_jkan(
        df_storico,
        html_path,
        data_ultimo_snapshot=data_snapshot,
        riepilogo_acronimi=riepilogo_acronimi,
    )

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
  <input>.html — report Scrum/Kanban omonimo del file Excel prodotto.
  dbJKAN.csv   — storico snapshot (cartella dello script).
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
