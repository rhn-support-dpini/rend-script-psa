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
    Le righe Description con prefisso "#" generano sotto-righe da colonna N;
    A–M sono merge verticali per Title, con bordo rosso pastello per card.
    Colonna J (InizioLavorazione(GG)), K (Waiting #), L (Totale Lavorazione),
    M (Period SUM), N (Giorni), O (TAG Temporali).
    Waiting #: giorni dall'ultimo tag "# Waiting -" a oggi, con sfondo giallo pastello.
    Period SUM: somma Giorni con tag in Progress; Totale Lavorazione: tag Lavorazione o in Progress.
    Colonna G (Estimate): valore numerico dal tag "# Estimate" in Description, non dal CSV.
    Giorni: se manca la 2ª data si usa oggi, eccetto tag Done (solo chiusura).
"""

import argparse
import csv
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
OUTPUT_SHEET = "data"
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
        "totale dei giorni in stato Lavorazione o in Progress",
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


def applica_colore_confronto_estimate(cella, estimate, valore):
    if estimate is None or valore is None:
        return
    if estimate >= valore:
        cella.fill = PASTEL_GREEN_FILL
    else:
        cella.fill = PASTEL_RED_FILL


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
    applica_colore_confronto_estimate(cella_lav, estimate, tot_lavorazione)
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


def scrivi_excel(card, output_path):
    righe = espandi_card_con_tag(card)
    df = pd.DataFrame(righe, columns=colonne_output())
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)

    wb = load_workbook(output_path)
    ws = wb[OUTPUT_SHEET]
    ws.cell(row=1, column=COL_TOTALE_LAVORAZIONE).value = TOTALE_LAVORAZIONE_COL
    ws.cell(row=1, column=COL_TAGS_SUM).value = PERIOD_SUM_HEADER
    gruppi = formatta_foglio_card(ws)
    riepilogo = riepilogo_da_gruppi(ws, gruppi)
    aggiungi_footer_data(ws, gruppi)
    crea_foglio_stat(wb, riepilogo)
    crea_foglio_graph(wb, riepilogo)
    wb.save(output_path)


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
    righe_output = len(espandi_card_con_tag(card))
    print(f"Sezioni kanban: {sezioni}")
    print(f"Card estratte: {len(card)}")
    print(f"Righe output: {righe_output}")
    print(f"Output: {output_path}")


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
