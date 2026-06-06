"""
elabora_JKAN.py — Estrae le card dalla Kanban MIRO "JBOSS-Barison"

Legge un export CSV di una board MIRO, individua la sezione Kanban con quel nome
e produce un file Excel con una riga per card (espansa per tag temporali in Description).

Utilizzo:
    python elabora_JKAN.py [input.csv]

    Default: input=2026-06-06-WIP.csv
    Output: stesso percorso e nome del CSV con estensione .xlsx
    Le righe Description con prefisso "#" generano righe replicate; colonne
    aggiuntive in coda: Giorni (Start Date → End Date) e TAG Temporali (testo tag).
"""

import csv
import os
import re
import sys
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

KANBAN_NAME = "JBOSS-Barison"
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
OUTPUT_SHEET = "data"
CENTER_COLS = {3, 4, 7}  # C=Status, D=Assignee, G=Estimate


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


def is_url_row(row):
    if not row:
        return False
    testo = normalizza_testo(row[0], compatta_spazi=True)
    if not testo:
        return True
    return testo.startswith("http://") or testo.startswith("https://")


def is_probabile_nome_kanban(row):
    if len(row) != 1:
        return False
    testo = normalizza_testo(row[0], compatta_spazi=True)
    if not testo or is_url_row(row):
        return False
    if len(testo) > 80:
        return False
    if testo.lower().startswith(("stat", "campi card", "descrizione", "kanban")):
        return False
    return True


def nome_kanban_da_contesto(rows, header_idx):
    """Ricostruisce il nome Kanban dalle righe che precedono l'header."""
    candidati = []
    for j in range(header_idx - 1, max(header_idx - 12, -1), -1):
        row = rows[j]
        if not row or all(not normalizza_testo(c) for c in row):
            continue
        if is_url_row(row):
            continue
        if len(row) == 1:
            testo = normalizza_testo(row[0], compatta_spazi=True)
            if is_probabile_nome_kanban(row):
                candidati.append(testo)
            elif len(testo) <= 80 and "barison" in testo.lower():
                candidati.append(testo)
        elif len(row) == 2 and normalizza_testo(row[0], compatta_spazi=True):
            primo = normalizza_testo(row[0], compatta_spazi=True)
            if re.search(r"jboss", primo, re.IGNORECASE):
                candidati.append(primo)

    if not candidati:
        return None

    barison = next((c for c in candidati if c.lower() == "barison"), None)
    jboss = next(
        (c for c in candidati if re.search(r"jboss", c, re.IGNORECASE)),
        None,
    )

    if barison and jboss:
        return KANBAN_NAME

    if barison:
        return KANBAN_NAME

    for candidato in candidati:
        if candidato.lower() == KANBAN_NAME.lower():
            return KANBAN_NAME

    return candidati[0]


def corrisponde_kanban_richiesta(nome_trovato):
    if not nome_trovato:
        return False
    trovato = normalizza_testo(nome_trovato, compatta_spazi=True).lower()
    richiesto = KANBAN_NAME.lower()
    if trovato == richiesto:
        return True
    if trovato.replace(" ", "") == richiesto.replace(" ", ""):
        return True
    if "jboss" in trovato and "barison" in trovato:
        return True
    if trovato == "barison":
        return True
    return False


def leggi_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def estrai_card_kanban(rows, kanban_name=KANBAN_NAME):
    card = []
    nome_sezione = None

    for i, row in enumerate(rows):
        if not is_kanban_header(row):
            continue

        nome = nome_kanban_da_contesto(rows, i)
        if not corrisponde_kanban_richiesta(nome):
            continue

        nome_sezione = nome or kanban_name
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
        break

    return nome_sezione, card


def estrai_tag_temporali(description):
    """Restituisce le righe di Description che iniziano con '#'."""
    if not description:
        return []
    tag = []
    for line in str(description).splitlines():
        if line.lstrip().startswith("#"):
            tag.append(line.strip())
    return tag


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


def giorni_tra_date(start_date, end_date):
    start = parse_data(start_date)
    end = parse_data(end_date)
    if start is None or end is None:
        return ""
    return (end - start).days


def espandi_card_con_tag(card):
    """Replica ogni card per tag temporale; aggiunge colonna TAG Temporali."""
    righe = []
    for record in card:
        tag_list = estrai_tag_temporali(record.get("Description", ""))
        giorni = giorni_tra_date(record.get("Start Date"), record.get("End Date"))
        if not tag_list:
            nuova = dict(record)
            nuova[GIORNI_COL] = giorni
            nuova[TAG_TEMPORALI_COL] = ""
            righe.append(nuova)
            continue
        for tag in tag_list:
            nuova = dict(record)
            nuova[GIORNI_COL] = giorni
            nuova[TAG_TEMPORALI_COL] = tag
            righe.append(nuova)
    return righe


def colonne_output():
    return KANBAN_COLUMNS + [GIORNI_COL, TAG_TEMPORALI_COL]


def scrivi_excel(card, output_path):
    righe = espandi_card_con_tag(card)
    df = pd.DataFrame(righe, columns=colonne_output())
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)

    wb = load_workbook(output_path)
    ws = wb[OUTPUT_SHEET]
    center = Alignment(horizontal="center", vertical="center")
    middle = Alignment(vertical="center")

    for row in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column,
    ):
        for cell in row:
            if cell.column in CENTER_COLS:
                cell.alignment = center
            else:
                cell.alignment = middle

    wb.save(output_path)


def elabora(input_csv):
    input_path = risolvi_percorso(input_csv)
    output_path = percorso_output_da_csv(input_path)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File di input non trovato: {input_path}")

    rows = leggi_csv(input_path)
    nome_sezione, card = estrai_card_kanban(rows)

    if not card:
        raise ValueError(
            f'Kanban "{KANBAN_NAME}" non trovata o senza card nel file {input_path}'
        )

    scrivi_excel(card, output_path)
    righe_output = len(espandi_card_con_tag(card))
    print(f"Kanban: {nome_sezione}")
    print(f"Card estratte: {len(card)}")
    print(f"Righe output: {righe_output}")
    print(f"Output: {output_path}")


def main():
    argv = sys.argv[1:]
    input_csv = argv[0] if len(argv) >= 1 else "2026-06-06-WIP.csv"
    elabora(input_csv)


if __name__ == "__main__":
    main()
