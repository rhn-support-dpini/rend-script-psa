"""
elabora_JKAN.py — Estrae le card dalla Kanban MIRO "JBOSS-Barison"

Legge un export CSV di una board MIRO, individua la sezione Kanban con quel nome
e produce un file Excel con una riga per card.

Utilizzo:
    python elabora_JKAN.py [input.csv] [output.xlsx]

    Default: input=2026-06-06-WIP.csv, output=elaborato_JKAN.xlsx
"""

import csv
import os
import re
import sys

import pandas as pd

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
OUTPUT_SHEET = "data"


def risolvi_percorso(nome_o_path):
    if os.path.isabs(nome_o_path):
        return nome_o_path
    return os.path.normpath(os.path.join(SCRIPT_DIR, nome_o_path))


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


def scrivi_excel(card, output_path):
    df = pd.DataFrame(card, columns=KANBAN_COLUMNS)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)


def elabora(input_csv, output_xlsx):
    input_path = risolvi_percorso(input_csv)
    output_path = risolvi_percorso(output_xlsx)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File di input non trovato: {input_path}")

    rows = leggi_csv(input_path)
    nome_sezione, card = estrai_card_kanban(rows)

    if not card:
        raise ValueError(
            f'Kanban "{KANBAN_NAME}" non trovata o senza card nel file {input_path}'
        )

    scrivi_excel(card, output_path)
    print(f'Kanban: {nome_sezione}')
    print(f'Card estratte: {len(card)}')
    print(f'Output: {output_path}')


def main():
    argv = sys.argv[1:]
    input_csv = argv[0] if len(argv) >= 1 else "2026-06-06-WIP.csv"
    output_xlsx = argv[1] if len(argv) >= 2 else "elaborato_JKAN.xlsx"
    elabora(input_csv, output_xlsx)


if __name__ == "__main__":
    main()
