"""
elabora_Ass.py — Export CSV PSA in Excel con riepiloghi Assignment

Legge un file CSV (export PSA/pianificazione) e produce un file .xlsx omonimo con tre fogli:
  - data               : dati sorgente copiati dal CSV
  - Assignment         : aggregazione per risorsa / milestone / progetto OPA
  - Assignment Status  : aggregazione per risorsa / status / forecast category

Utilizzo:
    python elabora_Ass.py input.csv
    python elabora_Ass.py -h
"""

import argparse
import os
import sys

import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

COL_RESOURCE = "Resource: Full Name"
COL_MILESTONE = "Assignment: Milestone: Milestone Name"
COL_OPA = "Project: OPA Project Number"
COL_PLANNED = "Assignment: Milestone: Planned Hours"
COL_ESTIMATED = "Estimated Hours"
COL_ACTUAL = "Actual Hours"
COL_PROJECT_NAME = "Project: Project Name"
COL_ASSIGNMENT_NAME = "Assignment: Assignment Name"
COL_STATUS = "Assignment: Status"
COL_FORECAST = "Assignment: Forecast Category"

SHEET_DATA = "data"
SHEET_ASSIGNMENT = "Assignment"
SHEET_STATUS = "Assignment Status"


def risolvi_percorso(nome_o_path):
    if os.path.isabs(nome_o_path):
        return nome_o_path
    return os.path.normpath(os.path.join(SCRIPT_DIR, nome_o_path))


def leggi_csv(path):
    """Legge un CSV PSA provando encoding e separatore."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            df = pd.read_csv(path, sep=None, engine="python", encoding=enc)
            if len(df.columns) > 1:
                return df
        except Exception:
            continue
    for sep in (";", ","):
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(path, sep=sep, encoding=enc)
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue
    raise ValueError(f"Impossibile leggere il CSV: {path}")


def _norm_col_name(name):
    return str(name).strip().casefold()


def trova_colonna(df, nome_atteso, alias=()):
    """Restituisce il nome reale della colonna nel DataFrame."""
    if nome_atteso in df.columns:
        return nome_atteso
    attesi = {_norm_col_name(nome_atteso)}
    for alt in alias:
        attesi.add(_norm_col_name(alt))
    for col in df.columns:
        if _norm_col_name(col) in attesi:
            return col
    raise KeyError(
        f"Colonna non trovata: {nome_atteso!r}. "
        f"Colonne disponibili: {', '.join(map(str, df.columns))}"
    )


def normalizza_ore(df, *colonne):
    for col in colonne:
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace(",", ".", regex=False),
            errors="coerce",
        ).fillna(0.0)
    return df


def percorso_output_da_csv(path_csv):
    base, _ = os.path.splitext(path_csv)
    return base + ".xlsx"


def tabella_assignment(df):
    col_res = trova_colonna(df, COL_RESOURCE, ("Full Name",))
    col_mil = trova_colonna(df, COL_MILESTONE)
    col_opa = trova_colonna(df, COL_OPA)
    col_planned = trova_colonna(df, COL_PLANNED, ("Planned Hours",))
    col_estimated = trova_colonna(df, COL_ESTIMATED)
    col_actual = trova_colonna(df, COL_ACTUAL)
    col_proj = trova_colonna(df, COL_PROJECT_NAME)
    col_assign = trova_colonna(df, COL_ASSIGNMENT_NAME, ("Assignment Name",))

    work = normalizza_ore(
        df.copy(), col_planned, col_estimated, col_actual
    )
    grouped = work.groupby([col_res, col_mil, col_opa], dropna=False, as_index=False).agg(
        **{
            COL_RESOURCE: (col_res, "first"),
            COL_MILESTONE: (col_mil, "first"),
            COL_OPA: (col_opa, "first"),
            COL_PLANNED: (col_planned, "sum"),
            COL_ESTIMATED: (col_estimated, "sum"),
            COL_ACTUAL: (col_actual, "sum"),
            COL_PROJECT_NAME: (col_proj, "first"),
            COL_ASSIGNMENT_NAME: (col_assign, "first"),
        }
    )
    return grouped[
        [
            COL_RESOURCE,
            COL_MILESTONE,
            COL_OPA,
            COL_PLANNED,
            COL_ESTIMATED,
            COL_ACTUAL,
            COL_PROJECT_NAME,
            COL_ASSIGNMENT_NAME,
        ]
    ]


def tabella_assignment_status(df):
    col_res = trova_colonna(df, COL_RESOURCE, ("Full Name",))
    col_status = trova_colonna(df, COL_STATUS, ("Assignment Status",))
    col_forecast = trova_colonna(df, COL_FORECAST)
    col_actual = trova_colonna(df, COL_ACTUAL)

    work = normalizza_ore(df.copy(), col_actual)
    grouped = work.groupby([col_res, col_status, col_forecast], dropna=False, as_index=False).agg(
        **{
            COL_RESOURCE: (col_res, "first"),
            COL_STATUS: (col_status, "first"),
            COL_FORECAST: (col_forecast, "first"),
            COL_ACTUAL: (col_actual, "sum"),
        }
    )
    return grouped[[COL_RESOURCE, COL_STATUS, COL_FORECAST, COL_ACTUAL]]


def scrivi_excel(path_output, df_data, df_assignment, df_status):
    with pd.ExcelWriter(path_output, engine="openpyxl") as writer:
        df_data.to_excel(writer, sheet_name=SHEET_DATA, index=False)
        df_assignment.to_excel(writer, sheet_name=SHEET_ASSIGNMENT, index=False)
        df_status.to_excel(writer, sheet_name=SHEET_STATUS, index=False)


def elabora(path_csv):
    path_csv = risolvi_percorso(path_csv)
    if not os.path.isfile(path_csv):
        raise FileNotFoundError(f"File non trovato: {path_csv}")
    if not path_csv.lower().endswith(".csv"):
        raise ValueError("Il file di input deve avere estensione .csv")

    df = leggi_csv(path_csv)
    df_assignment = tabella_assignment(df)
    df_status = tabella_assignment_status(df)

    path_output = percorso_output_da_csv(path_csv)
    scrivi_excel(path_output, df, df_assignment, df_status)
    return path_output


def main():
    parser = argparse.ArgumentParser(
        description="Converte un export CSV PSA in Excel con fogli data, Assignment e Assignment Status.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Esempi:
  python elabora_Ass.py export-assignment.csv
  python elabora_Ass.py /percorso/export-assignment.csv

Output:
  <input>.xlsx — stesso percorso del CSV, estensione .xlsx.
  Fogli: data, Assignment, Assignment Status.
""",
    )
    parser.add_argument(
        "input_csv",
        help="Percorso del file CSV di input (.csv)",
    )
    args = parser.parse_args()

    try:
        path_output = elabora(args.input_csv)
    except (FileNotFoundError, ValueError, KeyError) as exc:
        print(f"ERRORE: {exc}", file=sys.stderr)
        return 1

    print(f"File generato: {path_output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
