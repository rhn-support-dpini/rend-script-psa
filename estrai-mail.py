import os
import re

# Cartella export (default: JBOSS-Barison-Mail su Drive; override: EMAIL_DIR)
DIR_PATH = os.environ.get(
    "EMAIL_DIR",
    os.path.expanduser(
        "~/Library/CloudStorage/GoogleDrive-dpini@redhat.com/"
        "Il mio Drive/Customers/Miro/JBOSS-Barison-Mail"
    ),
)

# File di output (override: EMAIL_LST)
OUTPUT_FILE = os.environ.get("EMAIL_LST", "email.lst")


def analizza_file_email(filepath):
    """Legge un file .txt generato dallo script Apps Script

    ed estrae le informazioni DA, DATA e OGGETTO.
    """
    email_estratte = []

    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            contenuto = f.read()

        # Separa le singole email usando il delimitatore dello script Apps Script
        blocchi_email = contenuto.split(
            "=================================================="
        )

        for blocco in blocchi_email:
            if not blocco.strip():
                continue

            # Regex per estrarre le intestazioni principali
            da_match = re.search(r"^DA:\s*(.*)$", blocco, re.MULTILINE)
            data_match = re.search(r"^DATA:\s*(.*)$", blocco, re.MULTILINE)
            oggetto_match = re.search(r"^OGGETTO:\s*(.*)$", blocco, re.MULTILINE)

            da = da_match.group(1).strip() if da_match else "N/D"
            data = data_match.group(1).strip() if data_match else "N/D"
            oggetto = oggetto_match.group(1).strip() if oggetto_match else "N/D"

            # Aggiunge alla lista solo se abbiamo trovato informazioni valide
            if da != "N/D" or data != "N/D" or oggetto != "N/D":
                riga = f"DATA: {data} | DA: {da} | OGGETTO: {oggetto}"
                email_estratte.append(riga)

    except Exception as e:
        print(f"Errore nella lettura del file {filepath}: {e}")

    return email_estratte


def main():
    if not os.path.exists(DIR_PATH):
        print(f"Errore: Il percorso specificato non esiste:\n{DIR_PATH}")
        return

    # Recupera tutti i file presenti nella cartella
    file_list = [
        f
        for f in os.listdir(DIR_PATH)
        if os.path.isfile(os.path.join(DIR_PATH, f))
        and not f.startswith(".")  # Esclude file nascosti/sistema
    ]

    print(f"Trovati {len(file_list)} file da analizzare.")

    tutte_le_righe = []

    # Legge ed estrae le righe da tutti i file
    for file_name in file_list:
        full_path = os.path.join(DIR_PATH, file_name)
        print(f"Analisi di: {file_name}...")
        estratte = analizza_file_email(full_path)
        tutte_le_righe.extend(estratte)

    totale_iniziale = len(tutte_le_righe)

    # Rimuove i duplicati mantenendo l'ordine di estrazione
    righe_uniche = []
    viste = set()

    for riga in tutte_le_righe:
        if riga not in viste:
            viste.add(riga)
            righe_uniche.append(riga)

    totale_finale = len(righe_uniche)
    duplicati_rimossi = totale_iniziale - totale_finale

    # Scrive le righe uniche nel file output
    with open(OUTPUT_FILE, "w", encoding="utf-8") as out:
        for riga in righe_uniche:
            out.write(riga + "\n")

    # Resoconto a schermo
    print("\n--------------------------------------------------")
    print(f"Righe totali estrate (incluse doppie): {totale_iniziale}")
    print(f"Righe duplicate rimosse:              {duplicati_rimossi}")
    print(f"Righe uniche salvate in '{OUTPUT_FILE}':   {totale_finale}")
    print("--------------------------------------------------")


if __name__ == "__main__":
    main()
