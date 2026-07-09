#!/bin/bash

# Copia export MIRO (*-Intesa-JBOSS.csv) da Google Drive in Data/.
# Salta i file già presenti con lo stesso nome o già elaborati (prefisso x-).
# Percorsi sovrascrivibili con variabili d'ambiente DRIVE_DIR e LOCAL_DIR.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

DRIVE_DIR="${DRIVE_DIR:-$HOME/Library/CloudStorage/GoogleDrive-dpini@redhat.com/Il mio Drive/Customers/Miro}"
LOCAL_DIR="${LOCAL_DIR:-$SCRIPT_DIR/Data}"

if [ ! -d "$DRIVE_DIR" ]; then
    echo "Errore: la cartella sorgente su Google Drive non esiste o non è montata."
    echo "Verifica il percorso o imposta DRIVE_DIR: $DRIVE_DIR"
    exit 1
fi

mkdir -p "$LOCAL_DIR"

echo "Copia MIRO: $DRIVE_DIR/ → $LOCAL_DIR/"

copiati=0
saltati=0

for src in "$DRIVE_DIR"/*-Intesa-JBOSS.csv; do
    [ -e "$src" ] || continue

    base=$(basename "$src")

    if [ -e "$LOCAL_DIR/$base" ] || [ -e "$LOCAL_DIR/x-$base" ]; then
        saltati=$((saltati + 1))
        continue
    fi

    rsync -a "$src" "$LOCAL_DIR/"
    copiati=$((copiati + 1))
done

if [ "$copiati" -eq 0 ] && [ "$saltati" -eq 0 ]; then
    echo "Nessun file *-Intesa-JBOSS.csv trovato su Drive."
elif [ "$copiati" -eq 0 ]; then
    echo "Copia MIRO completata: nessun file nuovo ($saltati già presenti/elaborati)."
else
    echo "Copia MIRO completata: $copiati nuovo/i, $saltati già presenti/elaborati."
fi
