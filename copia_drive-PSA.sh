#!/bin/bash

# Copia export PSA (DvPFilter-RHCC-All-Assignments-Full) da Google Drive in Data/.
# Percorsi sovrascrivibili con variabili d'ambiente DRIVE_DIR e LOCAL_DIR.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

DRIVE_DIR="${DRIVE_DIR:-$HOME/Library/CloudStorage/GoogleDrive-dpini@redhat.com/Il mio Drive/Customers/DvPFilter-RHCC-All-Assignments-Full}"
LOCAL_DIR="${LOCAL_DIR:-$SCRIPT_DIR/Data}"

if [ ! -d "$DRIVE_DIR" ]; then
    echo "Errore: la cartella sorgente su Google Drive non esiste o non è montata."
    echo "Verifica il percorso o imposta DRIVE_DIR: $DRIVE_DIR"
    exit 1
fi

mkdir -p "$LOCAL_DIR"

echo "Copia PSA: $DRIVE_DIR/ → $LOCAL_DIR/"
rsync -av --ignore-existing "$DRIVE_DIR/" "$LOCAL_DIR/"

echo "Copia PSA completata."
