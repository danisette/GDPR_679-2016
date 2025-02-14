import sqlite3
import openpyxl
import sys
import os


def leggi_dati_excel(filename):
    # Apro il file Excel e prendo il foglio attivo
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    dati = []

    # Leggo tutte le righe a partire dalla seconda (salto l'intestazione)
    for row in ws.iter_rows(min_row=2, values_only=True):
        dati.append(row)
    return dati


def crea_tabella_sql(dati, db_name):
    # Apro una connessione al database (viene creato se non esiste)
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    # Creo la tabella 'utenti' se non c'è già
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS utenti (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT,
        cognome TEXT,
        email TEXT,
        telefono TEXT,
        codice TEXT
    )
    ''')

    # Inserisco i dati nella tabella
    cursor.executemany('''
    INSERT INTO utenti (nome, cognome, email, telefono, codice)
    VALUES (?, ?, ?, ?, ?)
    ''', dati)

    conn.commit()
    conn.close()
    print(f"Dati inseriti nel database '{db_name}'.")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Il file excel di input non è stato specificato tra gli argomenti.")
        sys.exit(1)

    excel_file = sys.argv[1]
    # Imposto il nome del database usando lo stesso nome del file Excel, con estensione .db
    db_file = os.path.splitext(excel_file)[0] + ".db"

    dati = leggi_dati_excel(excel_file)
    crea_tabella_sql(dati, db_file)
