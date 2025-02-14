import random
import string
from faker import Faker
import openpyxl
import sys

# Genera un codice alfanumerico di 10 caratteri, assicurando l'unicità rispetto a quelli già generati
def genera_codice_unico(codici_esistenti):
    while True:
        # Qua prenso solo 10 caratteri casuali tra lettere maiuscole e numeri
        caratteri = random.choices(string.ascii_uppercase + string.digits, k=10)
        # Mischio i caratteri per variare ulteriormente la sequenza
        random.shuffle(caratteri)
        codice = ''.join(caratteri)
        # Verifica che il codice non sia già stato usato
        if codice not in codici_esistenti:
            codici_esistenti.add(codice)
            return codice

# Crea una lista di utenti casuali
def genera_utenti(num_utenti):
    fake = Faker('it_IT')
    utenti = []
    codici_esistenti = set()

    for _ in range(num_utenti):
        # Genero i dati fittizi generando così utenti casuali e email e telefono casuali
        nome = fake.first_name()
        cognome = fake.last_name()
        email = fake.email()
        telefono = fake.phone_number()
        # Otteniamo un codice univoco per un utente
        codice = genera_codice_unico(codici_esistenti)
        utenti.append([nome, cognome, email, telefono, codice])

    return utenti

# Salva la lista degli utenti nel file excel passato negli argomenti del main
def salva_su_excel(utenti, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Utenti"
    # Inserisco l'intestazione
    ws.append(["Nome", "Cognome", "Email", "Telefono", "Codice"])

    # Aggiungo ogni utente come nuova riga nel foglio
    for utente in utenti:
        ws.append(utente)

    wb.save(filename)
    print(f"File Excel '{filename}' creato con successo.")

if __name__ == "__main__":
    # Controlla che sia stato passato il nome del file come argomento
    if len(sys.argv) != 2:
        print("Devi specificare il nome del file (inclusa l'estensione .xlsx).")
        sys.exit(1)

    utenti = genera_utenti(10)
    salva_su_excel(utenti, sys.argv[1])
