import pandas as pd
import os
from const import DRUGS_FILE_PATH, DRUGS_FIELDNAMES

def load_drugs():
    """
    Wczytuje dane leków z pliku XLSX. Tworzy pusty plik, jeśli nie istnieje.
    """
    if not os.path.exists(DRUGS_FILE_PATH):
        df = pd.DataFrame(columns=DRUGS_FIELDNAMES)
        df.to_excel(DRUGS_FILE_PATH, index=False)
        print("Utworzono nowy plik drugs.xlsx")
    else:
        df = pd.read_excel(DRUGS_FILE_PATH)
    return df

def show_drugs():
    """
    Wyświetla aktualną bazę leków.
    """
    df = load_drugs()
    print("\nAktualna baza leków:")
    print(df.to_string(index=False))

def add_drug(drug_id, name, on_recept, packages, date):
    """
    Dodaje nowy lek do bazy. Obsługuje wyjątek, gdy ID już istnieje.
    """
    df = load_drugs()
    if drug_id in df["ID"].values:
        raise ValueError(f"Lek o ID {drug_id} już istnieje.")
    new_row = pd.DataFrame([[drug_id, name, on_recept, packages, date]],
                           columns=df.columns)
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(DRUGS_FILE_PATH, index=False)
    print(f"Dodano lek: {name}")

def remove_drug(drug_id):
    """
    Usuwa lek z bazy po ID. Obsługuje wyjątek, gdy ID nie istnieje.
    """
    df = load_drugs()
    if drug_id not in df["ID"].values:
        raise ValueError(f"Nie znaleziono leku o ID {drug_id}.")
    df = df[df["ID"] != drug_id]
    df.to_excel(DRUGS_FILE_PATH, index=False)
    print(f"Usunięto lek o ID {drug_id}")
