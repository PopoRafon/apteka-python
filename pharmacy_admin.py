"""
NAZWA
    pharmacy_admin

OPIS
    Ten moduł implementuje funkcje do dodawania, usuwania i wyświetlania leków.
    Tworzy strukturę plików CSV w bazie danych do zarządzania lekami.

    Ten skrypt wymaga zainstalowania bibliotek pandas oraz openpyxl
    w środowisku Pythona, w którym jest uruchamiany.

FUNKCJE
    * load_drugs()

    * show_drugs()

    * add_drug(drug_id, name, on_recept, packages, date)

    * remove_drug(drug_id)

PRZYKŁADY
    add_drug(5, "ASPIRIN", "NO", 1000, "2026-06-01")
    remove_drug(2)
    show_drugs()
"""

import pandas as pd
import os
from const import DRUGS_FILE_PATH, DRUGS_FIELDNAMES

def load_drugs():
    """
    Wczytuje dane leków z pliku XLSX. Tworzy pusty plik, jeśli nie istnieje.

    Zwraca:
        dataframe (pd.DataFrame): dane dotyczace leków z pliku XLSX
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

    Zwraca:
        None
    """
    df = load_drugs()
    print("\nAktualna baza leków:")
    print(df.to_string(index=False))

def add_drug(drug_id, name, on_recept, packages, date):
    """
    Dodaje nowy lek do bazy. Obsługuje wyjątek, gdy ID już istnieje.

    Argumenty:
        drug_id (int): id nowego leku
        name (str): nazwa nowego leku
        on_recept (str): recepta na nowy lek
        packages (int): ilość nowego leku
        date (str): data dodania nowego leku

    Zwraca:
        None
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

    Argumenty:
        drug_id (int): id leku do usunięcia

    Zwraca:
        None
    """
    df = load_drugs()
    if drug_id not in df["ID"].values:
        raise ValueError(f"Nie znaleziono leku o ID {drug_id}.")
    df = df[df["ID"] != drug_id]
    df.to_excel(DRUGS_FILE_PATH, index=False)
    print(f"Usunięto lek o ID {drug_id}")
