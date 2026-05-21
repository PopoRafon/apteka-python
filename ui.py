"""
NAZWA
    ui

OPIS
    Ten moduł implementuje graficzny interfejs użytkownika (GUI) dla systemu
    zarządzania apteką przy użyciu bibliotek tkinter oraz ttk. Integruje on
    funkcjonalności administracyjne magazynu oraz kartotekę pacjentów
    wraz z modułem sprzedaży.

    Interfejs został podzielony na nowoczesny system czterech zakładek, co pozwala
    na przejrzystą i odizolowaną pracę na bazach danych w formatach .xlsx oraz .csv.

    Ten skrypt wymaga zainstalowania bibliotek tkinter, pandas oraz openpyxl
    w środowisku Pythona, w którym jest uruchamiany.

FUNKCJE
    Ten moduł zawiera następujące funkcje:
    * init_ui() - inicjalizuje główne okno, konfigurację i uruchamia pętlę mainloop

PRZYKŁADY
    init_ui()
"""

import pharmacy_admin
import customer
import tkinter as tk
import os
import pandas as pd
from tkinter import messagebox, ttk
from datetime import date
from const import DRUGS_FILE_PATH, CUSTOMER_FILE_PATH, DRUGS_FIELDNAMES

def init_ui():
    customer.init_database()
    root = tk.Tk()
    root.title("Apteka")
    root.geometry("600x500")
    root.configure(background="#f4f6f9")
    style=ttk.Style()
    style.theme_use('clam')
    style.configure("TNoteBook", background="#f4f6f9",borderwidth=0)
    style.configure("TNotebook.Tab", font=("Helvetica", 10,"bold"),padding=[15, 5], background="#e0e6ed", foreground="#2c3e50")
    style.map("TNotebook.Tab", background=[("selected", "#2c3e50")], foreground=[("selected", "white")])
    notebook = ttk.Notebook(root)
    notebook.pack(padx=15,pady=15, fill="both", expand=True)
    tab_magazyn=tk.Frame(notebook, bg="white")
    tab_pacjenci=tk.Frame(notebook, bg="white")
    tab_sprzedaz=tk.Frame(notebook, bg="white")
    tab_statystyki=tk.Frame(notebook, bg="white")
    notebook.add(tab_magazyn,text="📦 Magazyn")
    notebook.add(tab_pacjenci,text="👥 Pacjenci")
    notebook.add(tab_sprzedaz,text="🛒 Sprzedaz")
    notebook.add(tab_statystyki,text="📊 Statystyki")
    label_wyniki_klienci = tk.Label(tab_statystyki, text="Zarejestrowani pacjenci: Ładowanie...",font=("Helvetica", 11, "bold"), bg="#2c3e50", fg="#ecf0f1")
    label_wyniki_leki = tk.Label(tab_statystyki, text="Dostępne pozycje leków: Ładowanie...",font=("Helvetica", 11, "bold"), bg="#2c3e50", fg="#ecf0f1")

    def pokaz_statystyki():
        """
        Odczytuje pliki bazy danych i odświeża statystyki apteki wyświetlane w GUI.

        Wczytuje plik customer.csv (automatycznie wykrywając separator) oraz plik
        drugs.xlsx, a następnie aktualizuje etykiety tekstowe w zakładce 'Statystyki'
        o aktualną liczbę rekordów.

        Zwraca:
            None
        """
        if os.path.exists(CUSTOMER_FILE_PATH):
            try:
                df_c = pd.read_csv(CUSTOMER_FILE_PATH, sep=None, engine='python', encoding="utf-8")
                label_wyniki_klienci.config(text=f"Zarejestrowani pacjenci: {len(df_c)}")
            except Exception:
                label_wyniki_klienci.config(text="Zarejestrowani pacjenci: Błąd struktury pliku")
        if os.path.exists(DRUGS_FILE_PATH):
            try:
                df_d = pd.read_excel(DRUGS_FILE_PATH)
                df_d.columns = df_d.columns.str.strip()
                label_wyniki_leki.config(text=f"Dostępne pozycje leków: {len(df_d)}")
            except Exception:
                label_wyniki_leki.config(text="Dostępne pozycje leków: 0")

    def ustaw_placeholder(entry,tekst):
        """
        Tworzy dynamiczny tekst pomocniczy (placeholder) dla podanego pola wprowadzania.

        Wstawia jasnoszary tekst podpowiedzi do pola Entry. Rejestruje zdarzenia
        FocusIn (czyszczenie pola i zmiana koloru czcionki na ciemny przy kliknięciu)
        oraz FocusOut (przywrócenie podpowiedzi, jeśli użytkownik nic nie wpisał).

        Argumenty:
            entry (tk.Entry): okno formularza, do którego przypisana jest podpowiedź
            tekst (str): treść wyświetlanej podpowiedzi (np. "ID leku")

        Zwraca:
            None
        """
        entry.insert(0,tekst)
        entry.config(fg="#a0aab2")
        def entry_click(event):
            if entry.get()==tekst:
                entry.delete(0,tk.END)
                entry.config(fg="#2c3e50")
        def focusout(event):
            if entry.get()=="":
                entry.insert(0,tekst)
                entry.config(fg="#a0aab2")
        entry.bind("<FocusIn>",entry_click)
        entry.bind("<FocusOut>",focusout)

    btn_style = {"font": ("Helvetica", 10, "bold"), "fg": "white", "bd": 0, "cursor": "hand2", "relief": "flat"}
    tk.Label(tab_magazyn, text="ZARZĄDZANIE ASORTYMENTEM APTEKI", font=("Helvetica", 12, "bold"), bg="white",fg="#2c3e50").pack(pady=15)
    frame_inputs_lek=tk.Frame(tab_magazyn, bg="white")
    frame_inputs_lek.pack(pady=10)
    entry_id_leku = tk.Entry(frame_inputs_lek, width=8, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_id_leku, "ID leku")
    entry_id_leku.pack(side=tk.LEFT, padx=4, ipady=3)
    entry_nazwa_leku = tk.Entry(frame_inputs_lek, width=16, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_nazwa_leku, "Nazwa leku")
    entry_nazwa_leku.pack(side=tk.LEFT, padx=4, ipady=3)
    entry_recepta = tk.Entry(frame_inputs_lek, width=14, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_recepta, "Recepta? (YES/NO)")
    entry_recepta.pack(side=tk.LEFT, padx=4, ipady=3)
    entry_ilosc = tk.Entry(frame_inputs_lek, width=8, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_ilosc, "Ilość")
    entry_ilosc.pack(side=tk.LEFT, padx=4, ipady=3)

    def dodaj_lek():
        """
        Pobiera dane z formularza magazynu i dopisuje nowy lek do bazy danych Excel.

        Konwertuje ID oraz Ilość na typy liczbowe. Sprawdza w pliku drugs.xlsx,
        czy lek o podanym ID już istnieje. Jeśli nie, tworzy nowy wiersz DataFrame
        i dołącza ga do pliku Excel, automatycznie generując dzisiejszą datę wpisu.

        Zwraca:
            None
        """
        try:
            d_id=int(entry_id_leku.get())
            nazwa=entry_nazwa_leku.get()
            recepta=entry_recepta.get()
            ilosc=int(entry_ilosc.get())
            dzisiejsza_data = date.today().strftime("%Y-%m-%d")
            df=pd.read_excel(DRUGS_FILE_PATH)
            df.columns = df.columns.str.strip()
            if d_id in df["ID"].values:
                raise ValueError(f"BŁĄD: Lek o ID {d_id} już istnieje!")
            nowy_lek = pd.DataFrame([[d_id, nazwa, recepta, ilosc, dzisiejsza_data]], columns=DRUGS_FIELDNAMES)
            df=pd.concat([df,nowy_lek],ignore_index=True)
            df.to_excel(DRUGS_FILE_PATH,index=False)
            messagebox.showinfo("Sukces",f"Dodano lek: {nazwa}")
            pokaz_statystyki()
        except ValueError as e:
            messagebox.showerror("Error",str(e))
        except Exception:
            messagebox.showerror("Error","Wpisz poprawne dane!")

    tk.Button(tab_magazyn, text="➕ Dodaj Nowy Lek", command=dodaj_lek, bg="#2ecc71", **btn_style).pack(pady=10,ipadx=15,ipady=3)
    frame_del_lek = tk.Frame(tab_magazyn, bg="white")
    frame_del_lek.pack(pady=25)
    entry_usun_lek_id = tk.Entry(frame_del_lek, width=20, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_usun_lek_id, "ID leku do usunięcia")
    entry_usun_lek_id.pack(side=tk.LEFT, padx=5, ipady=3)

    def usun_lek():
        """
        Przygotowuje dane i wywołuje procedurę usuwania leku z magazynu.

        Pobiera liczbowe ID leku z formularza, czyści potencjalne błędy w strukturze
        nagłówków pliku Excel (usuwa białe znaki), a następnie wywołuje zewnętrzną
        funkcję pharmacy_admin.remove_drug() w celu usunięcia leku.

        Zwraca:
            None
        """
        try:
            d_id = int(entry_usun_lek_id.get())
            if os.path.exists(DRUGS_FILE_PATH):
                df = pd.read_excel(DRUGS_FILE_PATH)
                df.columns = df.columns.str.strip()
                df.to_excel(DRUGS_FILE_PATH, index=False)
            pharmacy_admin.remove_drug(d_id)
            messagebox.showinfo("Sukces", f"Usunięto lek o ID {d_id}")
            pokaz_statystyki()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Wpisz poprawne dane! Szczegóły: {e}")

    tk.Button(frame_del_lek, text="🗑️ Usuń Lek", command=usun_lek, bg="#e74c3c", **btn_style).pack(side=tk.LEFT,ipadx=10,ipady=3)
    tk.Label(tab_pacjenci, text="KARTOTEKA I REJESTRACJA PACJENTÓW", font=("Helvetica", 12, "bold"), bg="white",fg="#2c3e50").pack(pady=15)
    frame_p1=tk.Frame(tab_pacjenci, bg="white")
    frame_p1.pack(pady=5)
    entry_imie_klienta=tk.Entry(frame_p1, width=22, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_imie_klienta, "Imię i Nazwisko")
    entry_imie_klienta.pack(side=tk.LEFT, padx=3, ipady=3)
    entry_email_klienta=tk.Entry(frame_p1, width=18, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_email_klienta, "E-mail")
    entry_email_klienta.pack(side=tk.LEFT, padx=3, ipady=3)
    entry_telefon_klienta=tk.Entry(frame_p1, width=14, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_telefon_klienta, "Telefon")
    entry_telefon_klienta.pack(side=tk.LEFT, padx=3, ipady=3)
    frame_p2=tk.Frame(tab_pacjenci, bg="white")
    frame_p2.pack(pady=5)
    entry_ulica_klienta=tk.Entry(frame_p2, width=20, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_ulica_klienta, "Ulica i nr")
    entry_ulica_klienta.pack(side=tk.LEFT, padx=3, ipady=3)
    entry_miasto_klienta=tk.Entry(frame_p2, width=17, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_miasto_klienta, "Miasto")
    entry_miasto_klienta.pack(side=tk.LEFT, padx=3, ipady=3)
    entry_kraj_klienta=tk.Entry(frame_p2, width=17, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_kraj_klienta, "Kraj")
    entry_kraj_klienta.pack(side=tk.LEFT, padx=3, ipady=3)

    def dodaj_klienta():
        """
        Zbiera kompletne dane adresowe i rejestruje nowego klienta w bazie danych.

        Pobiera z interfejsu Imię i Nazwisko, E-mail, Telefon, Ulicę, Miasto oraz Kraj.
        Przekazuje pełen zestaw argumentów bezpośrednio do funkcji bazy danych
        customer.add_customer().

        Zwraca:
            None
        """
        imie=entry_imie_klienta.get()
        email=entry_email_klienta.get()
        telefon=entry_telefon_klienta.get()
        ulica=entry_ulica_klienta.get()
        miasto=entry_miasto_klienta.get()
        kraj=entry_kraj_klienta.get()
        wiadomosc=customer.add_customer(name=imie,email=email,phone=telefon,street=ulica,city=miasto,country=kraj)
        if wiadomosc is None:
            messagebox.showinfo("Sukces",f"Zarejestrowane pacjenta {imie}")
            pokaz_statystyki()
        else:
            messagebox.showerror("ERROR",wiadomosc)

    tk.Button(tab_pacjenci, text="👤 Zarejestruj Pacjenta", command=dodaj_klienta, bg="#3498db", **btn_style).pack(pady=10, ipadx=15, ipady=3)
    frame_del_pacjent=tk.Frame(tab_pacjenci, bg="white")
    frame_del_pacjent.pack(pady=20)
    entry_usun_klienta_id=tk.Entry(frame_del_pacjent, width=22, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_usun_klienta_id, "ID klienta do usunięcia")
    entry_usun_klienta_id.pack(side=tk.LEFT, padx=5, ipady=3)

    def usun_klienta():
        """
        Wyrejestrowuje klienta z bazy danych na podstawie podanego ID.

        Pobiera liczbowy identyfikator z pola tekstowego i przekazuje go do funkcji
        customer.remove_customer(). Jeśli profil istnieje, usuwa go z plików bazowych.

        Zwraca:
            None
        """
        try:
            k_id=int(entry_usun_klienta_id.get())
            wiadomosc=customer.remove_customer(id=k_id,name=None)
            if wiadomosc is None:
                messagebox.showinfo("Sukces",f"Usunięto klienta o ID: {k_id}")
                pokaz_statystyki()
            else:
                messagebox.showerror("ERROR",wiadomosc)
        except Exception:
            messagebox.showerror("Error","Wpisz poprawne dane!")

    tk.Button(frame_del_pacjent, text="🗑️ Usuń Profil", command=usun_klienta, bg="#e74c3c", **btn_style).pack(side=tk.LEFT, ipadx=10, ipady=3)
    frame_inputs_sprzedaz=tk.Frame(tab_sprzedaz, bg="white")
    frame_inputs_sprzedaz.pack(pady=15)
    entry_id_klienta=tk.Entry(frame_inputs_sprzedaz, width=14, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_id_klienta, "ID klienta")
    entry_id_klienta.pack(side=tk.LEFT, padx=4, ipady=3)
    entry_kup_lek=tk.Entry(frame_inputs_sprzedaz, width=18, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_kup_lek, "Nazwa leku")
    entry_kup_lek.pack(side=tk.LEFT, padx=4, ipady=3)
    entry_kup_ilosc=tk.Entry(frame_inputs_sprzedaz, width=10, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_kup_ilosc, "Ilość")
    entry_kup_ilosc.pack(side=tk.LEFT, padx=4, ipady=3)
    entry_kup_recepta=tk.Entry(frame_inputs_sprzedaz, width=17, font=("Helvetica", 10), bd=1, relief="solid")
    ustaw_placeholder(entry_kup_recepta, "Recepta? (YES/NO)")
    entry_kup_recepta.pack(side=tk.LEFT, padx=4, ipady=3)

    def kup():
        """
        Przetwarza transakcję zakupu leku przez klienta (zarządzanie koszykiem).

        Pobiera tekstowe ID klienta, nazwę leku, recepte na lek oraz żądaną ilość. Blokuje transakcję,
        jeśli domyślne podpowiedzi pól (placeholdery) nie zostały zmienione. Formatuje
        dane do struktury słownika i przekazuje je do funkcji customer.customer_buy_drug()
        w celu zmniejszenia zapasów i zapisania historii operacji.

        Zwraca:
            None
        """
        try:
            kup_id_str = entry_id_klienta.get().strip()
            lek_nazwa_input = entry_kup_lek.get().strip()
            ile_str = entry_kup_ilosc.get().strip()
            lek_recepta = entry_kup_recepta.get().strip()
            if kup_id_str == "ID klienta" or lek_nazwa_input == "Nazwa leku" or ile_str == "Ilość" or lek_recepta == "Recepta? (YES/NO)":
                messagebox.showerror("Błąd", "Wypełnij najpierw wszystkie pola formularza zakupu!")
                return
            kup_id = int(kup_id_str)
            ile = int(ile_str)
            koszyk = [{'amount': ile, 'drug_name': lek_nazwa_input, 'prescription': lek_recepta}]
            wyniki = customer.customer_buy_drug(kup_id, koszyk)
            if wyniki is None:
                messagebox.showinfo("Sukces",f"Zrealizowano zakup!\nLek: {lek_nazwa_input}\nIlość: {ile} szt.")
                pokaz_statystyki()
            else:
                messagebox.showerror("Odrzucono", f"Zakup odrzucony! {wyniki[0]}")
        except ValueError:
            messagebox.showerror("Błąd formatu", "ID klienta oraz Ilość muszą być liczbami całkowitymi!")
        except Exception as e:
            messagebox.showerror("Błąd systemowy", f"Szczegóły: {e}")

    tk.Button(tab_sprzedaz, text="💳 Zrealizuj Zakup (Koszyk)", command=kup, bg="#9b59b6", **btn_style).pack(pady=15, ipadx=20, ipady=4)
    tk.Label(tab_statystyki, text="MONITOROWANIE SYSTEMU DANYCH", font=("Helvetica", 12, "bold"), bg="#2c3e50",fg="white").pack(pady=25)
    label_wyniki_klienci = tk.Label(tab_statystyki, text="Zarejestrowani pacjenci: Ładowanie...", font=("Helvetica", 11, "bold"), bg="#2c3e50", fg="#ecf0f1")
    label_wyniki_klienci.pack(pady=10)
    label_wyniki_leki = tk.Label(tab_statystyki, text="Dostępne pozycje leków: Ładowanie...", font=("Helvetica", 11, "bold"), bg="#2c3e50", fg="#ecf0f1")
    label_wyniki_leki.pack(pady=10)
    tk.Button(tab_statystyki, text="🔄 Odśwież Dane", command=pokaz_statystyki, bg="#34495e", **btn_style).pack(pady=30, ipadx=10, ipady=3)
    pokaz_statystyki()
    root.mainloop()
