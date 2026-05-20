# Pharmacy – Moduł Administratora Bazy Leków

## Rola: Administrator Bazy i Towaru
Osoba odpowiedzialna za przygotowanie i utrzymanie bazy leków.

### Zadania:
- Tworzy strukturę plików (CSV dla leków, klientów i adresów).
- Pisze funkcje do dodawania, usuwania i wyświetlania leków.
- Obsługuje wyjątki (np. brak ID, duplikaty).
- Dba o poprawność danych i zapis do pliku.

### Pliki:
- DATABASE/drugs.csv – baza leków
- pharmacy_admin.py – funkcje administratora
- main.py – testowanie funkcji

### Przykład użycia:
```python
from pharmacy_admin import add_drug, remove_drug, show_drugs
add_drug(5, "ASPIRIN", "NO", 1000, "2026-06-01")
remove_drug(2)
show_drugs()
