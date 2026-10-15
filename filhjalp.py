"""
Hjälpfunktioner för filhantering som delas mellan skripten.
"""

import sys


def save_with_retry(save_func, output_file) -> None:
    """
    Försöker spara en fil. Vid PermissionError (t.ex. filen öppen i Excel)
    får användaren välja att försöka igen eller avsluta programmet.
    """
    while True:
        try:
            save_func()
            return
        except PermissionError:
            print(f"\nKunde inte skriva till '{output_file}'.")
            print("Filen kan vara öppen i ett annat program (t.ex. Excel).")
            print("  1. Försök igen")
            print("  2. Avsluta programmet")
            val = input("Välj (1/2): ").strip()
            if val == "2":
                print("Avslutar.")
                sys.exit(1)
            print("Försöker igen...")
