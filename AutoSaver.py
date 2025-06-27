import win32com.client
import time

def auto_save_documents(interval_sec):
    try:
        inventor = win32com.client.Dispatch("Inventor.Application")
        inventor.Visible = True
    except Exception as e:
        print("Nie udało się połączyć z Autodesk Inventor:", e)
        return

    last_saved = {}  # słownik: InternalName -> timestamp ostatniego zapisu

    print(f"Auto-zapisywanie co {interval_sec} sekund. Naciśnij Ctrl+C, aby zakończyć.")

    try:
        while True:
            docs = inventor.Documents
            now = time.time()

            if docs.Count == 0:
                print("Brak otwartych dokumentów.")
            else:
                for i in range(1, docs.Count + 1):
                    doc = docs.Item(i)
                    internal_name = doc.InternalName

                    # Jeśli dokument był zapisany wcześniej niż X sekund temu
                    last = last_saved.get(internal_name, 0)
                    time_since_last_save = now - last

                    if doc.Dirty:
                        if time_since_last_save >= interval_sec:
                            doc.Save()
                            last_saved[internal_name] = now
                            print(f"[Zapisano] {doc.DisplayName}")
                        else:
                            print(f"[Pominięto] {doc.DisplayName} - ostatni zapis {int(time_since_last_save)} sek. temu")
                    else:
                        print(f"[Bez zmian] {doc.DisplayName}")

            time.sleep(interval_sec)

    except KeyboardInterrupt:
        print("\nZatrzymano auto-zapisywanie.")

# --- START ---

if __name__ == "__main__":
    try:
        interval = int(input("Podaj interwał zapisu (w sekundach): "))
        auto_save_documents(interval)
    except ValueError:
        print("Nieprawidłowy czas. Podaj liczbę całkowitą.")
