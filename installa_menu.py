import sys
import os
from pathlib import Path
import winreg

KEY_PATH = r"Software\Classes\*\shell\convertPDF"

SUPPORTED_EXTS = [
    ".doc", ".docx", ".rtf", ".odt",
    ".xls", ".xlsx", ".xlsm", ".csv",
    ".ppt", ".pptx"
]

def trova_pythonw():
    """Trova il percorso di pythonw.exe."""
    # 1. Accanto all'interprete corrente
    cand = Path(sys.executable).parent / "pythonw.exe"
    if cand.exists():
        return cand

    # 2. Percorso standard di Python 3.12 dell'utente
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    p312 = Path(local_app_data) / "Programs" / "Python" / "Python312" / "pythonw.exe"
    if p312.exists():
        return p312

    # 3. Fallback: sys.executable
    return Path(sys.executable)

def installa():
    pythonw = trova_pythonw()
    script = (Path(__file__).parent / "convertPDF.py").resolve()

    if not script.exists():
        print(f"Errore: file non trovato: {script}")
        return False

    applies_to = " OR ".join(f"System.FileExtension:={ext}" for ext in SUPPORTED_EXTS)
    cmd = f'"{pythonw}" "{script}" "%1"'

    try:
        # Crea la chiave principale
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, KEY_PATH) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Converti in PDF")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "shell32.dll,275")
            winreg.SetValueEx(key, "AppliesTo", 0, winreg.REG_SZ, applies_to)

        # Crea la sottochiave command
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"{KEY_PATH}\\command") as cmd_key:
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, cmd)

        msg = (
            "Installazione completata con successo!\n\n"
            "Ora facendo tasto destro su file Word, Excel o PowerPoint troverai la voce 'Converti in PDF'.\n"
            "Il processo girerà in background e salverà il PDF nella stessa cartella del file."
        )
        print(msg)
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, msg, "convertPDF - Installato", 0x40)
        except Exception:
            pass
        return True
    except Exception as e:
        print(f"Errore durante l'installazione nel registro: {e}")
        return False

def disinstalla():
    try:
        # Cancella sottochiave command
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, f"{KEY_PATH}\\command")
        except FileNotFoundError:
            pass
        # Cancella chiave principale
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, KEY_PATH)

        msg = "Voce 'Converti in PDF' rimossa dal menu del tasto destro."
        print(msg)
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, msg, "convertPDF - Disinstallato", 0x40)
        except Exception:
            pass
        return True
    except FileNotFoundError:
        print("La voce non era presente nel registro.")
        return True
    except Exception as e:
        print(f"Errore durante la disinstallazione: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].lower() in ("--rimuovi", "--disinstalla", "-u", "/u"):
        disinstalla()
    else:
        installa()
