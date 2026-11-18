import sys
import os
import subprocess
from pathlib import Path

def mostra_errore(titolo, messaggio):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, str(messaggio), str(titolo), 0x10)  # 0x10 = MB_ICONERROR
    except Exception:
        pass

def notifica_successo(src_path, dst_path):
    try:
        import winsound
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass

# Verifica presenza di win32com
HAS_WIN32COM = False
try:
    import win32com.client
    HAS_WIN32COM = True
except ModuleNotFoundError:
    HAS_WIN32COM = False
    # Tenta installazione automatica in background di pywin32 per il futuro
    try:
        subprocess.Popen(
            [sys.executable, "-m", "pip", "install", "pywin32"],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        )
    except Exception:
        pass

WORD = {".doc", ".docx", ".rtf", ".odt"}
EXCEL = {".xls", ".xlsx", ".xlsm", ".csv"}
PPT = {".ppt", ".pptx"}

def adatta_tabelle_word_com(doc, forza_landscape=False, auto_landscape_colonne=5):
    wdAutoFitWindow = 2
    wdAlignRowCenter = 1
    wdOrientLandscape = 1

    ha_tabelle_larghe = False

    for table in doc.Tables:
        try:
            table.AutoFitBehavior(wdAutoFitWindow)
        except Exception:
            pass

        try:
            table.Rows.Alignment = wdAlignRowCenter
        except Exception:
            pass

        try:
            table.Rows(1).HeadingFormat = -1
        except Exception:
            pass

        try:
            table.Rows.AllowBreakAcrossPages = 0
        except Exception:
            pass

        try:
            if auto_landscape_colonne and table.Columns.Count >= auto_landscape_colonne:
                ha_tabelle_larghe = True
        except Exception:
            pass

    if forza_landscape or ha_tabelle_larghe:
        try:
            doc.PageSetup.Orientation = wdOrientLandscape
            for table in doc.Tables:
                table.AutoFitBehavior(wdAutoFitWindow)
        except Exception:
            pass

def converti_word_ps(src, dst, landscape=False):
    src_str = str(src).replace("'", "''")
    dst_str = str(dst).replace("'", "''")
    orient_cmd = "$doc.PageSetup.Orientation = 1;" if landscape else ""
    ps_code = f"""
    $ErrorActionPreference = 'Stop'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try {{
        $doc = $word.Documents.Open('{src_str}', $false, $true)
        try {{
            foreach ($tbl in $doc.Tables) {{
                try {{ $tbl.AutoFitBehavior(2) }} catch {{}}
                try {{ $tbl.Rows.Alignment = 1 }} catch {{}}
                try {{ $tbl.Rows.Item(1).HeadingFormat = -1 }} catch {{}}
                try {{ $tbl.Rows.AllowBreakAcrossPages = 0 }} catch {{}}
            }}
            {orient_cmd}
            $doc.ExportAsFixedFormat('{dst_str}', 17)
        }} finally {{
            $doc.Close($false)
        }}
    }} finally {{
        $word.Quit()
    }}
    """
    res = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_code],
        capture_output=True,
        text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    )
    if res.returncode != 0:
        err = res.stderr.strip() or res.stdout.strip()
        raise RuntimeError(f"PowerShell Word COM fallito: {err}")

def converti_excel_ps(src, dst):
    src_str = str(src).replace("'", "''")
    dst_str = str(dst).replace("'", "''")
    ps_code = f"""
    $ErrorActionPreference = 'Stop'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {{
        $wb = $excel.Workbooks.Open('{src_str}', $false, $true)
        try {{
            $wb.ExportAsFixedFormat(0, '{dst_str}')
        }} finally {{
            $wb.Close($false)
        }}
    }} finally {{
        $excel.Quit()
    }}
    """
    res = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_code],
        capture_output=True,
        text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    )
    if res.returncode != 0:
        err = res.stderr.strip() or res.stdout.strip()
        raise RuntimeError(f"PowerShell Excel COM fallito: {err}")

def converti_ppt_ps(src, dst):
    src_str = str(src).replace("'", "''")
    dst_str = str(dst).replace("'", "''")
    ps_code = f"""
    $ErrorActionPreference = 'Stop'
    $ppt = New-Object -ComObject PowerPoint.Application
    try {{
        $pres = $ppt.Presentations.Open('{src_str}', $true, $false, $false)
        try {{
            $pres.SaveAs('{dst_str}', 32)
        }} finally {{
            $pres.Close()
        }}
    }} finally {{
        $ppt.Quit()
    }}
    """
    res = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_code],
        capture_output=True,
        text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    )
    if res.returncode != 0:
        err = res.stderr.strip() or res.stdout.strip()
        raise RuntimeError(f"PowerShell PowerPoint COM fallito: {err}")

def converti(src, dst=None, adatta_tabelle=True, landscape=False):
    src = Path(src).resolve()
    dst = Path(dst).resolve() if dst else src.with_suffix(".pdf")
    ext = src.suffix.lower()

    if not src.exists():
        raise FileNotFoundError(f"File di origine non trovato: {src}")

    if ext in WORD:
        if HAS_WIN32COM:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            try:
                doc = app.Documents.Open(str(src), ReadOnly=True)
                if adatta_tabelle and doc.Tables.Count > 0:
                    adatta_tabelle_word_com(doc, forza_landscape=landscape)
                doc.ExportAsFixedFormat(str(dst), 17)
                doc.Close(False)
            finally:
                app.Quit()
        else:
            converti_word_ps(src, dst, landscape=landscape)

    elif ext in EXCEL:
        if HAS_WIN32COM:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                wb = app.Workbooks.Open(str(src), ReadOnly=True)
                wb.ExportAsFixedFormat(0, str(dst))
                wb.Close(False)
            finally:
                app.Quit()
        else:
            converti_excel_ps(src, dst)

    elif ext in PPT:
        if HAS_WIN32COM:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                pres = app.Presentations.Open(str(src), ReadOnly=True, WithWindow=False)
                pres.SaveAs(str(dst), 32)
                pres.Close()
            finally:
                app.Quit()
        else:
            converti_ppt_ps(src, dst)

    else:
        raise ValueError(f"Formato non supportato: {ext}")

    return dst

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: convertPDF.py [--landscape] <file1> [file2 ...]")
        sys.exit(1)

    landscape = "--landscape" in sys.argv
    files = [arg for arg in sys.argv[1:] if not arg.startswith("--")]

    is_gui = "pythonw" in sys.executable.lower()

    for f in files:
        try:
            res = converti(f, landscape=landscape)
            if is_gui:
                notifica_successo(f, res)
            else:
                print("OK ->", res)
        except Exception as e:
            err_msg = f"Impossibile convertire '{Path(f).name}':\n\n{e}"
            if is_gui:
                mostra_errore("Errore convertPDF", err_msg)
            else:
                print(f"Errore su {f}: {e}")