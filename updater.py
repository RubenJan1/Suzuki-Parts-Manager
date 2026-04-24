import os
import sys
import stat
import time
import shutil
import zipfile


def log(msg):
    print("[UPDATER]", msg)


# ─── Robuuste verwijdering (werkt ook bij OneDrive + read-only bestanden) ──────

def _force_writable(func, path, _exc_info):
    """Error-handler voor shutil.rmtree: zet bestand op schrijfbaar en probeer opnieuw."""
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        pass


def rmtree_robust(path, retries=6, delay=1.5):
    """
    Verwijdert een map met retries.
    Nodig omdat OneDrive bestanden tijdelijk kan vergrendelen tijdens sync.
    """
    for poging in range(retries):
        try:
            shutil.rmtree(path, onerror=_force_writable)
            return
        except Exception as e:
            if poging < retries - 1:
                log(f"Map verwijderen mislukt (poging {poging + 1}/{retries}): {e} — opnieuw proberen...")
                time.sleep(delay)
            else:
                raise RuntimeError(
                    f"Kan map niet verwijderen na {retries} pogingen: {path}\n\n"
                    "Tip: als de app in een OneDrive-map staat, zet OneDrive dan even op pauze "
                    "en probeer de update opnieuw."
                ) from e


# ─── Bestanden kopiëren ────────────────────────────────────────────────────────

def copy_robust(src, dst, retries=4, delay=1.0):
    """Kopieert een bestand met retries voor gelocked bestanden (OneDrive)."""
    for poging in range(retries):
        try:
            shutil.copy2(src, dst)
            return
        except Exception as e:
            if poging < retries - 1:
                time.sleep(delay)
            else:
                raise


def replace_files(src, dst):
    log("Bestanden vervangen...")
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)

        if os.path.isdir(s):
            if os.path.exists(d):
                rmtree_robust(d)
            shutil.copytree(s, d)
        else:
            copy_robust(s, d)


# ─── Wachten op afsluiting app ────────────────────────────────────────────────

def wait_for_app_to_close(pid: int):
    log(f"Wachten tot app sluit (PID {pid})...")
    import ctypes
    kernel32 = ctypes.windll.kernel32
    SYNCHRONIZE = 0x00100000
    handle = kernel32.OpenProcess(SYNCHRONIZE, False, pid)
    if handle:
        kernel32.WaitForSingleObject(handle, 30000)  # max 30 sec
        kernel32.CloseHandle(handle)
    time.sleep(1)


# ─── Hoofd ────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print("Dit programma wordt automatisch gestart door Suzuki Parts Manager.")
        print("Je hoeft het niet handmatig te openen.")
        input("\nDruk op Enter om te sluiten...")
        return

    zip_path = sys.argv[1]
    app_dir  = sys.argv[2]
    pid      = int(sys.argv[3]) if len(sys.argv) > 3 else None

    temp_dir = os.path.join(app_dir, "_update_tmp")

    # Verwijder eventuele resterende temp-map van vorige update
    if os.path.exists(temp_dir):
        log("Oude tijdelijke map opruimen...")
        rmtree_robust(temp_dir)

    os.makedirs(temp_dir)

    if pid:
        wait_for_app_to_close(pid)
    else:
        time.sleep(3)

    log("Uitpakken...")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    items = os.listdir(temp_dir)
    if not items:
        raise RuntimeError("Zip-bestand is leeg of kon niet worden uitgepakt.")

    inner_path = os.path.join(temp_dir, items[0])
    replace_files(inner_path, app_dir)

    log("Tijdelijke bestanden opruimen...")
    rmtree_robust(temp_dir)

    try:
        os.remove(zip_path)
    except Exception:
        pass

    exe_path = os.path.join(app_dir, "Suzuki Parts Manager.exe")
    log("Nieuwe versie starten...")
    os.startfile(exe_path)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"FOUT: {e}")
        input("\nDruk op Enter om te sluiten...")
