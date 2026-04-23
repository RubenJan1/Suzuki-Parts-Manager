import os
import sys
import time
import shutil
import zipfile

def log(msg):
    print("[UPDATER]", msg)

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

def extract_zip(zip_path, extract_to):
    log("Uitpakken...")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

def replace_files(src, dst):
    log("Bestanden vervangen...")
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)

        if os.path.isdir(s):
            if os.path.exists(d):
                shutil.rmtree(d)
            shutil.copytree(s, d)
        else:
            shutil.copy2(s, d)

def main():
    if len(sys.argv) < 3:
        print("Dit programma wordt automatisch gestart door Suzuki Parts Manager.")
        print("Je hoeft het niet handmatig te openen.")
        input("\nDruk op Enter om te sluiten...")
        return

    zip_path = sys.argv[1]
    app_dir = sys.argv[2]
    pid = int(sys.argv[3]) if len(sys.argv) > 3 else None

    temp_dir = os.path.join(app_dir, "_update_tmp")

    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    os.makedirs(temp_dir)

    if pid:
        wait_for_app_to_close(pid)
    else:
        time.sleep(3)

    extract_zip(zip_path, temp_dir)

    inner = os.listdir(temp_dir)[0]
    inner_path = os.path.join(temp_dir, inner)

    replace_files(inner_path, app_dir)

    shutil.rmtree(temp_dir)
    os.remove(zip_path)

    exe_path = os.path.join(app_dir, "Suzuki Parts Manager.exe")

    log("Nieuwe versie starten...")
    os.startfile(exe_path)

if __name__ == "__main__":
    main()