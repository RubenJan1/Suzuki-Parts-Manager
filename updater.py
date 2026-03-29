import os
import sys
import time
import shutil
import zipfile

def log(msg):
    print("[UPDATER]", msg)

def wait_for_app_to_close(app_path):
    log("Wachten tot app sluit...")
    time.sleep(2)

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
        log("FOUT: geen args")
        return

    zip_path = sys.argv[1]
    app_dir = sys.argv[2]

    temp_dir = os.path.join(app_dir, "_update_tmp")

    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    os.makedirs(temp_dir)

    wait_for_app_to_close(app_dir)

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