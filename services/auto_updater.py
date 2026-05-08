import urllib.request
import os
import subprocess
import sys
import zipfile
import platform


def download_file(url, dest):
    urllib.request.urlretrieve(url, dest)


def run_updater(download_url):
    app_dir = os.path.dirname(sys.executable)
    is_windows = platform.system() == "Windows"

    if is_windows:
        local_app_data = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    else:
        local_app_data = os.path.expanduser("~/Library/Application Support")

    download_dir = os.path.join(local_app_data, "Suzuki Parts Manager")
    os.makedirs(download_dir, exist_ok=True)

    zip_path = os.path.join(download_dir, "update.zip")

    download_file(download_url, zip_path)

    if is_windows:
        # Gebruik de NIEUWE updater.exe uit de gedownloade zip.
        # Dit voorkomt dat een verouderde bundeled updater.exe draait bij updates.
        updater_path = os.path.join(app_dir, "updater.exe")  # fallback: huidige versie
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                # updater.exe zit in de zip op: <map>/updater.exe
                entries = [n for n in zf.namelist() if n.endswith("/updater.exe") or n == "updater.exe"]
                if entries:
                    new_updater_path = os.path.join(download_dir, "updater_new.exe")
                    with zf.open(entries[0]) as src, open(new_updater_path, "wb") as dst:
                        dst.write(src.read())
                    updater_path = new_updater_path
        except Exception:
            pass  # bij fout: val terug op de bundeled updater.exe

        subprocess.Popen([
            updater_path,
            zip_path,
            app_dir,
            str(os.getpid()),
        ])
    else:
        updater_script = os.path.join(os.path.dirname(os.path.dirname(__file__)), "updater.py")
        subprocess.Popen([
            sys.executable,
            updater_script,
            zip_path,
            app_dir,
            str(os.getpid()),
        ])

    sys.exit(0)
