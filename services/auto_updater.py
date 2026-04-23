import urllib.request
import os
import subprocess
import sys

def download_file(url, dest):
    urllib.request.urlretrieve(url, dest)

def run_updater(download_url):
    app_dir = os.path.dirname(sys.executable)

    local_app_data = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    download_dir = os.path.join(local_app_data, "Suzuki Parts Manager")
    os.makedirs(download_dir, exist_ok=True)

    zip_path = os.path.join(download_dir, "update.zip")
    updater_path = os.path.join(app_dir, "updater.exe")

    download_file(download_url, zip_path)

    subprocess.Popen([
        updater_path,
        zip_path,
        app_dir,
        str(os.getpid()),
    ])

    sys.exit(0)