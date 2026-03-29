import urllib.request
import os
import subprocess
import sys

def download_file(url, dest):
    urllib.request.urlretrieve(url, dest)

def run_updater(download_url):
    app_dir = os.path.dirname(sys.executable)

    zip_path = os.path.join(app_dir, "update.zip")
    updater_path = os.path.join(app_dir, "updater.exe")

    download_file(download_url, zip_path)

    subprocess.Popen([
        updater_path,
        zip_path,
        app_dir
    ])

    sys.exit(0)