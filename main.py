import sys
import ctypes

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget,
    QWidget, QVBoxLayout, QPushButton,
    QLabel, QFileDialog, QMessageBox, QSplashScreen
)
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import QLockFile, Qt

import pandas as pd

from app_state import AppState
from tabs.tab_inboeken import TabInboeken
from tabs.tab_tradelist import TabTradelist
from tabs.tab_tlc_1322 import TabTLC1322
from tabs.tab_tlc_update import TabTLCUpdate
from tabs.tab_website_277 import TabWebsite277
from tabs.tab_factuurmaker import TabFactuurmaker
from tabs.tab_zoeklijst import TabZoeklijst
from tabs.tab_intro import TabIntro

from services.update_checker import check_github_release
from services.auto_updater import run_updater
from version import APP_VERSION, GITHUB_OWNER, GITHUB_REPO
from utils.paths import resource_path, output_root, appdata_root, get_lock_file
from tabs.tab_intro import TabIntro
from utils.theme import apply_theme

def create_single_instance_lock():
    lock = QLockFile(str(get_lock_file()))
    lock.setStaleLockTime(0)

    if not lock.tryLock(0):
        QMessageBox.warning(None, "App al geopend", "Suzuki Parts Manager is al geopend.")
        return None

    return lock


def create_splash(app: QApplication) -> QSplashScreen:
    splash_path = resource_path("assets/splash.png")

    pixmap = QPixmap(str(splash_path))
    if pixmap.isNull():
        pixmap = QPixmap(700, 380)
        pixmap.fill(Qt.white)

    splash = QSplashScreen(pixmap)
    splash.setWindowFlag(Qt.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()
    return splash


def splash_message(app: QApplication, splash: QSplashScreen, text: str):
    splash.showMessage(
        text,
        Qt.AlignBottom | Qt.AlignHCenter,
        Qt.black
    )
    app.processEvents()


def maybe_check_for_updates(parent=None):
    info = check_github_release(
        current_version=APP_VERSION,
        github_owner=GITHUB_OWNER,
        github_repo=GITHUB_REPO,
        timeout_seconds=3,
    )

    if info.error:
        return

    if not info.update_available:
        return

    msg = QMessageBox(parent)
    msg.setIcon(QMessageBox.Information)
    msg.setWindowTitle("Update beschikbaar")
    msg.setText(
        f"Er is een nieuwe versie beschikbaar.\n\n"
        f"Huidige versie: {info.current_version}\n"
        f"Nieuwe versie: {info.latest_version}"
    )

    if info.asset_name:
        msg.setInformativeText(f"Downloadbestand: {info.asset_name}")
    else:
        msg.setInformativeText("Er is een nieuwe release beschikbaar.")

    btn_download = msg.addButton("Download", QMessageBox.AcceptRole)
    msg.addButton("Later", QMessageBox.RejectRole)

    msg.exec()

    if msg.clickedButton() == btn_download:
        target = info.download_url or info.release_url
        if target:
            run_updater(target)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.app_state = AppState()

        self.setWindowTitle(f"Suzuki Parts Manager v{APP_VERSION}")
        self.setWindowIcon(QIcon(str(resource_path("assets/app_icon.png"))))
        self.resize(1100, 750)

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        apply_theme(self)

        # ✅ JUISTE Intro tab
        self.intro_tab = TabIntro(self.app_state, self.create_work_tabs)
        self.tabs.addTab(self.intro_tab, "Start")


    def create_work_tabs(self):
        self.tabs.clear()
        self.tabs.addTab(self.intro_tab, "Start")

        self.tab_inboeken = TabInboeken(self.app_state)
        self.tab_tradelist = TabTradelist(self.app_state)
        self.tab_1322 = TabTLC1322(self.app_state)
        self.tab_tlc_update = TabTLCUpdate(self.app_state)
        self.tab_website = TabWebsite277(self.app_state)
        self.tab_factuur = TabFactuurmaker(self.app_state)
        self.tab_zoeklijst = TabZoeklijst(self.app_state)

        self.tabs.addTab(self.tab_inboeken, "Inboeken")
        self.tabs.addTab(self.tab_tradelist, "Tradelist")
        self.tabs.addTab(self.tab_1322, "1322")
        self.tabs.addTab(self.tab_tlc_update, "TLC Update")
        self.tabs.addTab(self.tab_website, "Website 277")
        self.tabs.addTab(self.tab_factuur, "Factuurmaker")
        self.tabs.addTab(self.tab_zoeklijst, "Zoeklijst")


if __name__ == "__main__":
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
        "classic.suzuki.parts.manager"
    )

    app = QApplication(sys.argv)

    # 🔴 EERST: check of app al open is
    lock = create_single_instance_lock()
    if lock is None:
        sys.exit(0)

    # ✅ DAN PAS splash
    splash = create_splash(app)
    splash_message(app, splash, "App starten...")

    splash_message(app, splash, "Lokale mappen controleren...")
    appdata_root()
    output_root()

    splash_message(app, splash, "Hoofdvenster laden...")
    window = MainWindow()
    window.showMaximized()
    app.processEvents()

    splash_message(app, splash, "Controleren op updates...")
    maybe_check_for_updates(window)

    splash_message(app, splash, "Klaar met opstarten...")
    splash.finish(window)

    sys.exit(app.exec())