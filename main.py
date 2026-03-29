# ============================================================
# Suzuki Parts – Centrale Desktop App
# main.py
# ============================================================
from utils.theme import DARK_THEME, LIGHT_THEME
from PySide6.QtGui import QGuiApplication

import sys
import pandas as pd
from PySide6.QtGui import QIcon
import ctypes
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget,
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QFrame, QSizePolicy
)
from PySide6.QtCore import Qt, QLockFile

from app_state import AppState
from utils.paths import resource_path, output_root, appdata_root, get_lock_file

# Tabs
from tabs.tab_inboeken import TabInboeken
from tabs.tab_website_277 import TabWebsite277
from tabs.tab_tlc_1322 import TabTLC1322
from tabs.tab_tlc_update import TabTLCUpdate
from tabs.tab_tradelist import TabTradelist
from tabs.tab_factuurmaker import TabFactuurmaker
from tabs.tab_zoeklijst import TabZoeklijst

# ============================================================
# INTRO TAB – WC EXPORT UPLOAD
# ============================================================
def system_is_dark():
    palette = QGuiApplication.palette()
    return palette.window().color().lightness() < 128

def create_single_instance_lock():
    lock = QLockFile(str(get_lock_file()))
    lock.setStaleLockTime(0)
    if not lock.tryLock():
        QMessageBox.warning(None, "App al geopend", "Suzuki Parts Manager is al geopend.")
        return None
    return lock

class TabIntro(QWidget):
    def __init__(self, app_state: AppState, on_loaded_callback):
        super().__init__()
        if system_is_dark():
            self.setStyleSheet(DARK_THEME)
        else:
            self.setStyleSheet(LIGHT_THEME)
        self.app_state = app_state
        self.on_loaded_callback = on_loaded_callback
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(14)

        # Header
        title = QLabel("Suzuki Parts Manager")
        title.setObjectName("IntroTitle")
        title.setAlignment(Qt.AlignLeft)

        subtitle = QLabel("Laad eerst je WooCommerce product-export (CSV). Daarna worden alle tabs ontgrendeld.")
        subtitle.setWordWrap(True)
        subtitle.setObjectName("IntroSubtitle")
        subtitle.setAlignment(Qt.AlignLeft)

        header = QVBoxLayout()
        header.setSpacing(6)
        header.addWidget(title)
        header.addWidget(subtitle)

        # Cards row
        cards_row = QHBoxLayout()
        cards_row.setSpacing(14)

        def make_card(card_title: str, card_text: str):
            frame = QFrame()
            frame.setFrameShape(QFrame.StyledPanel)
            frame.setObjectName("IntroCard")
            frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)

            v = QVBoxLayout(frame)
            v.setContentsMargins(14, 12, 14, 12)
            v.setSpacing(8)

            t = QLabel(card_title)
            t.setObjectName("IntroCardTitle")
            t.setWordWrap(True)

            body = QLabel(card_text)
            body.setObjectName("IntroCardBody")
            body.setWordWrap(True)

            v.addWidget(t)
            v.addWidget(body)
            v.addStretch(1)
            return frame

        card1 = make_card(
            "Stap 1 — Exporteren",
            "WooCommerce → Producten → Exporteren\n"
            "• Kies: Alle producten\n"
            "• Exporteer als CSV\n"
            "• Download het bestand"
        )
        card2 = make_card(
            "Stap 2 — Uploaden",
            "Klik hieronder op ‘Upload WooCommerce CSV export’.\n"
            "De app gebruikt dit bestand als ‘één waarheid’ en schrijft nooit direct naar WooCommerce."
        )
        card3 = make_card(
            "Tip — Snelle check",
            "Na laden zie je:\n"
            "• Bestandsnaam en pad\n"
            "• Aantal producten\n"
            "• Aantal op voorraad\n"
            "Zo weet je meteen of je de juiste export hebt."
        )

        cards_row.addWidget(card1)
        cards_row.addWidget(card2)
        cards_row.addWidget(card3)

        # Action area
        action_frame = QFrame()
        action_frame.setFrameShape(QFrame.StyledPanel)
        action_frame.setObjectName("IntroAction")

        action_layout = QHBoxLayout(action_frame)
        action_layout.setContentsMargins(14, 12, 14, 12)
        action_layout.setSpacing(12)

        self.btn_upload = QPushButton("Upload WooCommerce CSV export")
        self.btn_upload.setObjectName("PrimaryButton")
        self.btn_upload.setFixedHeight(44)
        self.btn_upload.clicked.connect(self.load_wc_export)

        self.status = QLabel("❌ Nog geen WooCommerce-export geladen")
        self.status.setWordWrap(True)
        self.status.setObjectName("IntroStatus")

        self.file_info = QLabel("")
        self.file_info.setWordWrap(True)
        self.file_info.setObjectName("IntroFileInfo")

        left = QVBoxLayout()
        left.setSpacing(6)
        left.addWidget(self.status)
        left.addWidget(self.file_info)

        action_layout.addLayout(left, stretch=1)
        action_layout.addWidget(self.btn_upload)

        # Minimal, theme-friendly styling (works with both LIGHT_THEME & DARK_THEME)
        self.setStyleSheet(self.styleSheet() + """
            QLabel#IntroTitle { font-size: 26px; font-weight: 800; }
            QLabel#IntroSubtitle { font-size: 13px; opacity: 0.9; }
            QFrame#IntroCard, QFrame#IntroAction { border-radius: 12px; }
            QLabel#IntroCardTitle { font-size: 14px; font-weight: 700; }
            QLabel#IntroCardBody { font-size: 12px; }
            QLabel#IntroStatus { font-size: 13px; font-weight: 600; }
            QLabel#IntroFileInfo { font-size: 12px; opacity: 0.9; }
            QPushButton#PrimaryButton { font-size: 13px; font-weight: 700; padding: 10px 14px; }
        """)

        root.addLayout(header)
        root.addSpacing(4)
        root.addLayout(cards_row)
        root.addSpacing(4)
        root.addWidget(action_frame)
        root.addStretch(1)



    def load_wc_export(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecteer WooCommerce export",
            "",
            "CSV bestanden (*.csv)"
        )

        if not path:
            return

        try:
            df = pd.read_csv(path, dtype=str, low_memory=False)
        except Exception as e:
            QMessageBox.critical(self, "Fout", f"Kan CSV niet laden:\n{e}")
            return

        # Centrale state
        self.app_state.set_wc_export(df, path)

        # Kleine sanity checks / info
        n_rows = int(len(df))
        stock_col = "Voorraad" if "Voorraad" in df.columns else ("Stock" if "Stock" in df.columns else None)
        in_stock = None
        if stock_col:
            try:
                s = pd.to_numeric(df[stock_col], errors="coerce").fillna(0)
                in_stock = int((s > 0).sum())
            except Exception:
                in_stock = None

        self.status.setText("✅ WooCommerce-export geladen")
        info_lines = [
            f"Bestand: {path}",
            f"Producten: {n_rows:,}".replace(",", ".")
        ]
        if in_stock is not None:
            info_lines.append(f"Op voorraad (>0): {in_stock:,}".replace(",", "."))
        self.file_info.setText("\n".join(info_lines))

        QMessageBox.information(
            self,
            "Succes",
            "WooCommerce-export succesvol geladen.\n\n"
            "De applicatie wordt nu ontgrendeld."
        )

        self.on_loaded_callback()


# ============================================================
# MAIN WINDOW
# ============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowIcon(QIcon(str(resource_path("assets/app_icon.png"))))
        self.setWindowTitle("Suzuki Parts – Afboeken & Inboeken")

        self.app_state = AppState()

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Intro tab (altijd eerst)
        self.tab_intro = TabIntro(self.app_state, self.create_work_tabs)
        self.tabs.addTab(self.tab_intro, "Start")

    
    # --------------------------------------------------------
    # Werk-tabs pas maken NA WC upload
    # --------------------------------------------------------



    def create_work_tabs(self):
        # voorkom dubbel aanmaken
        if self.tabs.count() > 1:
            return

        self.tab_inboeken = TabInboeken(self.app_state)
        self.tab_277 = TabWebsite277(self.app_state)
        self.tab_1322 = TabTLC1322(self.app_state)
        self.tab_tlc_update = TabTLCUpdate(self.app_state)
        self.tab_tradelist = TabTradelist(self.app_state)
        self.tab_factuur = TabFactuurmaker(self.app_state)
        self.tab_zoeklijst = TabZoeklijst(self.app_state)

        self.tabs.addTab(self.tab_inboeken, "Inboeken")
        self.tabs.addTab(self.tab_277, "Website 277")
        self.tabs.addTab(self.tab_1322, "TLC 1322")
        self.tabs.addTab(self.tab_tlc_update, "TLC Update")
        self.tabs.addTab(self.tab_tradelist, "Tradelist")
        self.tabs.addTab(self.tab_factuur, "Factuurmaker")
        self.tabs.addTab(self.tab_zoeklijst, "Zoeklijst")

        # automatisch naar Inboeken
        self.tabs.setCurrentIndex(1)



# ============================================================
# APP START
# ============================================================

    
if __name__ == "__main__":


    

    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
        "classic.suzuki.parts.manager"
    )

    app = QApplication(sys.argv)

    appdata_root()
    output_root()

        
    lock = create_single_instance_lock()
    if lock is None:
        sys.exit(0)

    window = MainWindow()
    window.show()
    window.showMaximized()

    sys.exit(app.exec())
