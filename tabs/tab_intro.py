import re
from pathlib import Path
from datetime import datetime, date

import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFileDialog, QMessageBox, QFrame
)
from utils.theme import apply_theme, is_dark_mode
from version import APP_VERSION, GITHUB_OWNER, GITHUB_REPO
from services.update_checker import check_github_release
from services.auto_updater import run_updater

def extract_export_date_from_filename(path: str):
    """
    Verwacht iets als:
    wc-product-export-27-2-2026-1772200987668.csv
    """
    name = Path(path).stem
    m = re.search(r"(\d{1,2})-(\d{1,2})-(\d{4})", name)
    if not m:
        return None

    day, month, year = map(int, m.groups())
    try:
        return date(year, month, day)
    except ValueError:
        return None


def fmt_date(d: date) -> str:
    return d.strftime("%d-%m-%Y")


def make_card(title: str, text: str) -> QFrame:
    card = QFrame()
    card.setStyleSheet("""
        QFrame {
            border: 1px solid palette(mid);
            border-radius: 10px;
            background: palette(base);
            padding: 10px;
        }
        QLabel[role="card_title"] {
            font-size: 13px;
            font-weight: bold;
            color: palette(text);
        }
        QLabel[role="card_text"] {
            font-size: 12px;
            color: palette(text);
        }
    """)
    lay = QVBoxLayout(card)
    lay.setContentsMargins(12, 12, 12, 12)
    lay.setSpacing(6)

    lbl_title = QLabel(title)
    lbl_title.setProperty("role", "card_title")

    lbl_text = QLabel(text)
    lbl_text.setProperty("role", "card_text")
    lbl_text.setWordWrap(True)

    lay.addWidget(lbl_title)
    lay.addWidget(lbl_text)
    return card


class TabIntro(QWidget):
    def __init__(self, app_state, on_loaded_callback=None):
        super().__init__()
        self.app_state = app_state
        self.on_loaded_callback = on_loaded_callback
        self._build_ui()

    def _build_ui(self):
        apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # =========================
        # TITEL
        # =========================
        title = QLabel("Start — WooCommerce export laden")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        root.addWidget(title)

        # =========================
        # UITLEG
        # =========================
        uitleg = QLabel(
            "Laad eerst de WooCommerce product-export (CSV).\n\n"
            "Deze export is de basis voor zoeken, inboeken, website-acties en TLC.\n"
            "Gebruik bij voorkeur de export van vandaag. Een oudere export mag ook, "
            "maar dan krijg je een waarschuwing."
        )
        uitleg.setWordWrap(True)
        root.addWidget(uitleg)

        # =========================
        # STAPPEN / INFO BLOK
        # =========================
        stappen = QLabel(
            "Werkwijze:\n"
            "① Maak of download een WooCommerce CSV export\n"
            "② Klik op 'Upload WooCommerce CSV export'\n"
            "③ Controleer bestand, datum en aantallen\n"
            "④ Daarna worden de andere tabs bruikbaar"
        )
        stappen.setWordWrap(True)
        stappen.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 10px;
            }
        """)
        root.addWidget(stappen)

        # =========================
        # ACTIE
        # =========================
        self.btn_upload = QPushButton("Laad WooCommerce export")
        self.btn_upload.setObjectName("primary")
        self.btn_upload.setMinimumHeight(40)
        self.btn_upload.clicked.connect(self.load_wc_export)
        root.addWidget(self.btn_upload)

        # =========================
        # HERLAAD KNOP
        # =========================
        self.btn_reload = QPushButton("Andere export laden")
        self.btn_reload.setObjectName("secondary")
        self.btn_reload.setMinimumHeight(40)
        self.btn_reload.clicked.connect(self.load_wc_export)
        root.addWidget(self.btn_reload)
     

        # =========================
        # STATUS BLOK
        # =========================
        self.status = QLabel("Nog geen WooCommerce export geladen")
        self.status.setWordWrap(True)
        self.status.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 10px;
                font-weight: bold;
            }
        """)
        root.addWidget(self.status)

        # =========================
        # BESTANDSINFO
        # =========================
        self.file_info = QLabel("")
        self.file_info.setWordWrap(True)
        self.file_info.hide()
        self.file_info.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 10px;
            }
        """)
        root.addWidget(self.file_info)

        # =========================
        # WAARSCHUWING
        # =========================
        self.export_warning = QLabel("")
        self.export_warning.setWordWrap(True)
        self.export_warning.hide()
        _amber = "#fcd34d" if is_dark_mode(self) else "#b06a00"
        self.export_warning.setStyleSheet(f"""
            QLabel {{
                background: rgba(255, 170, 0, 0.14);
                border: 1px solid rgba(255, 170, 0, 0.35);
                border-radius: 8px;
                padding: 10px;
                color: {_amber};
                font-weight: bold;
            }}
        """)
        root.addWidget(self.export_warning)

        root.addStretch()

        # =========================
        # VERSIE + UPDATE KNOP
        # =========================
        versie_rij = QHBoxLayout()
        versie_rij.setSpacing(12)

        self.lbl_versie = QLabel(f"Versie {APP_VERSION}")
        self.lbl_versie.setStyleSheet("color: palette(mid); font-size: 11px;")
        versie_rij.addWidget(self.lbl_versie)

        self.btn_update = QPushButton("Controleer op updates")
        self.btn_update.setObjectName("secondary")
        self.btn_update.setFixedHeight(28)
        self.btn_update.clicked.connect(self._controleer_updates)
        versie_rij.addWidget(self.btn_update)
        versie_rij.addStretch()

        root.addLayout(versie_rij)

    def _controleer_updates(self):
        self.btn_update.setEnabled(False)
        self.btn_update.setText("Bezig met controleren...")

        info = check_github_release(
            current_version=APP_VERSION,
            github_owner=GITHUB_OWNER,
            github_repo=GITHUB_REPO,
            timeout_seconds=8,
        )

        self.btn_update.setEnabled(True)
        self.btn_update.setText("Controleer op updates")

        if info.error:
            QMessageBox.warning(
                self,
                "Kan niet controleren",
                f"Kon niet controleren op updates.\n\nFout: {info.error}\n\n"
                "Controleer je internetverbinding en probeer het opnieuw."
            )
            return

        if not info.update_available:
            QMessageBox.information(
                self,
                "Je hebt de nieuwste versie",
                f"Je hebt al de nieuwste versie ({APP_VERSION}).\nEr zijn geen updates beschikbaar."
            )
            return

        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Update beschikbaar")
        msg.setText(
            f"Er is een nieuwe versie beschikbaar!\n\n"
            f"Huidige versie:  {info.current_version}\n"
            f"Nieuwe versie:   {info.latest_version}"
        )
        msg.setInformativeText("Wil je de update nu downloaden en installeren?")
        btn_ja = msg.addButton("Ja, update nu", QMessageBox.AcceptRole)
        msg.addButton("Nee, later", QMessageBox.RejectRole)
        msg.exec()

        if msg.clickedButton() == btn_ja:
            target = info.download_url or info.release_url
            if target:
                run_updater(target)

    def load_wc_export(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecteer WooCommerce product-export",
            "",
            "CSV bestanden (*.csv)"
        )

        if not path:
            return

        try:
            df = pd.read_csv(path, dtype=str, low_memory=False)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Fout",
                f"Kan WooCommerce CSV niet laden:\n{e}"
            )
            return

        export_date = extract_export_date_from_filename(path)
        today = date.today()

        self.app_state.set_wc_export(df, path)

        n_rows = int(len(df))

        stock_col = None
        for col in ("Voorraad", "Stock", "stock_quantity"):
            if col in df.columns:
                stock_col = col
                break

        in_stock = None
        if stock_col:
            try:
                s = pd.to_numeric(df[stock_col], errors="coerce").fillna(0)
                in_stock = int((s > 0).sum())
            except Exception:
                in_stock = None

        if export_date and export_date == today:
            self.status.setText("✅ WooCommerce export geladen — export van vandaag")
            self.export_warning.hide()
        elif export_date:
            self.status.setText("⚠️ WooCommerce export geladen — oudere export")
            self.export_warning.setText(
                f"Waarschuwing: dit bestand lijkt niet van vandaag te zijn.\n"
                f"Exportdatum: {fmt_date(export_date)}\n"
                f"Vandaag: {fmt_date(today)}\n\n"
                "Je mag wel doorgaan. Dit is handig voor controle en foutonderzoek, "
                "maar let op dat collega’s weten dat er met een oude export gewerkt wordt."
            )
            self.export_warning.show()
        else:
            self.status.setText("⚠️ WooCommerce export geladen — datum niet herkend")
            self.export_warning.setText(
                "Waarschuwing: de datum kon niet uit de bestandsnaam worden gelezen.\n"
                "Controleer of je wel de juiste WooCommerce export hebt gekozen."
            )
            self.export_warning.show()

        info_lines = [
            f"Bestand: {Path(path).name}",
            f"Pad: {path}",
            f"Aantal producten: {n_rows:,}".replace(",", "."),
            f"Ingeladen op: {datetime.now().strftime('%d-%m-%Y %H:%M')}",
        ]

        if export_date:
            info_lines.append(f"Exportdatum uit bestandsnaam: {fmt_date(export_date)}")
        else:
            info_lines.append("Exportdatum uit bestandsnaam: niet gevonden")

        if in_stock is not None:
            info_lines.append(f"Op voorraad (>0): {in_stock:,}".replace(",", "."))

        self.file_info.setText("\n".join(info_lines))
        self.file_info.show()

        QMessageBox.information(
            self,
            "Succes",
            "WooCommerce export succesvol geladen.\n"
            "Alle tabs kunnen nu gebruikt worden."
        )

        if self.on_loaded_callback:
            self.on_loaded_callback()