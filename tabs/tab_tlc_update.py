# ============================================================
# tabs/tab_tlc_update.py
# TLC Update UI – 1 scherm, minimale handelingen
# ============================================================

import os
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog,
    QMessageBox, QTextEdit, QRadioButton,
    QLineEdit
)

from engines.engine_tlc_update import TLCUpdateEngine

from utils.paths import output_root

BASE_DIR = str(output_root() / "1322")


class TabTLCUpdate(QWidget):
    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state
        self.engine = TLCUpdateEngine()
        self.engine.set_base_dir(BASE_DIR)
        self._last_debug_dir = None

        self._build_ui()
        self._sync_status()

    def _active_tlc_path(self):
        return os.path.join(BASE_DIR, "TLC", "TLC_1.xlsx")

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)

        title = QLabel("TLC Update – Aanvullen / Krat vervangen")
        title.setStyleSheet("font-size: 20px; font-weight: 800;")
        root.addWidget(title)

        uitleg = QLabel(
            "Updatebestanden formaat (geen kolomnamen):\n"
            "• Kolom A = Title/Artikelnummer\n"
            "• Kolom B = Voorraad\n"
            "• Kolom C = Locatie\n"
            "• Kolom D = Prijs\n\n"
            "Regel: nooit samenvoegen op Title. Alleen match op Title+Locatie."
        )
        uitleg.setWordWrap(True)
        uitleg.setStyleSheet("opacity: 0.9;")
        root.addWidget(uitleg)

        # Status card
        self.lbl_status = QLabel("")
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setStyleSheet("""
            QLabel {
                background-color: palette(base);
                color: palette(text);
                border: 1px solid palette(mid);
                padding: 10px;
                border-radius: 10px;
            }
        """)
        root.addWidget(self.lbl_status)

        # Buttons row
        row_btn = QHBoxLayout()
        self.btn_open_tlc = QPushButton("Open TLC map")
        self.btn_open_tlc.clicked.connect(self.on_open_tlc)

        self.btn_add = QPushButton("Voeg updatebestand(en) toe")
        self.btn_add.clicked.connect(self.on_add_files)

        self.btn_clear = QPushButton("Reset selectie")
        self.btn_clear.clicked.connect(self.on_reset_files)

        row_btn.addWidget(self.btn_open_tlc)
        row_btn.addWidget(self.btn_add)
        row_btn.addWidget(self.btn_clear)
        row_btn.addStretch(1)
        root.addLayout(row_btn)

        # Mode
        # Mode (normaal)
        mode_box = QHBoxLayout()
        self.rb_merge = QRadioButton("Aanvullen / bijwerken")
        self.rb_merge.setChecked(True)
        mode_box.addWidget(self.rb_merge)
        mode_box.addStretch(1)
        root.addLayout(mode_box)

        # --- Gevorderd / correcties ---
        lbl_adv = QLabel("⚠️ Correcties (gevorderd)")
        lbl_adv.setStyleSheet("font-weight: 700; margin-top: 8px;")
        root.addWidget(lbl_adv)

        lbl_adv_info = QLabel(
            "Gebruik dit alleen bij fysieke herindeling of historische fouten.\n"
            "Deze actie vervangt een hele locatie (krat) en maakt altijd een backup + report."
        )
        lbl_adv_info.setWordWrap(True)
        lbl_adv_info.setStyleSheet("color: #666;")
        root.addWidget(lbl_adv_info)

        adv_box = QHBoxLayout()
        self.rb_replace = QRadioButton("Krat vervangen (locatie)")
        self.txt_loc = QLineEdit()
        self.txt_loc.setPlaceholderText("Bijv: 63 of 63,64")
        self.txt_loc.setFixedWidth(160)
        self.txt_loc.setEnabled(False)

        def on_mode_change():
            self.txt_loc.setEnabled(self.rb_replace.isChecked())

        self.rb_merge.toggled.connect(on_mode_change)
        self.rb_replace.toggled.connect(on_mode_change)

        adv_box.addWidget(self.rb_replace)
        adv_box.addWidget(self.txt_loc)
        adv_box.addStretch(1)
        root.addLayout(adv_box)

        # Run row
        row_run = QHBoxLayout()
        self.btn_run = QPushButton("START TLC UPDATE")
        self.btn_run.setFixedHeight(44)
        self.btn_run.setStyleSheet("font-weight: 800;")
        self.btn_run.clicked.connect(self.on_run_update)

        self.btn_open_debug = QPushButton("Open laatste debug")
        self.btn_open_debug.clicked.connect(self.on_open_last_debug)

        row_run.addWidget(self.btn_run)
        row_run.addWidget(self.btn_open_debug)
        row_run.addStretch(1)
        root.addLayout(row_run)

        # Log
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(200)
        self.log.setPlaceholderText("Log verschijnt hier...")
        root.addWidget(self.log, stretch=1)

    def _sync_status(self):
        p = self._active_tlc_path()
        if os.path.exists(p):
            mtime = os.path.getmtime(p)
            laatst = datetime.fromtimestamp(mtime).strftime("%d-%m-%Y %H:%M:%S")
            self.lbl_status.setText(
                "✅ Actieve TLC gevonden (wordt automatisch gebruikt)\n"
                f"Pad: {p}\n"
                f"Laatst gewijzigd: {laatst}\n"
                f"Updatebestanden geselecteerd: {len(self.engine.update_paths)}"
            )
        else:
            self.lbl_status.setText(
                "❌ Actieve TLC ontbreekt\n"
                f"Verwacht: {p}\n\n"
                "Klik 'Open TLC map' en plaats daar TLC_1.xlsx."
            )
        self.btn_run.setEnabled(os.path.exists(p))
        self.btn_run.setEnabled(os.path.exists(self._active_tlc_path()))


    def on_open_tlc(self):
        folder = os.path.join(BASE_DIR, "TLC")
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)
        self._sync_status()

    def on_add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecteer TLC updatebestand(en)",
            "",
            "Excel bestanden (*.xlsx)"
        )
        if not paths:
            return
        for p in paths:
            self.engine.add_update_file(p)

        self.log.append(f"✅ {len(paths)} updatebestand(en) toegevoegd.")
        self._sync_status()

    def on_reset_files(self):
        self.engine.clear_updates()
        self.log.clear()
        self._sync_status()

    def on_open_last_debug(self):
        if getattr(self, "_last_debug_dir", None) and os.path.exists(self._last_debug_dir):
            os.startfile(self._last_debug_dir)
            return

        QMessageBox.information(
            self,
            "Info",
            "Nog geen debug map gevonden.\n"
            "Tip: run eerst een update (ook als die faalt door conflicts, wordt debug dan opgeslagen)."
        )


    def on_run_update(self):
        try:
            self.engine.set_base_dir(BASE_DIR)

            mode = "MERGE"
            locs = None
            if self.rb_replace.isChecked():
                mode = "REPLACE_LOC"
                locs = self.txt_loc.text().strip()
            
            if mode == "REPLACE_LOC":
                reply = QMessageBox.warning(
                    self,
                    "⚠️ Krat vervangen (locatie)",
                    "Je staat op het punt om een hele locatie (krat) te vervangen.\n\n"
                    "Gebruik dit ALLEEN als:\n"
                    "• de krat opnieuw is gedaan / fysieke herindeling\n"
                    "• of de locatie historisch fout was\n\n"
                    "Er wordt altijd een backup + report gemaakt.\n\n"
                    "Weet je zeker dat je wilt doorgaan?",
                    QMessageBox.StandardButton.Cancel | QMessageBox.StandardButton.Ok
                )
                if reply != QMessageBox.StandardButton.Ok:
                    self.log.append("⏭️ Geannuleerd door gebruiker.")
                    return


            self.log.append("---- TLC UPDATE START ----")
            self.log.append(f"Mode: {mode}" + (f" | Locatie(s): {locs}" if locs else ""))

            res = self.engine.run_update(mode=mode, replace_locations=locs)

        except Exception as e:
            msg = str(e)

            # Probeer report-pad uit de foutmelding te halen
            report_path = None
            for line in msg.splitlines():
                line = line.strip()
                if line.lower().endswith(".xlsx") and "TLC_UPDATE_REPORT_" in line:
                    report_path = line
                    break
                # jouw fout heeft soms "Zie report:" op een regel en pad op volgende regel
                if line.lower().startswith("c:\\") and line.lower().endswith(".xlsx"):
                    report_path = line
                    break

            # Als report gevonden: zet last debug dir
            if report_path and os.path.exists(report_path):
                self._last_debug_dir = os.path.dirname(report_path)
                self.log.append(f"📄 Report gevonden: {report_path}")
                self.log.append(f"📁 Debug map: {self._last_debug_dir}")

            QMessageBox.critical(self, "Fout", msg)
            self.log.append(f"❌ Fout: {msg}")
            self._sync_status()
            return


        self._last_debug_dir = res.get("debug_dir")
        self.log.append("✅ Update klaar!")
        self.log.append(f"- TLC bijgewerkt: {res.get('active_tlc')}")
        self.log.append(f"- Backup gemaakt: {res.get('backup_tlc')}")
        self.log.append(f"- Report: {res.get('report')}")
        self.log.append(
            f"- Bijgewerkt: {res.get('updated')} | Nieuw: {res.get('added')} | Verwijderd: {res.get('removed')} | Warnings: {res.get('warnings')}"
        )

        QMessageBox.information(
            self,
            "TLC update klaar",
            "TLC update is succesvol uitgevoerd.\n\n"
            "Er is altijd een backup gemaakt en een report in DEBUG gezet."
        )

        self._sync_status()
