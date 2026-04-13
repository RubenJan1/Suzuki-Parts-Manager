# tabs/tab_zoeklijst.py
"""
Fix voor je crash:

Je had:
    self.btn_load_tlc_master ...
    inp_row.insertWidget(...)
BOVEN de regel:
    inp_row = QHBoxLayout()

Daardoor was inp_row nog niet aangemaakt -> UnboundLocalError.

Deze fixed versie:
- maakt inp_row eerst
- voegt knoppen daarna toe
- toont status: WC ✅/❌ en TLC ✅/❌
- tabel heeft Source kolom (WC/TLC/NO)
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional, List

from PySide6.QtCore import Qt
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTextEdit, QFileDialog, QLineEdit, QMessageBox, QTableWidget,
    QTableWidgetItem, QSizePolicy
)

from engines.engine_zoeklijst import EngineZoeklijst, extract_part_numbers_from_text, extract_part_numbers_from_xlsx
from utils.paths import output_root

class TabZoeklijst(QWidget):
    def __init__(self, app_state, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self.engine = EngineZoeklijst(output_dir=str(output_root() / "zoeklijst"))
        self._last_output_path: Optional[Path] = None
        self._parts: List[str] = []

        self._build_ui()

        if getattr(self.app_state, "wc_df", None) is not None:
            try:
                self.engine.load_website_df(self.app_state.wc_df)
            except Exception as e:
                self.lbl_status.setText(f"WC export fout: {e}")
            self._update_status()

    def _update_status(self):
        wc_ok = "✅" if getattr(self.app_state, "wc_df", None) is not None else "❌"
        tlc_ok = "✅" if self.engine.tlc_df is not None else "❌"
        self.lbl_status.setText(f"WC export {wc_ok} | TLC masterlijst {tlc_ok}")

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        top = QHBoxLayout()
        lbl_title = QLabel("Zoeklijst — Mail / XLSX parser")
        lbl_title.setStyleSheet("font-size: 18px; font-weight: 600;")
        top.addWidget(lbl_title)
        top.addStretch(1)

        self.btn_open_output = QPushButton("📁 Open output folder")
        self.btn_open_output.clicked.connect(self.open_output_folder)
        top.addWidget(self.btn_open_output)

        root.addLayout(top)

        self.lbl_status = QLabel("WC export ❌ | TLC masterlijst ❌")
        self.lbl_status.setStyleSheet("opacity: 0.85;")
        root.addWidget(self.lbl_status)

        # Input row (hier zat jouw bug)
        inp_row = QHBoxLayout()

        self.btn_load_tlc_master = QPushButton("📦 Laad TLC masterlijst (XLSX)…")
        self.btn_load_tlc_master.clicked.connect(self.on_load_tlc_master)
        inp_row.addWidget(self.btn_load_tlc_master)

        self.btn_load_xlsx = QPushButton("Upload zoeklijst (XLSX)…")
        self.btn_load_xlsx.clicked.connect(self.on_load_xlsx)
        inp_row.addWidget(self.btn_load_xlsx)

        self.btn_parse_text = QPushButton("Parse pasted text")
        self.btn_parse_text.clicked.connect(self.on_parse_text)
        inp_row.addWidget(self.btn_parse_text)

        inp_row.addStretch(1)

        self.btn_build = QPushButton("Maak XLSX report")
        self.btn_build.setStyleSheet("font-weight: 600; padding: 8px 14px;")
        self.btn_build.clicked.connect(self.on_build_report)
        inp_row.addWidget(self.btn_build)

        root.addLayout(inp_row)

        self.txt_paste = QTextEdit()
        self.txt_paste.setPlaceholderText(
            "Plak hier een lijst… (part numbers met of zonder streepjes)\n"
            "Tip: je mag de omschrijvingen laten staan; we vissen alleen de nummers eruit."
        )

        # auto-load TLC uit vaste map (hufterproof)
        tlc_path = str(output_root() / "1322" / "TLC" / "TLC_1.xlsx")
        if os.path.exists(tlc_path):
            try:
                self.engine.load_tlc_xlsx(tlc_path)
            except Exception as e:
                self.lbl_status.setText(f"TLC auto-load fout: {e}")
        self._update_status()

        self.txt_paste.setMinimumHeight(200)
        root.addWidget(self.txt_paste)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["Part number", "Source", "Found", "Stock", "Price", "Locatie", "Notes"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        root.addWidget(self.table, stretch=1)

        bottom = QHBoxLayout()
        bottom.addWidget(QLabel("Output:"))
        self.ed_filename = QLineEdit("zoeklijst_report.xlsx")
        self.ed_filename.setPlaceholderText("Bestandsnaam output")
        bottom.addWidget(self.ed_filename, stretch=1)

        self.btn_open_last = QPushButton("Open laatste output")
        self.btn_open_last.clicked.connect(self.open_last_output)
        bottom.addWidget(self.btn_open_last)

        root.addLayout(bottom)

    def _ensure_wc_loaded(self) -> bool:
        if getattr(self.app_state, "wc_df", None) is None:
            QMessageBox.warning(self, "WC export ontbreekt", "Laad eerst de WooCommerce CSV via de Start-tab.")
            return False
        if self.engine.website_df is None:
            try:
                self.engine.load_website_df(self.app_state.wc_df)
            except Exception as e:
                QMessageBox.critical(self, "WC export fout", str(e))
                return False
        return True

    def on_load_tlc_master(self):
        path, _ = QFileDialog.getOpenFileName(self, "Kies TLC masterlijst", "", "Excel (*.xlsx *.xls)")
        if not path:
            return
        try:
            self.engine.load_tlc_xlsx(path)
            self._update_status()
            QMessageBox.information(self, "Geladen", f"TLC masterlijst geladen:\n{os.path.basename(path)}")
        except Exception as e:
            QMessageBox.critical(self, "TLC masterlijst fout", str(e))

    def on_load_xlsx(self):
        if not self._ensure_wc_loaded():
            return
        path, _ = QFileDialog.getOpenFileName(self, "Kies zoeklijst", "", "Excel (*.xlsx *.xls)")
        if not path:
            return
        try:
            parts = extract_part_numbers_from_xlsx(path)
            self._parts = parts
            self._preview(parts, source=f"XLSX_toggle: {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.critical(self, "XLSX fout", str(e))

    def on_parse_text(self):
        if not self._ensure_wc_loaded():
            return
        text = self.txt_paste.toPlainText()
        parts = extract_part_numbers_from_text(text)
        self._parts = parts
        self._preview(parts, source="Pasted text")

    def _preview(self, parts: List[str], source: str):
        if not parts:
            self.lbl_status.setText(f"{source} — geen part numbers gevonden ❌")
            self.table.setRowCount(0)
            return
        self._update_status()
        report = self.engine.build_report(parts)
        self._fill_table(report)

    def _fill_table(self, df):
        self.table.setRowCount(0)
        cols = ["Part number", "Source", "Found", "Stock", "Price", "Locatie", "Notes"]
        for _, row in df.iterrows():
            r = self.table.rowCount()
            self.table.insertRow(r)
            for c, col in enumerate(cols):
                item = QTableWidgetItem(str(row.get(col, "")))
                if col == "Found" and str(row.get(col, "")) == "NO":
                    item.setForeground(Qt.red)
                self.table.setItem(r, c, item)

    def on_build_report(self):
        if not self._ensure_wc_loaded():
            return
        if not self._parts:
            QMessageBox.information(self, "Geen input", "Upload een XLSX of plak text en klik 'Parse'.")
            return

        filename = (self.ed_filename.text() or "").strip() or "zoeklijst_report.xlsx"
        if not filename.lower().endswith(".xlsx"):
            filename += ".xlsx"

        report = self.engine.build_report(self._parts)
        try:
            out_path = self.engine.export_report_xlsx(report, filename=filename)
            self._last_output_path = out_path
            QMessageBox.information(self, "Klaar", f"Report opgeslagen:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Opslaan fout", str(e))

    def open_output_folder(self):
        out_dir = Path(self.engine.output_dir).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(out_dir.as_uri())

    def open_last_output(self):
        if self._last_output_path and self._last_output_path.exists():
            QDesktopServices.openUrl(self._last_output_path.resolve().as_uri())
        else:
            QMessageBox.information(self, "Geen output", "Er is nog geen output gemaakt.")
