from __future__ import annotations

import os
import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit,
    QRadioButton, QComboBox, QDoubleSpinBox,
    QTableWidget, QHeaderView, QSizePolicy,
    QGroupBox, QButtonGroup, QTableWidgetItem, QMessageBox, QFileDialog,
    QScrollArea, QFrame,
)

from engines.engine_factuurmaker import FactuurMakerEngine
from utils.paths import output_root, resource_path
from utils.theme import apply_theme
from PySide6.QtGui import QColor

def currency(value) -> str:
    """Format euro with European separators."""
    try:
        v = float(value)
        return f"€ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "€ 0,00"



class TabFactuurmaker(QWidget):
    """
    Factuurmaker (Factuur + Creditfactuur in één tab)

    PRO v3 UI fixes:
    - Preview krijgt standaard veel meer ruimte op 14–15,6"
    - Type toggle + Actions in één rij (minder “lege lucht”)
    - Slimme kolombreedtes (minder horizontale scrollbar)
    - Splitter handle breder (makkelijk slepen)
    """

    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state
        self.engine = FactuurMakerEngine()
        self._loaded_bron: str | None = None
        self._build_ui()
        self._validate_form()

    def showEvent(self, event):
        super().showEvent(event)
        self._refresh_banner()

    def _maak_cms_banner(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("cms_banner")
        frame.setStyleSheet("""
            QFrame#cms_banner {
                background: rgba(234,88,12,0.12);
                border: 2px solid #ea580c;
                border-radius: 10px;
            }
        """)
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(10)

        self.banner_lbl = QLabel()
        self.banner_lbl.setWordWrap(True)
        self.banner_lbl.setStyleSheet(
            "font-size: 13px; font-weight: bold; color: #9a3412;"
            "background: transparent; border: none;"
        )
        lay.addWidget(self.banner_lbl)

        rij = QHBoxLayout()
        self.btn_laad_277 = QPushButton("Laad 277-orders in factuur")
        self.btn_laad_277.setObjectName("primary")
        self.btn_laad_277.setMinimumHeight(40)
        self.btn_laad_277.clicked.connect(lambda: self._laad_uit_queue("277"))

        self.btn_laad_1322 = QPushButton("Laad 1322-orders in factuur")
        self.btn_laad_1322.setObjectName("primary")
        self.btn_laad_1322.setMinimumHeight(40)
        self.btn_laad_1322.clicked.connect(lambda: self._laad_uit_queue("1322"))

        rij.addWidget(self.btn_laad_277)
        rij.addWidget(self.btn_laad_1322)
        rij.addStretch()
        lay.addLayout(rij)

        frame.hide()
        return frame

    def _refresh_banner(self):
        try:
            from services.cms_queue import pending_counts, pending_factuurnummers, has_pending
            counts = pending_counts()
            n277   = counts.get("277", 0)
            n1322  = counts.get("1322", 0)

            if n277 == 0 and n1322 == 0:
                self.banner_cms.hide()
                return

            delen = []
            if n277 > 0:
                fnrs = pending_factuurnummers("277")
                fnr_str = ", ".join(fnrs[:5]) + (" ..." if len(fnrs) > 5 else "")
                delen.append(f"277: {n277} regel(s)  —  ordernr: {fnr_str}")
            if n1322 > 0:
                fnrs = pending_factuurnummers("1322")
                fnr_str = ", ".join(fnrs[:5]) + (" ..." if len(fnrs) > 5 else "")
                delen.append(f"1322: {n1322} regel(s)  —  ordernr: {fnr_str}")

            self.banner_lbl.setText(
                "⚠️  Er staan CMS-orders klaar voor facturering:\n\n"
                + "\n".join(delen)
                + "\n\nKlik hieronder op de juiste knop om de orders te laden."
            )
            self.btn_laad_277.setEnabled(n277 > 0)
            self.btn_laad_1322.setEnabled(n1322 > 0)
            self.banner_cms.show()
        except Exception:
            self.banner_cms.hide()

    def _laad_uit_queue(self, bron: str):
        try:
            from services.cms_queue import get_pending
            import pandas as pd
            entries = get_pending(bron)
            if not entries:
                return

            regels = []
            for entry in entries:
                for r in entry.get("regels", []):
                    regels.append({
                        "Artikel":      r.get("title", ""),
                        "Omschrijving": r.get("omschrijving", ""),
                        "Aantal":       int(r.get("geleverd", 0)),
                        "Prijs":        float(r.get("prijs", 0.0)),
                    })

            if not regels:
                return

            df = pd.DataFrame(regels)
            df["Aantal"] = pd.to_numeric(df["Aantal"], errors="coerce").fillna(0).astype(int)
            df["Prijs"]  = pd.to_numeric(df["Prijs"],  errors="coerce").fillna(0.0).astype(float)

            self.engine.clear_bestellingen()
            self.engine.work_df = df
            self.engine.merge_work_df()
            self.engine.sort_work_df()

            self._loaded_bron = bron
            self.refresh_preview()
            self._validate_form()
            self.lbl_status.setText(
                f"{len(regels)} regel(s) geladen uit {bron}-queue"
            )
        except Exception as e:
            QMessageBox.critical(self, "Fout", f"Kan orders niet laden:\n{e}")

    @staticmethod
    def _pas_tabel_hoogte_aan(tabel: QTableWidget):
        hoogte = (
            tabel.horizontalHeader().height()
            + tabel.verticalHeader().length()
            + tabel.frameWidth() * 2
            + 4
        )
        tabel.setFixedHeight(max(hoogte, 80))

    def _build_ui(self):
        apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # =========================
        # TITEL (vast, buiten scroll)
        # =========================
        title = QLabel("Factuurmaker")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        root.addWidget(title)

        # =========================
        # CMS WEEKFACTUUR BANNER (vast, buiten scroll)
        # =========================
        self.banner_cms = self._maak_cms_banner()
        root.addWidget(self.banner_cms)

        # Scroll area voor de rest
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        content = QWidget()
        content.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        vbox = QVBoxLayout(content)
        vbox.setContentsMargins(0, 0, 4, 0)
        vbox.setSpacing(10)

        scroll.setWidget(content)
        root.addWidget(scroll, stretch=1)

        # Alias zodat de rest van _build_ui gewoon 'root' kan blijven gebruiken
        root = vbox

        # =========================
        # TYPE EERST KIEZEN
        # =========================
        box_type = QGroupBox("1. Documenttype")
        box_type_l = QVBoxLayout(box_type)

        row_type = QHBoxLayout()
        row_type.addWidget(QLabel("Kies documenttype:"))

        self.rb_invoice = QRadioButton("Factuur")
        self.rb_credit = QRadioButton("Creditfactuur")
        self.rb_invoice.setChecked(True)

        self.doc_group = QButtonGroup(self)
        self.doc_group.addButton(self.rb_invoice)
        self.doc_group.addButton(self.rb_credit)

        self.rb_invoice.toggled.connect(self.on_doc_type_changed)
        self.rb_credit.toggled.connect(self.on_doc_type_changed)

        row_type.setSpacing(12)
        row_type.addWidget(self.rb_invoice)
        row_type.addWidget(self.rb_credit)
        row_type.addStretch()

        box_type_l.addLayout(row_type)

        row_credit = QGridLayout()
        row_credit.setHorizontalSpacing(12)
        row_credit.setVerticalSpacing(6)

        self.lbl_original_invoice = QLabel("Originele factuur")
        self.txt_original_invoice = QLineEdit()
        self.txt_original_invoice.setPlaceholderText("Verplicht bij creditfactuur")

        self.lbl_credit_reason = QLabel("Reden")
        self.cmb_credit_reason = QComboBox()
        self.cmb_credit_reason.addItems([
            "Retour",
            "Te veel gefactureerd",
            "Verkeerd artikel",
            "Niet geleverd",
            "Overig",
        ])

        row_credit.addWidget(self.lbl_original_invoice, 0, 0)
        row_credit.addWidget(self.txt_original_invoice, 0, 1)
        row_credit.addWidget(self.lbl_credit_reason, 0, 2)
        row_credit.addWidget(self.cmb_credit_reason, 0, 3)

        box_type_l.addLayout(row_credit)
        root.addWidget(box_type)

        # =========================
        # CMS BESTANDEN LADEN
        # =========================
        box_files = QGroupBox("2. CMS orders laden")
        box_files_l = QVBoxLayout(box_files)

        row_files = QHBoxLayout()

        btn_add = QPushButton("Add CMS orders (.xlsx)")
        btn_add.setObjectName("primary")
        btn_add.clicked.connect(self.on_add_files)

        btn_reset = QPushButton("Reset")
        btn_reset.setObjectName("danger")
        btn_reset.clicked.connect(self.on_reset)

        self.lbl_status = QLabel("No CMS orders loaded")

        row_files.addWidget(btn_add)
        row_files.addWidget(btn_reset)
        row_files.addStretch()
        row_files.addWidget(self.lbl_status)

        box_files_l.addLayout(row_files)
        root.addWidget(box_files)

        # =========================
        # KLANT / DOCUMENTGEGEVENS
        # =========================
        box_details = QGroupBox("3. Klant- en documentgegevens")
        box_details_l = QVBoxLayout(box_details)

        form = QGridLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)

        self.txt_invoice = QLineEdit(self.engine.invoice_number)
        self.txt_invoice.setPlaceholderText("Leeg laten voor automatisch nummer")

        self.txt_supplier = QLineEdit()
        self.txt_supplier.setPlaceholderText("Optioneel (CMS ref / klantnummer)")

        self.spin_shipping = QDoubleSpinBox()
        self.spin_shipping.setDecimals(2)
        self.spin_shipping.setMaximum(9999)
        self.spin_shipping.setSuffix(" €")
        self.spin_shipping.valueChanged.connect(self._update_engine)
        self.spin_shipping.valueChanged.connect(self._validate_form)

        self.txt_billto = QLineEdit(self.engine.bill_to)

        self.txt_address = QTextEdit()
        self.txt_address.setPlainText(self.engine.billing_address)
        self.txt_address.setFixedHeight(72)

        form.addWidget(QLabel("Documentnummer"), 0, 0)
        form.addWidget(self.txt_invoice, 0, 1)
        self.txt_original_invoice.textChanged.connect(self._validate_form)

        form.addWidget(QLabel("Klant-/CMS nummer"), 0, 2)
        form.addWidget(self.txt_supplier, 0, 3)

        form.addWidget(QLabel("Verzend/handling"), 1, 0)
        form.addWidget(self.spin_shipping, 1, 1)

        form.addWidget(QLabel("Bill to"), 1, 2)
        form.addWidget(self.txt_billto, 1, 3)
        self.txt_billto.textChanged.connect(self._validate_form)

        form.addWidget(QLabel("Billing address"), 2, 0)
        form.addWidget(self.txt_address, 2, 1, 1, 3)
        self.txt_address.textChanged.connect(self._validate_form)

        box_details_l.addLayout(form)
        root.addWidget(box_details)

        # =========================
        # PREVIEW / REGELS
        # =========================
        box_preview = QGroupBox("4. Regels controleren en aanpassen")
        box_preview_l = QVBoxLayout(box_preview)
        box_preview_l.setSpacing(8)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Zoek part number in preview...")
        self.txt_search.textChanged.connect(self.on_search_changed)
        box_preview_l.addWidget(self.txt_search)

        self.table = QTableWidget(0, 5)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.table.itemChanged.connect(self.on_table_edited)
        self.table.setHorizontalHeaderLabels(["Part Number", "Description", "Qty", "Price", "Total"])
        self.table.setAlternatingRowColors(True)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        self.table.verticalHeader().setDefaultSectionSize(28)

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        box_preview_l.addWidget(self.table)

        row_edit = QHBoxLayout()

        btn_add_row = QPushButton("➕  Artikel zelf toevoegen")
        btn_add_row.setObjectName("secondary")
        btn_add_row.setToolTip(
            "Gebruik dit als een artikel niet automatisch gevonden is,\n"
            "maar je het toch wilt factureren."
        )
        btn_add_row.clicked.connect(self.on_add_row)

        btn_remove = QPushButton("Verwijder geselecteerde regel")
        btn_remove.setObjectName("danger")
        btn_remove.clicked.connect(self.on_remove_row)

        row_edit.addWidget(btn_add_row)
        row_edit.addWidget(btn_remove)
        row_edit.addStretch()

        box_preview_l.addLayout(row_edit)
        root.addWidget(box_preview, stretch=1)

        # =========================
        # GENEREREN
        # =========================
        box_actions = QGroupBox("5. Genereren")
        box_actions_l = QHBoxLayout(box_actions)
        self.lbl_validation = QLabel("")
        self.lbl_validation.setWordWrap(True)
        self.lbl_validation.setStyleSheet("""
            QLabel {
                background: palette(base);
                border: 1px solid palette(mid);
                border-radius: 8px;
                padding: 10px;
                font-weight: bold;
            }
        """)

        box_actions_l.addWidget(self.lbl_validation)

        self.btn_generate = QPushButton("Generate invoice (PDF)")
        self.btn_generate.setObjectName("primary")
        self.btn_generate.setMinimumHeight(40)
        self.btn_generate.clicked.connect(self.on_generate)

        btn_open = QPushButton("Open output folder")
        btn_open.setObjectName("secondary")
        btn_open.clicked.connect(self.on_open_output)

        box_actions_l.addStretch()
        box_actions_l.addWidget(self.btn_generate)
        box_actions_l.addWidget(btn_open)

        root.addWidget(box_actions)

        # =========================
        # BESTAANDE FACTUUR LADEN
        # =========================
        box_herstel = QGroupBox("6. Bestaande factuur laden (fout herstellen)")
        box_herstel_l = QVBoxLayout(box_herstel)

        lbl_herstel = QLabel(
            "Laad een eerder gegenereerde factuur om deze te bewerken en opnieuw te genereren."
        )
        lbl_herstel.setWordWrap(True)
        box_herstel_l.addWidget(lbl_herstel)

        rij_herstel = QHBoxLayout()

        btn_laden_json = QPushButton("Laden uit JSON (snel)")
        btn_laden_json.setObjectName("secondary")
        btn_laden_json.setMinimumHeight(36)
        btn_laden_json.setToolTip(
            "Selecteer het .json-bestand naast de PDF.\n"
            "Alle gegevens worden exact hersteld."
        )
        btn_laden_json.clicked.connect(self._herstel_uit_json)

        btn_laden_pdf = QPushButton("Laden uit PDF")
        btn_laden_pdf.setObjectName("secondary")
        btn_laden_pdf.setMinimumHeight(36)
        btn_laden_pdf.setToolTip(
            "Selecteer de originele PDF.\n"
            "De tekst wordt automatisch uitgelezen (minder nauwkeurig)."
        )
        btn_laden_pdf.clicked.connect(self._herstel_uit_pdf)

        rij_herstel.addWidget(btn_laden_json)
        rij_herstel.addWidget(btn_laden_pdf)
        rij_herstel.addStretch()
        box_herstel_l.addLayout(rij_herstel)

        root.addWidget(box_herstel)

        self.on_doc_type_changed()
    # =====================================================
    # HERSTEL BESTAANDE FACTUUR
    # =====================================================

    def _herstel_uit_data(self, data: dict):
        """Vul alle formuliervelden vanuit een geladen factuur-dict."""
        self.engine.clear_bestellingen()
        self.engine.work_df = data["df"].copy()
        self.engine.merge_work_df()
        self.engine.sort_work_df()

        self.txt_invoice.setText(data.get("invoice_number", ""))
        self.txt_supplier.setText(data.get("supplier_number", ""))
        self.txt_billto.setText(data.get("bill_to", ""))
        self.txt_address.setPlainText(data.get("billing_address", ""))
        self.spin_shipping.setValue(abs(float(data.get("verzendkosten", 0.0))))

        doc_type = data.get("document_type", "invoice")
        if doc_type == "credit":
            self.rb_credit.setChecked(True)
        else:
            self.rb_invoice.setChecked(True)

        orig = data.get("original_invoice_number", "")
        self.txt_original_invoice.setText(orig)

        reason = data.get("credit_reason", "")
        idx = self.cmb_credit_reason.findText(reason)
        if idx >= 0:
            self.cmb_credit_reason.setCurrentIndex(idx)

        self._loaded_bron = None
        self.refresh_preview()
        self._validate_form()
        self.lbl_status.setText(
            f"{len(self.engine.work_df)} regel(s) geladen — pas aan en genereer opnieuw"
        )

    def _herstel_uit_json(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Selecteer factuur-JSON", str(self.engine.output_dir), "JSON bestanden (*.json)"
        )
        if not path:
            return
        try:
            data = self.engine.load_draft_json(path)
            self._herstel_uit_data(data)
        except Exception as e:
            QMessageBox.critical(self, "Fout", f"Kan JSON niet laden:\n{e}")

    def _herstel_uit_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Selecteer factuur-PDF", str(self.engine.output_dir), "PDF bestanden (*.pdf)"
        )
        if not path:
            return
        try:
            from services.factuur_pdf_parser import parse_invoice_pdf
            data = parse_invoice_pdf(path)
            self._herstel_uit_data(data)
        except Exception as e:
            QMessageBox.critical(self, "Fout", f"Kan PDF niet inlezen:\n{e}")

    # =====================================================
    # LOGIC
    # =====================================================
    def _preview_issues(self):
        errors = []
        warnings = []

        if self.engine.work_df is None or self.engine.work_df.empty:
            return errors, warnings

        df = self.engine.work_df

        for idx, row in df.iterrows():
            artikel = str(row.get("Artikel", "") or "").strip()
            omschrijving = str(row.get("Omschrijving", "") or "").strip()

            try:
                aantal = int(float(row.get("Aantal", 0)))
            except Exception:
                aantal = 0

            try:
                prijs = float(row.get("Prijs", 0) or 0)
            except Exception:
                prijs = 0.0

            row_no = idx + 1

            # 🔴 FOUTEN
            if not artikel:
                errors.append(f"Regel {row_no}: artikelnummer ontbreekt")

            if not omschrijving:
                errors.append(f"Regel {row_no}: omschrijving ontbreekt")

            if aantal == 0:
                errors.append(f"Regel {row_no}: aantal is 0")

            # 🟡 WAARSCHUWINGEN
            if prijs == 0:
                warnings.append(f"Regel {row_no}: prijs is 0,00")

            if prijs < 0:
                warnings.append(f"Regel {row_no}: negatieve prijs")

        return errors, warnings
    
    def _validate_form(self):
        errors = []

        # ===== Document type
        is_credit = self.rb_credit.isChecked()

        # ===== Klantgegevens
        if not (self.txt_billto.text() or "").strip():
            errors.append("Geen klant (Bill to) ingevuld")

        if not (self.txt_address.toPlainText() or "").strip():
            errors.append("Geen adres ingevuld")

        # ===== Credit check
        if is_credit:
            if not (self.txt_original_invoice.text() or "").strip():
                errors.append("Creditfactuur zonder originele factuur")

        # ===== Data check
        if self.engine.work_df is None or self.engine.work_df.empty:
            errors.append("Geen regels (producten) geladen")

        preview_errors, preview_warnings = self._preview_issues()

        if preview_errors:
            errors.append(f"{len(preview_errors)} foutieve previewregel(s)")
            errors.extend(preview_errors[:2])

        # warnings alleen tonen, niet blokkeren
        warning_text = ""
        if preview_warnings:
            warning_text = "\n⚠️ Waarschuwingen:\n- " + "\n- ".join(preview_warnings[:2])
        if errors:
            self.lbl_validation.setText(
                "❌ Niet klaar:\n- " + "\n- ".join(errors) + warning_text
            )
            self.lbl_validation.setStyleSheet("""
                QLabel {
                    background: rgba(255, 0, 0, 0.08);
                    border: 1px solid rgba(255, 0, 0, 0.4);
                    border-radius: 8px;
                    padding: 10px;
                    font-weight: bold;
                    color: #a00000;
                }
            """)
            self.btn_generate.setEnabled(False)

        else:
            text = "✅ Klaar om te genereren"
            if preview_warnings:
                text += warning_text

            self.lbl_validation.setText(text)

            self.lbl_validation.setStyleSheet("""
                QLabel {
                    background: rgba(0, 150, 0, 0.08);
                    border: 1px solid rgba(0, 150, 0, 0.4);
                    border-radius: 8px;
                    padding: 10px;
                    font-weight: bold;
                    color: #006600;
                }
            """)
            self.btn_generate.setEnabled(True)

        # ===== Totaal check (optioneel maar sterk)
        try:
            df = self.engine.work_df
            if df is not None and not df.empty:
                total = (df["Aantal"] * df["Prijs"]).sum() + self.engine.verzendkosten
                if abs(total) < 0.01:
                    errors.append("Totaal is 0.00")
        except Exception:
            pass

        # ===== RESULTAAT
        if errors:
            self.lbl_validation.setText("❌ Niet klaar:\n- " + "\n- ".join(errors))
            self.lbl_validation.setStyleSheet("""
                QLabel {
                    background: rgba(255, 0, 0, 0.08);
                    border: 1px solid rgba(255, 0, 0, 0.4);
                    border-radius: 8px;
                    padding: 10px;
                    font-weight: bold;
                    color: #a00000;
                }
            """)
            self.btn_generate.setEnabled(False)
        else:
            self.lbl_validation.setText("✅ Klaar om te genereren")
            self.lbl_validation.setStyleSheet("""
                QLabel {
                    background: rgba(0, 150, 0, 0.08);
                    border: 1px solid rgba(0, 150, 0, 0.4);
                    border-radius: 8px;
                    padding: 10px;
                    font-weight: bold;
                    color: #006600;
                }
            """)
            self.btn_generate.setEnabled(True)

    def _update_engine(self):
        self.engine.verzendkosten = float(self.spin_shipping.value())

    def on_doc_type_changed(self):
        is_credit = self.rb_credit.isChecked()

        self.lbl_original_invoice.setVisible(is_credit)
        self.txt_original_invoice.setVisible(is_credit)
        self.lbl_credit_reason.setVisible(is_credit)
        self.cmb_credit_reason.setVisible(is_credit)
        
        self.btn_generate.setText("Generate credit invoice (PDF)" if is_credit else "Generate invoice (PDF)")
        self._validate_form()
        self.refresh_preview()

    def on_add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select CMS orders", "", "Excel files (*.xlsx)")
        if not paths:
            return
        for p in paths:
            self.engine.add_cms_bestelling(p)
        self.refresh_preview()
        self._validate_form()

    def on_reset(self):
        if QMessageBox.question(self, "Reset", "Remove all CMS orders?") != QMessageBox.Yes:
            return
        self.engine.clear_bestellingen()
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        self.table.blockSignals(False)
        self.lbl_status.setText("No CMS orders loaded")
        self._validate_form()

    def refresh_preview(self):
        df = self.engine.work_df
        if df is None:
            return

        is_credit = self.rb_credit.isChecked()
        sign = -1 if is_credit else 1

        self.table.blockSignals(True)
        self.table.setRowCount(0)
        self.lbl_status.setText(f"{len(self.engine.cms_paths)} CMS order file(s) loaded")

        for _, r in df.iterrows():
            row = self.table.rowCount()
            self.table.insertRow(row)

            artikel = str(r.get("Artikel", "") or "").strip()
            omschrijving = str(r.get("Omschrijving", "") or "").strip()

            try:
                qty_raw = int(r["Aantal"])
            except Exception:
                qty_raw = 0

            try:
                price = float(r["Prijs"])
            except Exception:
                price = 0.0

            qty = qty_raw * sign
            total = qty * price

            item_artikel = QTableWidgetItem(artikel)
            item_omschrijving = QTableWidgetItem(omschrijving)
            item_qty = QTableWidgetItem(str(qty))
            item_price = QTableWidgetItem(currency(price))
            item_total = QTableWidgetItem(currency(total))
            warn_bg = QColor(255, 245, 200)   # geel
            error_bg = QColor(255, 220, 220)  # licht rood

            # 🔴 fouten
            if not artikel:
                item_artikel.setBackground(error_bg)

            if not omschrijving:
                item_omschrijving.setBackground(error_bg)

            if qty_raw == 0:
                item_qty.setBackground(error_bg)

            # 🟡 waarschuwingen
            if price == 0:
                item_price.setBackground(warn_bg)

            if price < 0:
                item_price.setBackground(warn_bg)

            self.table.setItem(row, 0, item_artikel)
            self.table.setItem(row, 1, item_omschrijving)
            self.table.setItem(row, 2, item_qty)
            self.table.setItem(row, 3, item_price)
            self.table.setItem(row, 4, item_total)

        self.table.blockSignals(False)
        self.on_search_changed(self.txt_search.text())
        self._pas_tabel_hoogte_aan(self.table)

    def _parse_float_eu(self, s: str) -> float:
        s = (s or "").strip().replace("€", "").strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        return float(s)

    def on_generate(self):
        try:
            bill_to = (self.txt_billto.text() or "").strip()
            addr = (self.txt_address.toPlainText() or "").strip()
            if not bill_to:
                QMessageBox.warning(self, "Missing info", "Vul 'Bill to' in.")
                return
            if not addr:
                QMessageBox.warning(self, "Missing info", "Vul 'Billing address' in.")
                return

            is_credit = self.rb_credit.isChecked()

            inv_no = (self.txt_invoice.text() or "").strip()
            self.engine.invoice_number = inv_no  # engine may auto-number if blank

            self.engine.supplier_number = (self.txt_supplier.text() or "").strip()
            self.engine.bill_to = bill_to
            self.engine.billing_address = addr
            self.engine.logo_path = str(resource_path("assets/logo.png"))
            self._update_engine()

            self.engine.document_type = "credit" if is_credit else "invoice"
            self.engine.original_invoice_number = (self.txt_original_invoice.text() or "").strip()
            self.engine.credit_reason = (self.cmb_credit_reason.currentText() or "").strip()

            if self.engine.document_type == "credit" and not self.engine.original_invoice_number:
                QMessageBox.warning(self, "Missing info", "Voor een creditfactuur is 'Originele factuur' verplicht.")
                return

            path = self.engine.generate_pdf()
            self.txt_invoice.setText(self.engine.invoice_number)

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return

        QMessageBox.information(self, "Document aangemaakt", f"PDF opgeslagen:\n{path}")

        # Vraag om queue te markeren als verwerkt
        if self._loaded_bron:
            antw = QMessageBox.question(
                self,
                "Orders verwerkt?",
                f"Wil je de {self._loaded_bron}-orders markeren als verwerkt?\n\n"
                "Ze verdwijnen dan uit de lijst voor de weekfactuur.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if antw == QMessageBox.Yes:
                try:
                    from services.cms_queue import mark_verwerkt
                    mark_verwerkt(self._loaded_bron)
                except Exception:
                    pass
                self._loaded_bron = None
                self._refresh_banner()

    def on_open_output(self):
        folder = output_root() / "facturen"
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))

    def on_search_changed(self, text: str):
        q = (text or "").strip().lower()
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            part = item.text().lower() if item else ""
            self.table.setRowHidden(r, bool(q) and (q not in part))

    def on_table_edited(self, item):
        row = item.row()
        col = item.column()
        mapping = {0: "Artikel", 1: "Omschrijving", 2: "Aantal", 3: "Prijs"}
        if col not in mapping:
            return

        field = mapping[col]
        try:
            if field == "Aantal":
                v = int(float(item.text().strip()))
                self.engine.work_df.at[row, field] = abs(v)
            elif field == "Prijs":
                self.engine.work_df.at[row, field] = self._parse_float_eu(item.text())
            else:
                self.engine.work_df.at[row, field] = item.text()
        except Exception:
            return

        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()
        self._validate_form()

    def on_remove_row(self):
        row = self.table.currentRow()
        if row < 0:
            return
        if QMessageBox.question(self, "Remove item", "Remove selected row from document?") != QMessageBox.Yes:
            return
        self.engine.work_df = self.engine.work_df.drop(self.engine.work_df.index[row]).reset_index(drop=True)
        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()
        self._validate_form()

    def on_add_row(self):
        if self.engine.work_df is None:
            self.engine.clear_bestellingen()
        new_row = {
            "Artikel": "",
            "Omschrijving": "",
            "Aantal": 1,
            "Prijs": 0.0
        }
        self.engine.work_df = pd.concat([self.engine.work_df, pd.DataFrame([new_row])], ignore_index=True)
        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()
        self._validate_form()
