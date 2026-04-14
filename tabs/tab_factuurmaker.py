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
    QGroupBox, QButtonGroup, QTableWidgetItem, QMessageBox, QFileDialog
)

from engines.engine_factuurmaker import FactuurMakerEngine
from utils.paths import output_root
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
        self._build_ui()
        self._validate_form()

    def _build_ui(self):
        apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # =========================
        # TITEL
        # =========================
        title = QLabel("Factuurmaker")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        root.addWidget(title)

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
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.table.itemChanged.connect(self.on_table_edited)
        self.table.setHorizontalHeaderLabels(["Part Number", "Description", "Qty", "Price", "Total"])
        self.table.setAlternatingRowColors(True)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.table.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
        self.table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        self.table.verticalHeader().setDefaultSectionSize(28)

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        box_preview_l.addWidget(self.table, stretch=1)

        row_edit = QHBoxLayout()

        btn_add_row = QPushButton("Add manual item")
        btn_add_row.setObjectName("secondary")
        btn_add_row.clicked.connect(self.on_add_row)

        btn_remove = QPushButton("Remove selected row")
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

        self.on_doc_type_changed()
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
            self.engine.logo_path = "assets/logo.png"
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

        QMessageBox.information(self, "Document created", f"PDF generated:\n{path}")

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
