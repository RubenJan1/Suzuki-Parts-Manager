from __future__ import annotations

import os
import pandas as pd

from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
    QLineEdit, QTextEdit, QDoubleSpinBox, QComboBox, QRadioButton, QButtonGroup,
    QSplitter, QSizePolicy, QHeaderView
)

from engines.engine_factuurmaker import FactuurMakerEngine


def currency(value) -> str:
    """Format euro with European separators."""
    try:
        v = float(value)
        return f"€ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "€ 0,00"

def _is_dark_mode(widget: QWidget) -> bool:
    bg = widget.palette().color(QPalette.Window)
    return bg.lightness() < 128


def _apply_theme(widget: QWidget) -> None:
    """
    Clean professional business/ERP theme:
    - Compact sizing (better for 14–15,6")
    - Clear selection states
    - Subtle borders, less rounding
    - Buttons not oversized
    """
    dark = _is_dark_mode(widget)

    if dark:
        style = """
        /* ===== Base ===== */
        QWidget {
            background-color: #111827;
            color: #E5E7EB;
            font-size: 10pt;
        }
        QLabel { color: #E5E7EB; }

        /* ===== Inputs ===== */
        QLineEdit, QTextEdit, QDoubleSpinBox, QComboBox {
            background-color: #0B1220;
            color: #E5E7EB;
            border: 1px solid #273244;
            padding: 6px 8px;
            border-radius: 4px;
        }
        QLineEdit:focus, QTextEdit:focus, QDoubleSpinBox:focus, QComboBox:focus {
            border: 1px solid #60A5FA;
        }

        /* ===== Buttons ===== */
        QPushButton {
            padding: 6px 12px;
            border-radius: 4px;
            font-size: 10pt;
            min-height: 32px;
            border: 1px solid transparent;
        }
        QPushButton#primary {
            background-color: #2563EB;
            color: #FFFFFF;
        }
        QPushButton#primary:hover { background-color: #1D4ED8; }
        QPushButton#secondary {
            background-color: #1F2937;
            color: #E5E7EB;
            border: 1px solid #273244;
        }
        QPushButton#secondary:hover { background-color: #243244; }
        QPushButton#danger {
            background-color: #DC2626;
            color: #FFFFFF;
        }
        QPushButton#danger:hover { background-color: #B91C1C; }

        /* ===== Radio buttons (clean, not pill) ===== */
        QRadioButton {
            spacing: 8px;
            padding: 2px 4px;
            font-weight: 500;
        }
        QRadioButton::indicator {
            width: 16px;
            height: 16px;
        }
        QRadioButton::indicator:unchecked {
            border: 1px solid #3B4A63;
            border-radius: 8px;
            background: #0B1220;
        }
        QRadioButton::indicator:checked {
            border: 1px solid #2563EB;
            border-radius: 8px;
            background: #2563EB;
        }

        /* ===== Table ===== */
        QTableWidget {
            background-color: #0B1220;
            alternate-background-color: #0E172A;
            color: #E5E7EB;
            border: 1px solid #273244;
            border-radius: 6px;
            gridline-color: #1F2A3A;
        }
        QHeaderView::section {
            background-color: #0E172A;
            color: #E5E7EB;
            font-weight: 600;
            padding: 7px 8px;
            border: 0px;
            border-bottom: 1px solid #273244;
        }
        QTableWidget::item {
            padding-left: 6px;
            padding-right: 6px;
        }
        QTableWidget::item:selected {
            background-color: #1D4ED8;
            color: #FFFFFF;
        }

        /* ===== Scrollbars ===== */
        QScrollBar:vertical { width: 12px; background: #0B1220; }
        QScrollBar::handle:vertical { background: #273244; border-radius: 6px; min-height: 28px; }
        QScrollBar::handle:vertical:hover { background: #334155; }
        QScrollBar:horizontal { height: 12px; background: #0B1220; }
        QScrollBar::handle:horizontal { background: #273244; border-radius: 6px; min-width: 28px; }
        QScrollBar::handle:horizontal:hover { background: #334155; }

        /* ===== Splitter handle ===== */
        QSplitter::handle {
            background: #273244;
        }
        """
    else:
        style = """
        /* ===== Base ===== */
        QWidget {
            background-color: #F3F4F6;
            color: #111827;
            font-size: 10pt;
        }
        QLabel { color: #111827; }

        /* ===== Inputs ===== */
        QLineEdit, QTextEdit, QDoubleSpinBox, QComboBox {
            background-color: #FFFFFF;
            color: #111827;
            border: 1px solid #D1D5DB;
            padding: 6px 8px;
            border-radius: 4px;
        }
        QLineEdit:focus, QTextEdit:focus, QDoubleSpinBox:focus, QComboBox:focus {
            border: 1px solid #2563EB;
        }

        /* ===== Buttons ===== */
        QPushButton {
            padding: 6px 12px;
            border-radius: 4px;
            font-size: 10pt;
            min-height: 32px;
            border: 1px solid transparent;
        }
        QPushButton#primary {
            background-color: #2563EB;
            color: #FFFFFF;
        }
        QPushButton#primary:hover { background-color: #1D4ED8; }
        QPushButton#secondary {
            background-color: #FFFFFF;
            color: #111827;
            border: 1px solid #D1D5DB;
        }
        QPushButton#secondary:hover { background-color: #F3F4F6; }
        QPushButton#danger {
            background-color: #DC2626;
            color: #FFFFFF;
        }
        QPushButton#danger:hover { background-color: #B91C1C; }

        /* ===== Radio buttons (clean, not pill) ===== */
        QRadioButton {
            spacing: 8px;
            padding: 2px 4px;
            font-weight: 500;
        }
        QRadioButton::indicator {
            width: 16px;
            height: 16px;
        }
        QRadioButton::indicator:unchecked {
            border: 1px solid #9CA3AF;
            border-radius: 8px;
            background: #FFFFFF;
        }
        QRadioButton::indicator:checked {
            border: 1px solid #2563EB;
            border-radius: 8px;
            background: #2563EB;
        }

        /* ===== Table ===== */
        QTableWidget {
            background-color: #FFFFFF;
            alternate-background-color: #F9FAFB;
            color: #111827;
            border: 1px solid #E5E7EB;
            border-radius: 6px;
            gridline-color: #EDEDED;
        }
        QHeaderView::section {
            background-color: #F9FAFB;
            color: #111827;
            font-weight: 600;
            padding: 7px 8px;
            border: 0px;
            border-bottom: 1px solid #E5E7EB;
        }
        QTableWidget::item {
            padding-left: 6px;
            padding-right: 6px;
        }
        QTableWidget::item:selected {
            background-color: #DBEAFE;
            color: #111827;
        }

        /* ===== Scrollbars ===== */
        QScrollBar:vertical { width: 12px; background: #F3F4F6; }
        QScrollBar::handle:vertical { background: #D1D5DB; border-radius: 6px; min-height: 28px; }
        QScrollBar::handle:vertical:hover { background: #9CA3AF; }
        QScrollBar:horizontal { height: 12px; background: #F3F4F6; }
        QScrollBar::handle:horizontal { background: #D1D5DB; border-radius: 6px; min-width: 28px; }
        QScrollBar::handle:horizontal:hover { background: #9CA3AF; }

        /* ===== Splitter handle ===== */
        QSplitter::handle {
            background: #E5E7EB;
        }
        """

    widget.setStyleSheet(style)


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

    def _build_ui(self):
        _apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)

        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(10)
        root.addWidget(splitter)

        # =========================
        # TOP: Orders + Details + Type/Actions
        # =========================
        top = QWidget()
        top_l = QVBoxLayout(top)
        top_l.setSpacing(10)
        top_l.setContentsMargins(4, 4, 4, 4)

        # STEP 1
        lbl_step1 = QLabel("① CMS orders")
        lbl_step1.setStyleSheet("font-size: 12.5pt; font-weight: bold;")
        top_l.addWidget(lbl_step1)

        row_files = QHBoxLayout()
        btn_add = QPushButton("Add CMS orders (.xlsx)")
        btn_add.setObjectName("primary")
        btn_add.clicked.connect(self.on_add_files)

        btn_reset = QPushButton("Reset")
        btn_reset.setObjectName("danger")
        btn_reset.clicked.connect(self.on_reset)

        self.lbl_status = QLabel("No CMS orders loaded")
        self.lbl_status.setStyleSheet("color: #6B7280;")

        row_files.addWidget(btn_add)
        row_files.addWidget(btn_reset)
        row_files.addStretch()
        row_files.addWidget(self.lbl_status)
        top_l.addLayout(row_files)

        # STEP 2
        lbl_step2 = QLabel("② Document details")
        lbl_step2.setStyleSheet("font-size: 12.5pt; font-weight: bold;")
        top_l.addWidget(lbl_step2)

        form = QGridLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(6)

        self.txt_invoice = QLineEdit(self.engine.invoice_number)
        self.txt_invoice.setPlaceholderText("Leeg laten voor automatisch nummer")

        self.txt_supplier = QLineEdit()
        self.txt_supplier.setPlaceholderText("Optioneel (CMS ref / klantnummer)")

        self.spin_shipping = QDoubleSpinBox()
        self.spin_shipping.setDecimals(2)
        self.spin_shipping.setMaximum(9999)
        self.spin_shipping.setSuffix(" €")
        self.spin_shipping.valueChanged.connect(self._update_engine)

        self.txt_billto = QLineEdit(self.engine.bill_to)

        self.txt_address = QTextEdit()
        self.txt_address.setPlainText(self.engine.billing_address)  # ensures real multiline
        self.txt_address.setFixedHeight(72)

        form.addWidget(QLabel("Documentnummer"), 0, 0)
        form.addWidget(self.txt_invoice, 0, 1)

        form.addWidget(QLabel("Klant-/CMS nummer"), 0, 2)
        form.addWidget(self.txt_supplier, 0, 3)

        form.addWidget(QLabel("Verzend/handling"), 1, 0)
        form.addWidget(self.spin_shipping, 1, 1)

        form.addWidget(QLabel("Bill to"), 1, 2)
        form.addWidget(self.txt_billto, 1, 3)

        form.addWidget(QLabel("Billing address"), 2, 0)
        form.addWidget(self.txt_address, 2, 1, 1, 3)

        top_l.addLayout(form)

        # Type + Actions in one row (less wasted space)
        row_type_actions = QHBoxLayout()
        row_type_actions.setSpacing(10)

        self.rb_invoice = QRadioButton("Factuur")
        self.rb_credit = QRadioButton("Creditfactuur")
        self.rb_invoice.setChecked(True)

        self.doc_group = QButtonGroup(self)
        self.doc_group.addButton(self.rb_invoice)
        self.doc_group.addButton(self.rb_credit)
        self.rb_invoice.toggled.connect(self.on_doc_type_changed)
        self.rb_credit.toggled.connect(self.on_doc_type_changed)

        self.lbl_original_invoice = QLabel("Originele factuur")
        self.txt_original_invoice = QLineEdit()
        self.txt_original_invoice.setPlaceholderText("Factuurnummer waar deze credit bij hoort (verplicht)")

        self.lbl_credit_reason = QLabel("Reden")
        self.cmb_credit_reason = QComboBox()
        self.cmb_credit_reason.addItems(["Retour", "Te veel gefactureerd", "Verkeerd artikel", "Niet geleverd", "Overig"])

        # Left side: type toggles + credit fields (stacked compact)
        left = QVBoxLayout()
        left.setSpacing(6)

        row_type = QHBoxLayout()
        row_type.addWidget(QLabel("Type"))
        row_type.addSpacing(8)
        row_type.addWidget(self.rb_invoice)
        row_type.addWidget(self.rb_credit)
        row_type.addStretch()
        left.addLayout(row_type)

        row_credit = QGridLayout()
        row_credit.setHorizontalSpacing(12)
        row_credit.setVerticalSpacing(6)
        row_credit.addWidget(self.lbl_original_invoice, 0, 0)
        row_credit.addWidget(self.txt_original_invoice, 0, 1)
        row_credit.addWidget(self.lbl_credit_reason, 0, 2)
        row_credit.addWidget(self.cmb_credit_reason, 0, 3)
        left.addLayout(row_credit)

        left_wrap = QWidget()
        left_wrap.setLayout(left)

        # Right side: actions
        actions = QHBoxLayout()
        actions.setSpacing(10)

        self.btn_generate = QPushButton("Generate invoice (PDF)")
        self.btn_generate.setObjectName("primary")
        self.btn_generate.clicked.connect(self.on_generate)

        btn_open = QPushButton("Open output folder")
        btn_open.setObjectName("secondary")
        btn_open.clicked.connect(self.on_open_output)

        actions.addWidget(self.btn_generate)
        actions.addWidget(btn_open)

        actions_wrap = QWidget()
        actions_wrap.setLayout(actions)

        row_type_actions.addWidget(left_wrap, 1)
        row_type_actions.addStretch()
        row_type_actions.addWidget(actions_wrap, 0)

        top_l.addLayout(row_type_actions)

        splitter.addWidget(top)

        # =========================
        # BOTTOM: Preview
        # =========================
        bottom = QWidget()
        bottom_l = QVBoxLayout(bottom)
        bottom_l.setSpacing(8)
        bottom_l.setContentsMargins(4, 4, 4, 4)

        lbl_step3 = QLabel("③ Preview (sleep de scheidslijn om groter te maken)")
        lbl_step3.setStyleSheet("font-size: 12.5pt; font-weight: bold;")
        bottom_l.addWidget(lbl_step3)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Search part number...")
        self.txt_search.textChanged.connect(self.on_search_changed)
        bottom_l.addWidget(self.txt_search)

        self.table = QTableWidget(0, 5)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.table.itemChanged.connect(self.on_table_edited)
        self.table.setHorizontalHeaderLabels(["Part Number", "Description", "Qty", "Price", "Total"])
        self.table.setAlternatingRowColors(True)

        # Better scrolling on small screens
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.table.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
        self.table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        self.table.verticalHeader().setDefaultSectionSize(28)

        # Smarter column widths (reduce horizontal scroll)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Part
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)           # Description
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Qty
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Price
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Total

        bottom_l.addWidget(self.table, stretch=1)

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
        bottom_l.addLayout(row_edit)

        splitter.addWidget(bottom)

        # Preview first: make bottom larger by default
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([240, 780])

        self.on_doc_type_changed()

    # =====================================================
    # LOGIC
    # =====================================================

    def _update_engine(self):
        self.engine.verzendkosten = float(self.spin_shipping.value())

    def on_doc_type_changed(self):
        is_credit = self.rb_credit.isChecked()

        self.lbl_original_invoice.setVisible(is_credit)
        self.txt_original_invoice.setVisible(is_credit)
        self.lbl_credit_reason.setVisible(is_credit)
        self.cmb_credit_reason.setVisible(is_credit)

        self.btn_generate.setText("Generate credit invoice (PDF)" if is_credit else "Generate invoice (PDF)")
        self.refresh_preview()

    def on_add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select CMS orders", "", "Excel files (*.xlsx)")
        if not paths:
            return
        for p in paths:
            self.engine.add_cms_bestelling(p)
        self.refresh_preview()

    def on_reset(self):
        if QMessageBox.question(self, "Reset", "Remove all CMS orders?") != QMessageBox.Yes:
            return
        self.engine.clear_bestellingen()
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        self.table.blockSignals(False)
        self.lbl_status.setText("No CMS orders loaded")

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

            qty = int(r["Aantal"]) * sign
            price = float(r["Prijs"])
            total = qty * price

            self.table.setItem(row, 0, QTableWidgetItem(str(r["Artikel"])))
            self.table.setItem(row, 1, QTableWidgetItem(str(r["Omschrijving"])))
            self.table.setItem(row, 2, QTableWidgetItem(str(qty)))
            self.table.setItem(row, 3, QTableWidgetItem(currency(price)))
            self.table.setItem(row, 4, QTableWidgetItem(currency(total)))

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
        folder = os.path.abspath("output/facturen")
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

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

    def on_add_row(self):
        if self.engine.work_df is None:
            self.engine.clear_bestellingen()
        new_row = {"Artikel": "MANUAL", "Omschrijving": "Handmatig toegevoegd", "Aantal": 1, "Prijs": 0.0}
        self.engine.work_df = pd.concat([self.engine.work_df, pd.DataFrame([new_row])], ignore_index=True)
        self.engine.merge_work_df()
        self.engine.sort_work_df()
        self.refresh_preview()
