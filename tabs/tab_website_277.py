
# ============================================================
# tabs/tab_website_277.py
# UI – Website / CMS 277
# Batch-afboeken (alleen voorraad)
# ============================================================

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog,
    QMessageBox, QTextEdit
)
from PySide6.QtCore import Qt
import os
from datetime import datetime

from engines.engine_website_277 import Website277Engine
from services.batch_state import BatchStore
from services.batch_merge_277 import (
    load_changes,
    merge_changes,
    build_update_from_changes,
    save_merged_files,
)
from utils.paths import output_root
from utils.theme import apply_theme

class TabWebsite277(QWidget):
    def __init__(self, app_state):
        super().__init__()
        self.app_state = app_state
        self.engine = Website277Engine(app_state)
        self.batch_store = BatchStore()
        self._build_ui()

    # --------------------------------------------------------
    # UI
    # --------------------------------------------------------
    def _build_ui(self):
        apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # =========================
        # TITEL
        # =========================
        title = QLabel("Website / 277 – Voorraad afboeken")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        root.addWidget(title)

        # =========================
        # UITLEG
        # =========================
        uitleg = QLabel(
            "Verwerkt website/CMS bestellingen (277).\n\n"
            "① Zorg dat WooCommerce export geladen is\n"
            "② Voeg CMS bestanden toe\n"
            "③ Start afboeken\n\n"
            "• Alleen voorraad wordt aangepast\n"
            "• Tekorten worden apart gelogd"
        )
        uitleg.setWordWrap(True)
        root.addWidget(uitleg)

        # =========================
        # STATUS BLOK
        # =========================
        self.lbl_wc = QLabel()
        self.lbl_batch = QLabel()

        for lbl in (self.lbl_wc, self.lbl_batch):
            lbl.setWordWrap(True)
            lbl.setStyleSheet("""
                QLabel {
                    background: palette(base);
                    border: 1px solid palette(mid);
                    border-radius: 8px;
                    padding: 10px;
                }
            """)

        root.addWidget(self.lbl_wc)
        root.addWidget(self.lbl_batch)

        self._update_wc()
        self._update_batch_status()

        # =========================
        # ACTIES (boven)
        # =========================
        row_actions = QHBoxLayout()

        btn_add = QPushButton("Add CMS 277 orders")
        btn_add.setObjectName("primary")
        btn_add.clicked.connect(self.on_add)

        btn_reset = QPushButton("Reset")
        btn_reset.setObjectName("secondary")
        btn_reset.clicked.connect(self.on_reset)

        self.lbl_count = QLabel("0 bestanden geladen")

        row_actions.addWidget(btn_add)
        row_actions.addWidget(btn_reset)
        row_actions.addStretch()
        row_actions.addWidget(self.lbl_count)

        root.addLayout(row_actions)

        # =========================
        # RUN ACTIES (duidelijk)
        # =========================
        row_run = QHBoxLayout()

        btn_run = QPushButton("Start afboeken")
        btn_run.setObjectName("primary")
        btn_run.setMinimumHeight(40)
        btn_run.clicked.connect(self.on_run)

        btn_open = QPushButton("Open output")
        btn_open.setObjectName("secondary")
        btn_open.clicked.connect(self.on_open_output)

        btn_mark = QPushButton("Markeer als geïmporteerd")
        btn_mark.setObjectName("secondary")
        btn_mark.clicked.connect(self.on_mark_imported)

        row_run.addWidget(btn_run)
        row_run.addWidget(btn_open)
        row_run.addWidget(btn_mark)
        row_run.addStretch()

        root.addLayout(row_run)

        # =========================
        # LOG
        # =========================
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Log verschijnt hier...")
        root.addWidget(self.log, stretch=1)

        self._sync_label()

    # --------------------------------------------------------
    # HELPERS
    # --------------------------------------------------------
    def _update_wc(self):
        if self.app_state.wc_path and os.path.exists(self.app_state.wc_path):
            mtime = os.path.getmtime(self.app_state.wc_path)
            laatst = datetime.fromtimestamp(mtime).strftime("%d-%m-%Y %H:%M:%S")
            self.lbl_wc.setText(f"✅ WC-export in gebruik:\n{self.app_state.wc_path}\nLaatst gewijzigd: {laatst}")
        else:
            self.lbl_wc.setText("❌ Geen WC-export geladen")


    def _update_batch_status(self):
        batch = self.batch_store.get_latest_open_batch("277")
        if batch:
            created = batch.get("created_at", "-")
            update_path = batch.get("update_path", "-")
            self.lbl_batch.setText(
                "⚠️ Open website-update aanwezig\n"
                f"Batch: {batch.get('batch_id', '-')}\n"
                f"Aangemaakt: {created}\n"
                f"Status: {batch.get('status', '-')}\n"
                f"Updatebestand:\n{update_path}"
            )
        else:
            self.lbl_batch.setText("✅ Geen open website-update voor 277")

    def _run_merge_with_open_batch(self, open_batch: dict):
        # 1) draai eerst de nieuwe run gewoon normaal
        try:
            new_result = self.engine.run()
        except Exception as e:
            QMessageBox.critical(self, "Fout", str(e))
            return

        # 2) lees oude changes + nieuwe changes
        old_changes_path = open_batch.get("debug_changes_path", "")
        new_changes_path = new_result.get("debug_changes_path", "")

        try:
            old_df = load_changes(old_changes_path)
            new_df = load_changes(new_changes_path)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Merge fout",
                f"CHANGES bestand kon niet gelezen worden.\n\n{e}"
            )
            return

        # 3) merge changes
        try:
            merged_changes_df = merge_changes(old_df, new_df)
            merged_update_df = build_update_from_changes(
                merged_changes_df,
                self.app_state.wc_df
            )
            merged_files = save_merged_files(merged_changes_df, merged_update_df)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Merge fout",
                f"Samenvoegen is mislukt.\n\n{e}"
            )
            return

        # 4) oude open batch op MERGED zetten
        self.batch_store.mark_merged(
            open_batch["batch_id"],
            merged_files["batch_id"]
        )

        # 5) nieuwe gewone batch ook op MERGED zetten
        self.batch_store.create_batch(new_result)
        self.batch_store.mark_merged(
            new_result["batch_id"],
            merged_files["batch_id"]
        )

        # 6) merged batch opslaan als nieuwe open batch
        merged_batch = {
            "batch_id": merged_files["batch_id"],
            "tab": "277",
            "status": "PENDING_IMPORT",
            "update_path": merged_files["update_path"],
            "pick_path": "",
            "tekort_path": "",
            "stats_path": "",
            "debug_changes_path": merged_files["debug_changes_path"],
            "wc_path": str(self.app_state.wc_path or ""),
            "cms_paths": list(self.engine.cms_paths),
            "paths": [merged_files["update_path"], merged_files["debug_changes_path"]],
            "merge_source": True,
            "merged_from": [open_batch["batch_id"], new_result["batch_id"]],
        }
        self.batch_store.create_batch(merged_batch)

        # 7) UI/log opschonen
        self.log.append("\nSamenvoegen voltooid.")
        self.log.append(f"Oude batch samengevoegd: {open_batch['batch_id']}")
        self.log.append(f"Nieuwe batch samengevoegd: {new_result['batch_id']}")
        self.log.append(f"Nieuwe merged batch: {merged_files['batch_id']}")
        self.log.append(f"- {merged_files['update_path']}")
        self.log.append(f"- {merged_files['debug_changes_path']}")

        self.engine.clear()
        self._sync_label()
        self._update_batch_status()

        QMessageBox.information(
            self,
            "Samenvoegen klaar",
            "De open batch en de nieuwe run zijn samengevoegd.\n\n"
            "Er staat nu 1 nieuwe merged website-update open op WACHT OP IMPORT."
        )
    def _sync_label(self):
        self.lbl_count.setText(f"{len(self.engine.cms_paths)} CMS 277 bestand(en) toegevoegd")

    # --------------------------------------------------------
    # ACTIONS
    # --------------------------------------------------------
    def on_add(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecteer CMS 277 bestellingen",
            "",
            "Excel (*.xlsx)"
        )
        if not paths:
            return

        for p in paths:
            self.engine.add_cms_277(p)

        self.log.append(f"{len(paths)} CMS 277 bestand(en) toegevoegd")
        self._sync_label()

    def on_run(self):
        pending = self.batch_store.get_open_batches("277")

        if pending:
            open_batch = pending[-1]

            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Open website-update")
            msg.setText(
                "Er staat al een open 277 website-update.\n\n"
                f"Batch: {open_batch.get('batch_id', '-')}\n"
                f"Bestand:\n{open_batch.get('update_path', '-')}\n\n"
                "Wat wil je doen?"
            )

            btn_cancel = msg.addButton("Annuleren", QMessageBox.RejectRole)
            btn_imported = msg.addButton("Eerst als geïmporteerd markeren", QMessageBox.AcceptRole)
            btn_merge = msg.addButton("Samenvoegen met nieuwe run", QMessageBox.ActionRole)

            msg.exec()

            clicked = msg.clickedButton()

            if clicked == btn_cancel:
                self._update_batch_status()
                return

            if clicked == btn_imported:
                ok = self.batch_store.mark_imported(open_batch["batch_id"])
                if ok:
                    self.log.append(f"Batch gemarkeerd als geïmporteerd: {open_batch['batch_id']}")
                    self._update_batch_status()
                    QMessageBox.information(
                        self,
                        "Bijgewerkt",
                        "De open batch is gemarkeerd als geïmporteerd.\n"
                        "Start de run daarna opnieuw."
                    )
                else:
                    QMessageBox.warning(
                        self,
                        "Niet gelukt",
                        "De batch kon niet als geïmporteerd gemarkeerd worden."
                    )
                return

            if clicked == btn_merge:
                self._run_merge_with_open_batch(open_batch)
                return

            return

        # normale run als er geen open batch is
        try:
            result = self.engine.run()
        except Exception as e:
            QMessageBox.critical(self, "Fout", str(e))
            return

        self.batch_store.create_batch(result)

        self.log.append("\nVerwerking voltooid. Output:")
        for r in result.get("paths", []):
            self.log.append(f"- {r}")

        self.log.append(
            f"\nNieuwe open batch aangemaakt: {result.get('batch_id', '-')}\n"
            "Status: WACHT OP IMPORT"
        )

        self.engine.clear()
        self._sync_label()
        self._update_batch_status()

        QMessageBox.information(
            self,
            "277 klaar",
            "277 afboeken is afgerond.\n\n"
            "Let op: de website-update staat nu nog op WACHT OP IMPORT.\n"
            "Markeer hem pas als geïmporteerd nadat WP All Import gedaan is."
        )

    def on_mark_imported(self):
        batch = self.batch_store.get_latest_open_batch("277")
        if not batch:
            QMessageBox.information(self, "Geen open batch", "Er staat geen open 277-batch.")
            return

        ok = self.batch_store.mark_imported(batch["batch_id"])
        if not ok:
            QMessageBox.warning(self, "Niet gelukt", "Batch kon niet als geïmporteerd gemarkeerd worden.")
            return

        self.log.append(f"Batch gemarkeerd als geïmporteerd: {batch['batch_id']}")
        self._update_batch_status()
        QMessageBox.information(self, "Bijgewerkt", "De open 277-batch is gemarkeerd als geïmporteerd.")

    def on_reset(self):
        self.engine.clear()
        self.log.clear()
        self._sync_label()
        self._update_batch_status()

    def on_open_output(self):
        folder = output_root() / "277"
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))