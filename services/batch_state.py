# services/batch_state.py
import json
from pathlib import Path
from datetime import datetime


class BatchStore:
    def __init__(self, state_file="data/batch_state.json"):
        self.state_file = Path(state_file)
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        if not self.state_file.exists():
            self._save({"batches": []})

    def _load(self):
        try:
            with open(self.state_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"batches": []}

    def _save(self, data):
        with open(self.state_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def create_batch(self, batch_data: dict):
        data = self._load()
        batch = dict(batch_data)
        batch.setdefault("created_at", datetime.now().isoformat(timespec="seconds"))
        batch.setdefault("status", "PENDING_IMPORT")
        data["batches"].append(batch)
        self._save(data)

    def get_open_batches(self, tab: str):
        data = self._load()
        return [
            b for b in data.get("batches", [])
            if str(b.get("tab")) == str(tab) and b.get("status") == "PENDING_IMPORT"
        ]

    def get_latest_open_batch(self, tab: str):
        batches = self.get_open_batches(tab)
        if not batches:
            return None
        return sorted(batches, key=lambda x: x.get("created_at", ""), reverse=True)[0]

    def mark_imported(self, batch_id: str):
        data = self._load()
        changed = False
        for b in data.get("batches", []):
            if b.get("batch_id") == batch_id:
                b["status"] = "IMPORTED"
                b["imported_at"] = datetime.now().isoformat(timespec="seconds")
                changed = True
                break
        if changed:
            self._save(data)
        return changed

    def mark_merged(self, batch_id: str, merged_into_batch_id: str):
        data = self._load()
        changed = False
        for b in data.get("batches", []):
            if b.get("batch_id") == batch_id:
                b["status"] = "MERGED"
                b["merged_at"] = datetime.now().isoformat(timespec="seconds")
                b["merged_into_batch_id"] = merged_into_batch_id
                changed = True
                break
        if changed:
            self._save(data)
        return changed

    def get_batch(self, batch_id: str):
        data = self._load()
        for b in data.get("batches", []):
            if b.get("batch_id") == batch_id:
                return b
        return None