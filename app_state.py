from datetime import datetime
import uuid
from utils.paths import appdata_root, output_root

class AppState:
    def __init__(self):
        self.wc_df = None
        self.wc_path = None
        self.loaded_at = None
        self.session_id = uuid.uuid4().hex[:8]

        self.appdata_dir = appdata_root()
        self.output_dir = output_root()

    def set_wc_export(self, df, path):
        self.wc_df = df
        self.wc_path = path
        self.loaded_at = datetime.now()