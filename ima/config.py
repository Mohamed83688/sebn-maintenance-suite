"""IMA Config — reads from JSON files, mirrors PMA ConfigManager for web use."""
import os, json, shutil, tempfile

class IMAConfig:
    def __init__(self):
        env_data = os.environ.get("SEBN_DATA_DIR")
        if env_data:
            self.base_dir = os.path.abspath(env_data)
        else:
            # Check relative to repo root: <repo>/project/data
            cand_pma_data = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'project', 'data'))
            if os.path.isdir(cand_pma_data):
                self.base_dir = cand_pma_data
            elif os.path.isdir(os.path.abspath('project/data')):
                self.base_dir = os.path.abspath('project/data')
            else:
                self.base_dir = os.path.join(os.path.expanduser("~"), "Documents", "SEBN-TN IMA APP")
        try:
            os.makedirs(self.base_dir, exist_ok=True)
        except Exception:
            pass
        self.active_base = self.base_dir
        self.admin_config_file = os.path.join(self.base_dir, "ima_admin_config.json")
        self._init_defaults()

    def _init_defaults(self):
        if not os.path.exists(self.admin_config_file):
            self._save(self.admin_config_file, {"username": "admin", "password": "SEBN-TN03"})

    def _load(self, path, default=None):
        try:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except: pass
        return default or {}

    def _save(self, path, data):
        try:
            fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), text=True)
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            shutil.move(tmp, path)
            return True
        except: return False

    def get_admin_credentials(self):
        return self._load(self.admin_config_file, {"username": "admin", "password": "SEBN-TN03"})

    def set_admin_credentials(self, username, password):
        self._save(self.admin_config_file, {"username": username, "password": password})
