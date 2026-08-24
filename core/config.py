import json
import os

class ConfigManager:
    def __init__(self):
        # Support SEBN_DATA_DIR environment variable for persistent storage (e.g. Railway volume at /data)
        env_data = os.environ.get("SEBN_DATA_DIR")
        if env_data:
            self.active_base = os.path.abspath(env_data)
        else:
            self.active_base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
        os.makedirs(self.active_base, exist_ok=True)
        
        self.dirs = {
            "archives": os.path.join(self.active_base, "Archives"),
            "filled": os.path.join(self.active_base, "PPE_Filled"),
            "ppe": os.path.join(self.active_base, "PPE_Templates"),
            "tech": os.path.join(self.active_base, "Technicians")
        }
        
        # Ensure directories exist
        for d in self.dirs.values():
            os.makedirs(d, exist_ok=True)
            
        # Ensure configs exist
        self._ensure_configs()

    def _ensure_configs(self):
        admin_path = os.path.join(self.active_base, "admin_config.json")
        if not os.path.exists(admin_path):
            self.save_json(admin_path, {"username": "admin", "password": "admin2026"})
            
        shift_path = os.path.join(self.active_base, "shift_config.json")
        if not os.path.exists(shift_path):
            self.save_json(shift_path, {"A": "A", "B": "B", "C": "C"})
            
        email_path = os.path.join(self.active_base, "email_config.json")
        if not os.path.exists(email_path):
            self.save_json(email_path, {"sender": "", "password": "", "recipient": ""})
            
        ppe_path = os.path.join(self.active_base, "ppe_config.json")
        if not os.path.exists(ppe_path):
            self.save_json(ppe_path, {"group_to_template": {}})

    def load_json(self, filepath, default=None):
        if default is None:
            default = {}
        if os.path.exists(filepath):
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return default

    def save_json(self, filepath, data):
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

    def get_admin_credentials(self):
        admin_path = os.path.join(self.active_base, "admin_config.json")
        return self.load_json(admin_path, {"username": "admin", "password": "admin2026"})

    def get_shift_passwords(self):
        shift_path = os.path.join(self.active_base, "shift_config.json")
        return self.load_json(shift_path, {"A": "A", "B": "B", "C": "C"})

    def get_last_excel_path(self):
        path_file = os.path.join(self.active_base, "last_excel.txt")
        if os.path.exists(path_file):
            try:
                with open(path_file, 'r', encoding='utf-8') as f:
                    p = f.read().strip()
                    if p and os.path.exists(p):
                        return p
            except Exception:
                pass
        return None

    def set_last_excel_path(self, path):
        path_file = os.path.join(self.active_base, "last_excel.txt")
        with open(path_file, 'w', encoding='utf-8') as f:
            f.write(path)

    def get_schedules_dir(self):
        """Returns the dedicated directory for PMA schedule Excel files."""
        d = os.path.join(self.active_base, "Schedules")
        os.makedirs(d, exist_ok=True)
        return d

    def get_all_excel_paths(self):
        """
        Discovers and returns all valid schedule Excel files.
        First checks data/Schedules/ (dedicated folder),
        then falls back to active_base for backward-compat (current_schedule.xlsx).
        """
        paths = []
        schedules_dir = self.get_schedules_dir()

        # Prefer files in Schedules/ subfolder
        if os.path.isdir(schedules_dir):
            for f in sorted(os.listdir(schedules_dir)):
                if f.startswith('~$') or f.startswith('.'):
                    continue
                if f.lower().endswith(('.xlsx', '.xlsm')):
                    full_p = os.path.join(schedules_dir, f)
                    if os.path.isfile(full_p) and full_p not in paths:
                        paths.append(full_p)

        # Backward-compat: also check last_excel.txt path (may point to data/ root)
        if not paths:
            last_p = self.get_last_excel_path()
            if last_p and os.path.exists(last_p) and last_p not in paths:
                paths.append(last_p)

        return paths

    def list_schedule_files(self):
        """Returns metadata list for all discovered schedule files (for admin UI)."""
        result = []
        for p in self.get_all_excel_paths():
            try:
                stat = os.stat(p)
                result.append({
                    'path':     p,
                    'filename': os.path.basename(p),
                    'size_kb':  round(stat.st_size / 1024, 1),
                })
            except Exception:
                pass
        return result

    def delete_schedule_file(self, filename):
        """Safely delete a schedule Excel file by basename."""
        schedules_dir = self.get_schedules_dir()
        target = os.path.join(schedules_dir, os.path.basename(filename))
        if os.path.isfile(target):
            os.remove(target)
            return True
        # Also check in active_base (backward-compat)
        target2 = os.path.join(self.active_base, os.path.basename(filename))
        if os.path.isfile(target2):
            os.remove(target2)
            # Update last_excel.txt if it pointed to this file
            if self.get_last_excel_path() == target2:
                # Point to first remaining file
                remaining = self.get_all_excel_paths()
                if remaining:
                    self.set_last_excel_path(remaining[0])
            return True
        return False

    def get_last_tech(self, name, mat):
        cache_path = os.path.join(self.active_base, "last_tech.json")
        return self.load_json(cache_path, {"name": "", "matricule": ""})

    def set_last_tech(self, name, mat):
        cache_path = os.path.join(self.active_base, "last_tech.json")
        self.save_json(cache_path, {"name": name, "matricule": mat})
