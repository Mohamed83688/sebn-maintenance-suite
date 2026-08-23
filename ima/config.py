import os

class IMAConfig:
    def __init__(self):
        self.active_base = os.path.join(os.path.expanduser('~'), 'Documents', 'SEBN-TN IMA APP')
        os.makedirs(self.active_base, exist_ok=True)
        
        self._ensure_admin_config()

    def _ensure_admin_config(self):
        admin_path = os.path.join(self.active_base, "admin_config.txt")
        if not os.path.exists(admin_path):
            self.set_admin_credentials("admin", "admin2026")

    def get_admin_credentials(self):
        admin_path = os.path.join(self.active_base, "admin_config.txt")
        if os.path.exists(admin_path):
            try:
                with open(admin_path, 'r', encoding='utf-8') as f:
                    lines = f.read().strip().split('\n')
                    if len(lines) >= 2:
                        return {"username": lines[0].strip(), "password": lines[1].strip()}
            except Exception:
                pass
        return {"username": "admin", "password": "admin2026"}

    def set_admin_credentials(self, username, password):
        admin_path = os.path.join(self.active_base, "admin_config.txt")
        with open(admin_path, 'w', encoding='utf-8') as f:
            f.write(f"{username}\n{password}\n")
