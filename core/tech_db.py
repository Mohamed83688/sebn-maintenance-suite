import json
import os
import pandas as pd

class TechnicianDatabase:
    def __init__(self, active_base):
        self.tech_dir = os.path.join(active_base, "Technicians")
        os.makedirs(self.tech_dir, exist_ok=True)

    def _get_path(self, matricule):
        safe_mat = str(matricule).replace("/", "_").replace("\\", "_").strip()
        return os.path.join(self.tech_dir, f"{safe_mat}.json")

    def get_profile(self, matricule):
        path = self._get_path(matricule)
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return None

    def update_tech(self, profile_dict):
        mat = profile_dict.get('matricule')
        if not mat:
            return False
            
        path = self._get_path(mat)
        
        # Ensure minimum structure if missing
        if 'exams' not in profile_dict:
            profile_dict['exams'] = []
            
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(profile_dict, f, indent=4, ensure_ascii=False)
            return True
        except Exception:
            return False

    def get_dashboard_summary(self):
        profiles = []
        for filename in os.listdir(self.tech_dir):
            if filename.endswith(".json"):
                path = os.path.join(self.tech_dir, filename)
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        profiles.append(json.load(f))
                except Exception:
                    continue
        return profiles

    def import_from_excel(self, path):
        if not os.path.exists(path):
            return False, "Le fichier n'existe pas."
            
        try:
            df = pd.read_excel(path)
            # Require standard columns: Nom, Matricule, Equipe
            col_map = {c: str(c).strip().lower() for c in df.columns}
            df.rename(columns=col_map, inplace=True)
            
            # Identify columns
            name_col = next((c for c in df.columns if 'nom' in c), None)
            mat_col = next((c for c in df.columns if 'mat' in c), None)
            shift_col = next((c for c in df.columns if 'equip' in c or 'équipe' in c or 'shift' in c), None)
            
            if not name_col or not mat_col:
                return False, "Colonnes requises (Nom, Matricule) introuvables."
                
            count = 0
            for idx, row in df.iterrows():
                mat = str(row[mat_col]).strip()
                name = str(row[name_col]).strip()
                if not mat or not name or mat == 'nan' or name == 'nan':
                    continue
                    
                shift = str(row[shift_col]).strip() if shift_col else ""
                if shift == 'nan': shift = ""
                
                # Check if exists to preserve exams
                profile = self.get_profile(mat)
                if not profile:
                    profile = {
                        "name": name,
                        "matricule": mat,
                        "shift": shift,
                        "hire_date": "",
                        "exams": []
                    }
                else:
                    # Update basic details
                    profile['name'] = name
                    profile['shift'] = shift
                    
                self.update_tech(profile)
                count += 1
                
            return True, f"Importé {count} techniciens avec succès."
        except Exception as e:
            return False, f"Erreur lors de l'importation: {e}"
