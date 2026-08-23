import os
import pandas as pd
from datetime import datetime

def export_interventions_to_excel(db):
    interventions = db.get_all_interventions()
    if not interventions:
        return None
    
    df = pd.DataFrame(interventions)
    
    # Ensure export directory exists
    export_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'exports')
    os.makedirs(export_dir, exist_ok=True)
    
    filename = f"IMA_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = os.path.join(export_dir, filename)
    
    df.to_excel(filepath, index=False)
    return filepath
