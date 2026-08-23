import os
import time

def get_vault_files():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    archives_dir = os.path.join(base_dir, '..', 'archives')
    os.makedirs(archives_dir, exist_ok=True)
    
    files_list = []
    for f in os.listdir(archives_dir):
        path = os.path.join(archives_dir, f)
        if os.path.isfile(path):
            stat = os.stat(path)
            size_mb = stat.st_size / (1024 * 1024)
            size_str = f"{size_mb:.1f} MB" if size_mb >= 1 else f"{stat.st_size / 1024:.0f} KB"
            date_str = time.strftime('%d %b %Y', time.localtime(stat.st_mtime))
            
            ext = f.split('.')[-1].lower() if '.' in f else ''
            files_list.append({
                'name': f,
                'size': size_str,
                'date': date_str,
                'ext': ext
            })
    return files_list
