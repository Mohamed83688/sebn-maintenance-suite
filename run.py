import sys
import os
import socket

# Ensure the root of the project is in PYTHONPATH
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from web_portal.app import app

def get_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "localhost"

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    ip = get_ip()
    print("=" * 60)
    print("  SEBN-TN Maintenance Suite — Portail Unifié")
    print("=" * 60)
    print(f"  Port:     {port}")
    print(f"  Local:    http://localhost:{port}")
    print(f"  Réseau:   http://{ip}:{port}")
    print("=" * 60)
    
    # Run the application
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
