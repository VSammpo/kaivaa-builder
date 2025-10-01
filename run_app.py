"""
Script de lancement de l'application Streamlit
"""

import subprocess
import sys

if __name__ == "__main__":
    subprocess.run([
        sys.executable,
        "-m",
        "streamlit",
        "run",
        "frontend/Home.py",
        "--server.port=8501",
        "--server.headless=true"
    ])