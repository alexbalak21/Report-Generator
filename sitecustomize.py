import os
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
CACHE_DIR = ROOT_DIR / "__pycache__"

os.environ.setdefault("PYTHONPYCACHEPREFIX", str(CACHE_DIR))
