import os
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
CACHE_DIR = ROOT_DIR / "__pycache__"
os.environ.setdefault("PYTHONPYCACHEPREFIX", str(CACHE_DIR))

sys.path.insert(0, str(ROOT_DIR))
from app.app import main

if __name__ == "__main__":
    main()
