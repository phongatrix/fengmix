import sys
import os
from pathlib import Path

# Support multiple environments (local, Vercel root, Vercel function directory)
current_dir = Path(__file__).resolve().parent

candidates = [
    current_dir.parent / "backend",
    current_dir / "backend",
    Path("/var/task/backend"),
    Path("/var/task"),
]

for p in candidates:
    if p.exists() and str(p) not in sys.path:
        sys.path.insert(0, str(p))

from main import app
