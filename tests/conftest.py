# -*- coding: utf-8 -*-
"""Pytest bootstrap for stable local imports.

Ensures project packages (``core``, ``gui``, ``utils``, etc.) are importable
even when tests are started from inside ``tests/``.
"""

from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
project_root_str = str(PROJECT_ROOT)
if project_root_str not in sys.path:
    sys.path.insert(0, project_root_str)
