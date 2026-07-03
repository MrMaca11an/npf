"""
Корневой conftest.py: добавляет src/ в sys.path, чтобы pytest мог импортировать npf_recon.
"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))
