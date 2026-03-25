"""
conftest.py — pytest 共通フィクスチャ
"""
import os
import sys

# backend/ を PYTHONPATH に追加
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../../backend"))
