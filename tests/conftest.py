"""Hace importable `lib.*` y los tools desde tests/ sin instalar el paquete."""
import os
import sys

TOOLS = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "tools"))
if TOOLS not in sys.path:
    sys.path.insert(0, TOOLS)
