import streamlit as st
from pathlib import Path
import importlib.util, sys
page = Path("pages/product_intelligence.py")
spec = importlib.util.spec_from_file_location("product_intelligence", page)
mod  = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
