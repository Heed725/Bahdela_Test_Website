"""Launch the attendance app from the existing Streamlit Cloud entry point."""

from pathlib import Path
from runpy import run_path

run_path(str(Path(__file__).with_name("App.py")), run_name="__main__")
