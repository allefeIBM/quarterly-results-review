#!/usr/bin/env python3
"""
Startup helper for the Quarterly Results Review Generator.
Run:  python3 start.py
Then open:  http://localhost:5050
"""
import subprocess
import sys
import os
import webbrowser
import time
import shutil


PORT = 5050
URL = f"http://localhost:{PORT}"

print("=" * 60)
print("  Quarterly Results Review — Presentation Generator")
print("  IBM Field Marketing · 2026")
print("=" * 60)
print()

# Check Python dependencies
try:
    import flask
    import openpyxl
    import pptx
    import pytesseract
    from PIL import Image
    from lxml import etree
    from flask_cors import CORS
except ImportError as e:
    print(f"Missing dependency: {e}")
    print("Installing required packages…")
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "flask", "flask-cors", "openpyxl", "python-pptx", "lxml", "pillow", "pytesseract"])
    print()

# Check Tesseract OCR
if not shutil.which("tesseract"):
    print("Warning: Tesseract OCR was not found.")
    print("The app will still open, but automatic extraction from screenshots will be unavailable until Tesseract is installed.")
    print()

print(f"Starting server on {URL}")
print("Press Ctrl+C to stop.\n")

os.chdir(os.path.dirname(os.path.abspath(__file__)))
server = subprocess.Popen([sys.executable, "app.py"])

try:
    time.sleep(2.5)
    webbrowser.open(URL)
    server.wait()
except KeyboardInterrupt:
    server.terminate()
    raise
