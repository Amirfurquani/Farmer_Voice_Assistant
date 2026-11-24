import os
import sys
import subprocess
from pathlib import Path

# Paths
BASE_DIR = Path(r"c:\Users\maila\Desktop\Multilingual-Agri-AI-Agent")
SCREENSHOTS_DIR = BASE_DIR / "assets" / "screenshots"
LATENCY_IMG = BASE_DIR / "assets" / "latency.svg"

DOCX_OUT = BASE_DIR / "Voice_First_Multilingual_Agricultural_Assistant_Report.docx"
HTML_OUT = BASE_DIR / "Voice_First_Multilingual_Agricultural_Assistant_Report.html"
PDF_OUT = BASE_DIR / "Voice_First_Multilingual_Agricultural_Assistant_Report.pdf"

print("Starting generation script setup...")
