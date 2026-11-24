# -*- coding: utf-8 -*-
"""
Generate comprehensive, formal technical documentation for:
Voice-First Multilingual Agricultural Assistant (Kisan Mitra)
Produces:
  1. Voice_First_Multilingual_Agricultural_Assistant_Report.docx
  2. Voice_First_Multilingual_Agricultural_Assistant_Report.html
  3. Voice_First_Multilingual_Agricultural_Assistant_Report.pdf (via Chrome headless)
"""
import os
import sys
import subprocess
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = Path(r"c:\Users\maila\Desktop\Multilingual-Agri-AI-Agent")
SCREENSHOTS_DIR = BASE_DIR / "assets" / "screenshots"

DOCX_PATH = BASE_DIR / "Voice_First_Multilingual_Agricultural_Assistant_Report.docx"
HTML_PATH = BASE_DIR / "Voice_First_Multilingual_Agricultural_Assistant_Report.html"
PDF_PATH = BASE_DIR / "Voice_First_Multilingual_Agricultural_Assistant_Report.pdf"

print("Script initialized.")
