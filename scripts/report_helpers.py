import os
import sys
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_border(cell, **kwargs):
    """
    Set cell borders: top, bottom, left, right
    kwargs: top={"sz": 4, "val": "single", "color": "000000"}
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        r'<w:tcBorders {}><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
    )
    tcPr.append(tcBorders)

def set_table_borders(table):
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            # simple formal borders: black line top & bottom for header, subtle borders for rows
            tcPr = cell._tc.get_or_add_tcPr()
            if i == 0:
                borders = parse_xml(
                    r'<w:tcBorders {}><w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                    r'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                    r'<w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
                )
            elif i == len(table.rows) - 1:
                borders = parse_xml(
                    r'<w:tcBorders {}><w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
                    r'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                    r'<w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
                )
            else:
                borders = parse_xml(
                    r'<w:tcBorders {}><w:top w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
                    r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
                    r'<w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
                )
            tcPr.append(borders)

print("Helper definitions loaded successfully.")
