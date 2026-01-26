import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import datetime
import os
import re
import subprocess
import requests
import tempfile
from decimal import Decimal, ROUND_HALF_UP

# Намагаємось імпортувати бібліотеку для суми прописом
try:
    from num2words import num2words
except ImportError:
    num2words = None

# ШЛЯХ ДО ШАБЛОНІВ
TPL_DIR = "" 

# ==============================================================================
# 1. ТЕХНІЧНІ ФУНКЦІЇ
# ==============================================================================

def precise_round(number):
    return float(Decimal(str(number)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def format_num(n):
    return f"{precise_round(n):,.2f}".replace(",", " ").replace(".", ",")

def calculate_row(price_from_st, qty):
    """
    Повертає чисту ціну за одиницю (з бази) та суму рядка БЕЗ податків.
    """
    p_unit = precise_round(price_from_st)
    row_sum = precise_round(p_unit * qty)
    return p_unit, row_sum

def amount_to_text_uk(amount):
    val = precise_round(amount)
    grn = int(val)
    kop = int(round((val - grn) * 100))
    if num2words is None:
        return f"{format_num(val)} грн."
    try:
        words = num2words(grn, lang='uk').capitalize()
        return f"{words} гривень, {kop:02d} коп."
    except:
        return f"{format_num(val)} грн."

# --- ФУНКЦІЇ ФОРМАТУВАННЯ DOCX ---
def apply_font_style(run, size=12, bold=False, italic=False):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    r = run._element
    r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:ascii'), 'Times New Roman')
    r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:hAnsi'), 'Times New Roman')

def set_cell_style(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]; p.alignment = align
    run = p.add_run(str(text))
    apply_font_style(run, 12, bold, italic)

# --- ОСНОВНА ФУНКЦІЯ ЗАПОВНЕННЯ ТАБЛИЦІ ---
def fill_document_table(doc, items, is_fop, is_specification=False):
    target_table = None
    for tbl in doc.tables:
        if any("Найменування" in cell.text for cell in tbl.rows[0].cells):
            target_table = tbl
            break
    if not target_table: return 0

    total_no_tax = 0
    cols = len(target_table.columns)
    
    # Групування по категоріях
    categories = {}
    for it in items:
        cat = it['cat'].upper()
        if cat not in categories: categories[cat] = []
        categories[cat].append(it)

    for cat_name, cat_items in categories.items():
        row_cat = target_table.add_row()
        row_cat.cells[0].merge(row_cat.cells[cols-1])
        set_cell_style(row_cat.cells[0], cat_name, WD_ALIGN_PARAGRAPH.CENTER, italic=True)
        
        for it in cat_items:
            # ЦІНА ТА СУМА В РЯДКАХ ЗАВЖДИ БЕЗ ПДВ
            p_unit, row_sum = calculate_row(it['p'], it['qty'])
            total_no_tax += row_sum
            
            r = target_table.add_row()
            set_cell_style(r.cells[0], it['name'])
            if cols >= 4:
                set_cell_style(r.cells[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(r.cells[2], format_num(p_unit), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(r.cells[3], format_num(row_sum), WD_ALIGN_PARAGRAPH.RIGHT)

    # РОЗРАХУНОК ПОДАТКІВ ДЛЯ ПІДСУМКУ
    if is_fop:
        tax_amount = precise_round(total_no_tax * 0.06)
        label_tax = "Податкове навантаження 6%:"
    else:
        tax_amount = precise_round(total_no_tax * 0.20)
        label_tax = "ПДВ (20%):"
        
    grand_total = precise_round(total_no_tax + tax_amount)

    # ФОРМУВАННЯ ПІДВАЛУ ТАБЛИЦІ
    if is_fop and is_specification:
        # Для ФОП Специфікація: тільки один рядок (Загальна сума з ПДВ)
        r = target_table.add_row()
        r.cells[0].merge(r.cells[cols-2])
        set_cell_style(r.cells[0], "ЗАГАЛЬНА СУМА, грн:", WD_ALIGN_PARAGRAPH.LEFT, True)
        set_cell_style(r.cells[cols-1], format_num(grand_total), WD_ALIGN_PARAGRAPH.RIGHT, True)
    else:
        # Для ТОВ (завжди) та для ФОП (тільки в КП)
        sub_label = "РАЗОМ (без ПДВ), грн:" if not is_fop else "РАЗОМ (без навантаження), грн:"
        total_label = "ЗАГАЛЬНА СУМА з ПДВ, грн:" if not is_fop else "ЗАГАЛЬНА СУМА, грн:"
        
        f_rows = [
            (sub_label, total_no_tax, False), 
            (label_tax, tax_amount, False), 
            (total_label, grand_total, True)
        ]
        for label, val, is_bold in f_rows:
            r = target_table.add_row()
            r.cells[0].merge(r.cells[cols-2])
            set_cell_style(r.cells[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            set_cell_style(r.cells[cols-1], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)
            
    return grand_total

# ... (інші функції завантаження бази, Telegram, PDF залишаються без змін)
