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

TPL_DIR = "" 

# ==============================================================================
# 1. ТЕХНІЧНІ ФУНКЦІЇ
# ==============================================================================

def precise_round(number):
    return float(Decimal(str(number)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def format_num(n):
    return f"{precise_round(n):,.2f}".replace(",", " ").replace(".", ",")

def calculate_row(price_from_st, qty, is_fop, is_spec):
    # ПРАВКА: Для специфікації ФОП ціна одиниці множиться на 1.06
    if is_fop and is_spec:
        p_unit = precise_round(price_from_st * 1.06)
    else:
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

@st.cache_data(ttl=3600)
def load_full_database_from_gsheets():
    try:
        if "gcp_service_account" not in st.secrets: return {}
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], 
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )
        gc = gspread.authorize(creds)
        sh = gc.open("База_Товарів")
        full_base = {}
        for sheet in sh.worksheets():
            category_name = sheet.title
            data = sheet.get_all_records()
            items_in_cat = {str(row.get('Назва', '')).strip(): float(str(row.get('Ціна', '0')).replace(" ", "").replace(",", ".")) 
                            for row in data if row.get('Назва')}
            if items_in_cat: full_base[category_name] = items_in_cat
        return full_base
    except: return {}

VENDORS = {
    "ТОВ «ТАЛО»": {"full": "ТОВ «ТАЛО»", "short": "Олексій КРАМАРЕНКО", "inn": "32670939", "adr": "03113, м. Київ, проспект Перемоги, будинок 68/1 офіс 62", "iban": "UA_________________________", "bank": "АТ «УКРСИББАНК»", "tax_label": "ПДВ (20%)", "tax_rate": 0.20},
    "ФОП Крамаренко Олексій Сергійович": {"full": "ФОП Крамаренко Олексій Сергійович", "short": "Олексій КРАМАРЕНКО", "inn": "3048920896", "adr": "02156 м. Київ, вул. Кіото 9, кв. 40", "iban": "UA423348510000000026009261015", "bank": "АТ «ПУМБ»", "tax_label": "6%", "tax_rate": 0.06},
    "ФОП Шилова Ксенія Вікторівна": {"full": "ФОП Шилова Ксенія Вікторівна", "short": "Ксенія ШИЛОВА", "inn": "3237308989", "adr": "20901 м. Чигирин, вул. Миру 4, кв. 43", "iban": "UA433220010000026007350102344", "bank": "АТ УНІВЕРСАЛ БАНК", "tax_label": "6%", "tax_rate": 0.06}
}

# ==============================================================================
# 2. ФОРМАТУВАННЯ ТА ЗАПОВНЕННЯ ТАБЛИЦІ
# ==============================================================================

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

def replace_with_formatting(doc, reps):
    for p in doc.paragraphs:
        for k, v in reps.items():
            placeholder = f"{{{{{k}}}}}"
            if placeholder in p.text:
                full_text = p.text.replace(placeholder, str(v))
                p.text = ""
                if ":" in full_text:
                    parts = full_text.split(":", 1)
                    r1 = p.add_run(parts[0] + ":")
                    apply_font_style(r1, 12, bold=True)
                    r2 = p.add_run(parts[1])
                    apply_font_style(r2, 12, bold=False)
                else:
                    r = p.add_run(full_text)
                    apply_font_style(r, 12)

def fill_document_table(doc, items, is_fop, label_name):
    target_table = None
    for tbl in doc.tables:
        if any("Найменування" in cell.text for cell in tbl.rows[0].cells):
            target_table = tbl
            break
    if not target_table: return 0

    total_sum_for_bottom = 0
    cols = len(target_table.columns)
    is_spec = "Специфікація" in label_name

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
            # Для Специфікації ФОП ціна вже включає 6%
            p_unit, row_sum = calculate_row(it['p'], it['qty'], is_fop, is_spec)
            total_sum_for_bottom += row_sum
            r = target_table.add_row()
            set_cell_style(r.cells[0], it['name'])
            if cols >= 4:
                set_cell_style(r.cells[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(r.cells[2], format_num(p_unit), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(r.cells[3], format_num(row_sum), WD_ALIGN_PARAGRAPH.RIGHT)

    if is_fop and is_spec:
        # ПОВЕРНЕННЯ: Специфікація ФОП - один рядок
        r = target_table.add_row()
        r.cells[0].merge(r.cells[cols-2])
        set_cell_style(r.cells[0], "ЗАГАЛЬНА СУМА, грн:", WD_ALIGN_PARAGRAPH.LEFT, True)
        set_cell_style(r.cells[cols-1], format_num(total_sum_for_bottom), WD_ALIGN_PARAGRAPH.RIGHT, True)
        return total_sum_for_bottom
    else:
        # КП ФОП або ТОВ (будь-що) - три рядки
        tax_rate = 0.06 if is_fop else 0.20
        # Якщо це КП ФОП, ми рахували total_sum_for_bottom на чистих цінах. Додаємо податок.
        tax_amount = precise_round(total_sum_for_bottom * tax_rate)
        grand_total = precise_round(total_sum_for_bottom + tax_amount)

        sub_label = "РАЗОМ, грн:" if is_fop else "РАЗОМ, грн:"
        tax_label = "Податкове навантаження 6%:" if is_fop else "ПДВ (20%):"
        total_label = "ЗАГАЛЬНА СУМА, грн:" if is_fop else "ЗАГАЛЬНА СУМА з ПДВ, грн:"
        
        for lab, val, bld in [(sub_label, total_sum_for_bottom, False), (tax_label, tax_amount, False), (total_label, grand_total, True)]:
            r = target_table.add_row()
            r.cells[0].merge(r.cells[cols-2])
            set_cell_style(r.cells[0], lab, WD_ALIGN_PARAGRAPH.LEFT, bld)
            set_cell_style(r.cells[cols-1], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, bld)
        return grand_total

# ==============================================================================
# 3. STREAMLIT ІНТЕРФЕЙС
# ==============================================================================

st.set_page_config(page_title="Talo Generator", layout="wide")
st.title("⚡ Генератор КП")

EQUIPMENT_BASE = load_full_database_from_gsheets()
if "generated_files" not in st.session_state: st.session_state.generated_files = None

with st.expander("📌 Основні дані", expanded=True):
    c1, c2 = st.columns(2)
    vendor_choice = c1.selectbox("Виконавець:", list(VENDORS.keys()))
    is_fop = "ФОП" in vendor_choice
    v = VENDORS[vendor_choice]
    customer = c1.text_input("Замовник", "ОСББ")
    address = c1.text_input("Адреса об'єкта", "м. Київ")
    kp_num = c2.text_input("Номер КП", "1223.25")
    manager = c2.text_input("Відповідальний", "Олексій Крамаренко")
    date_str = c2.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
    phone = c2.text_input("Телефон", "+380 (67) 477-17-18")
    email = c2.text_input("E-mail", "o.kramarenko@talo.com.ua")

st.subheader("📦 Специфікація")
items_to_generate = []
if EQUIPMENT_BASE:
    tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
    for i, cat in enumerate(EQUIPMENT_BASE.keys()):
        with tabs[i]:
            sel = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
            for name in sel:
                base_p = EQUIPMENT_BASE[cat][name]
                cn, cq, cp = st.columns([4, 1, 2])
                cn.write(f"**{name}**")
                q = cq.number_input("К-сть", 1, 500, 1, key=f"qty_{cat}_{name}")
                p = cp.number_input("Ціна (чиста)", 0.0, 1000000.0, float(base_p), key=f"prc_{cat}_{name}")
                items_to_generate.append({"name": name, "qty": q, "p": p, "cat": cat})

if items_to_generate:
    # ВИВЕДЕННЯ СУМИ
    total_pure = sum(it['p'] * it['qty'] for it in items_to_generate)
    tax_rate = 0.06 if is_fop else 0.20
    tax_val = precise_round(total_pure * tax_rate)
    total_with_tax = total_pure + tax_val
    
    c_info1, c_info2 = st.columns(2)
    with c_info1:
        st.info(f"**Для КП ({'6%' if is_fop else '20%'} ПДВ):**\n\nРазом: {format_num(total_with_tax)} грн.")
    with c_info2:
        if is_fop:
            st.success(f"**Для Специфікації (ціна + 6%):**\n\nРазом: {format_num(total_with_tax)} грн.")
        else:
            st.success(f"**Для ТОВ (ПДВ):**\n\nРазом: {format_num(total_with_tax)} грн.")

    if st.button("📄 ЗГЕНЕРУВАТИ ДОКУМЕНТИ", use_container_width=True):
        reps = {"vendor_name": v["full"], "vendor_address": v["adr"], "vendor_inn": v["inn"], "vendor_iban": v["iban"], 
                "vendor_bank": v["bank"], "vendor_email": email, "customer": customer, "address": address, "kp_num": kp_num, "date": date_str, "manager": manager, "phone": phone}
        
        results = {}
        file_map = {"КП": "template.docx", "Специфікація_ОБЛ": "template_postavka.docx", "Специфікація_РОБ": "template_roboti.docx"}
        
        for label, tpl in file_map.items():
            if os.path.exists(tpl):
                doc = Document(tpl)
                it_fill = items_to_generate
                if "ОБЛ" in label: it_fill = [i for i in items_to_generate if "роботи" not in i["cat"].lower()]
                if "РОБ" in label: it_fill = [i for i in items_to_generate if "роботи" in i["cat"].lower()]
                
                if it_fill:
                    actual_total = fill_document_table(doc, it_fill, is_fop, label)
                    reps["total_sum_digits"] = format_num(actual_total)
                    reps["total_sum_words"] = amount_to_text_uk(actual_total)
                    replace_with_formatting(doc, reps)
                    buf = BytesIO(); doc.save(buf); buf.seek(0)
                    results[label] = {"name": f"{label}_{kp_num}.docx", "data": buf}
        
        st.session_state.generated_files = results
        st.rerun()

if st.session_state.generated_files:
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(f"💾 {info['name']}", info['data'], info['name'])
