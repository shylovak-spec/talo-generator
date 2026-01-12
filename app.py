import streamlit as st
import gspread
import requests
import subprocess
import tempfile
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import datetime
import re
import os
from decimal import Decimal, ROUND_HALF_UP

# Спробуємо імпортувати num2words для суми прописом
try:
    from num2words import num2words
except ImportError:
    num2words = None

# ================== НАЛАШТУВАННЯ ТА ДАНІ ==================
VENDORS = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВ «ТАЛО»",
        "short": "Олексій КРАМАРЕНКО",
        "inn": "32670939",
        "adr": "03113, м. Київ, проспект Перемоги, будинок 68/1 офіс 62",
        "iban": "_________",
        "bank": "АТ «УКРСИББАНК»",
        "tax_label": "ПДВ (20%)",
        "tax_rate": 0.20
    },
    "ФОП Крамаренко Олексій Сергійович": {
        "full": "ФОП Крамаренко Олексій Сергійович",
        "short": "Олексій КРАМАРЕНКО",
        "inn": "3048920896",
        "adr": "02156 м. Київ, вул. Кіото 9, кв. 40",
        "iban": "UA423348510000000026009261015",
        "bank": "АТ «ПУМБ»",
        "tax_label": "Податкове навантаження (5%)",
        "tax_rate": 0.05
    },
    "ФОП Шилова Ксенія Вікторівна": {
        "full": "ФОП Шилова Ксенія Вікторівна",
        "short": "Ксенія ШИЛОВА",
        "inn": "3237308989",
        "adr": "20901 м. Чигирин, вул. Миру 4, кв. 43",
        "iban": "UA433220010000026007350102344",
        "bank": "АТ УНІВЕРСАЛ БАНК",
        "tax_label": "Податкове навантаження (5%)",
        "tax_rate": 0.05
    }
}

# ================== ДОПОМІЖНІ ФУНКЦІЇ ==================
def precise_round(number, decimals=2):
    return float(Decimal(str(number)).quantize(Decimal('1.' + '0' * decimals), rounding=ROUND_HALF_UP))

def format_num(n):
    return f"{n:,.2f}".replace(",", " ").replace(".", ",")

def amount_to_text_uk(amount):
    val = int(precise_round(amount, 0))
    if num2words is None: return f"{format_num(amount)} грн."
    try:
        words = num2words(val, lang='uk').capitalize()
        return f"{words} гривень 00 копійок"
    except: return f"{format_num(amount)} грн."

def set_document_font(doc):
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

def set_cell_style(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

def replace_text_globally(doc, reps):
    """ Надійний метод заміни тексту у всьому документі (включаючи таблиці) """
    for key, val in reps.items():
        placeholder = f"{{{{{key}}}}}"
        for p in doc.paragraphs:
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(val))
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(val))

def find_main_table(doc):
    """ Шукає таблицю, яка містить заголовок специфікації """
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "найменування" in cell.text.lower() or "товар" in cell.text.lower():
                    return table
    return doc.tables[0] if doc.tables else None

# ================== ЛОГІКА ТАБЛИЦЬ ==================
def fill_document_table(tbl, items, tax_label, tax_rate, is_fop=False):
    if not tbl: return 0
    
    def get_category_name(item_cat):
        c = item_cat.lower()
        if "роботи" in c or "послуги" in c: return "РОБОТИ"
        if any(x in c for x in ["матеріал", "кабель", "щит", "комплект"]): return "МАТЕРІАЛИ"
        return "ОБЛАДНАННЯ"

    grouped_items = {"ОБЛАДНАННЯ": [], "МАТЕРІАЛИ": [], "РОБОТИ": []}
    grand_total = 0
    
    for it in items:
        cat_key = get_category_name(it['cat'])
        # Розрахунок суми: ціна за од (вже з 1.6) * кількість
        row_sum = precise_round(it['p'] * it['qty'])
        grand_total += row_sum
        grouped_items[cat_key].append({'name': it['name'], 'qty': it['qty'], 'p': it['p'], 'sum': row_sum})

    sections_order = ["ОБЛАДНАННЯ", "МАТЕРІАЛИ", "РОБОТИ"]
    col_count = len(tbl.columns)
    
    for section in sections_order:
        sec_items = grouped_items[section]
        if not sec_items: continue
        
        row_h = tbl.add_row().cells
        if col_count >= 4: row_h[0].merge(row_h[col_count-1])
        set_cell_style(row_h[0], section, WD_ALIGN_PARAGRAPH.CENTER, bold=True)
        
        for it in sec_items:
            r = tbl.add_row().cells
            set_cell_style(r[0], it['name'], WD_ALIGN_PARAGRAPH.LEFT)
            if col_count >= 4:
                set_cell_style(r[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(r[2], format_num(it['p']), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(r[3], format_num(it['sum']), WD_ALIGN_PARAGRAPH.RIGHT)

    if is_fop:
        footer_rows = [("ЗАГАЛЬНА ВАРТІСТЬ, грн:", precise_round(grand_total), True)]
    else:
        pure_sum = precise_round(grand_total / (1 + tax_rate))
        tax_val = precise_round(grand_total - pure_sum)
        footer_rows = [
            ("РАЗОМ (без ПДВ), грн:", pure_sum, False), 
            (f"{tax_label}:", tax_val, False), 
            ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", precise_round(grand_total), True)
        ]

    for label, val, is_bold in footer_rows:
        row = tbl.add_row().cells
        if col_count >= 4:
            row[0].merge(row[2])
            set_cell_style(row[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            set_cell_style(row[3], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)
            
    return grand_total

# ================== СЕРВІСНІ ФУНКЦІЇ ==================
@st.cache_data(ttl=600)
def load_full_database_from_gsheets():
    try:
        credentials_info = st.secrets["gcp_service_account"]
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=scope)
        gc = gspread.authorize(creds)
        sh = gc.open("База_Товарів")
        full_base = {}
        for sheet in sh.worksheets():
            data = sheet.get_all_records()
            items_in_cat = {str(row.get('Назва', '')).strip(): float(str(row.get('Ціна', 0)).replace(" ", "").replace(",", ".")) for row in data if row.get('Назва')}
            if items_in_cat: full_base[sheet.title] = items_in_cat
        return full_base
    except Exception as e:
        st.error(f"⚠️ Помилка завантаження бази: {e}"); return {}

def save_to_google_sheets(row_data):
    try:
        credentials_info = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        gc = gspread.authorize(creds)
        sh = gc.open("Реєстр КП Talo")
        sh.get_worksheet(0).append_row(row_data)
    except: pass

# ================== ІНТЕРФЕЙС STREAMLIT ==================
st.set_page_config(page_title="Talo Generator v2", layout="wide")
EQUIPMENT_BASE = load_full_database_from_gsheets()

if EQUIPMENT_BASE:
    if "generated_files" not in st.session_state: st.session_state.generated_files = None
    if "selected_items" not in st.session_state: st.session_state.selected_items = {}

    st.title("⚡ Talo Generator: Робоча версія")

    with st.expander("📌 Основна інформація", expanded=True):
        col1, col2 = st.columns(2)
        vendor_choice = col1.selectbox("Виконавець:", list(VENDORS.keys()))
        is_fop_selected = "ФОП" in vendor_choice 
        v = VENDORS[vendor_choice]
        
        customer = col1.text_input("Замовник", "ОСББ")
        address = col1.text_input("Адреса об'єкта")
        kp_num = col2.text_input("Номер КП/Специфікації", "1223.25")
        manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
        date_str = col2.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
        phone = col2.text_input("Телефон", "+380 (67) 477-17-18")
        email = col2.text_input("E-mail", "o.kramarenko@talo.com.ua")

    st.subheader("📝 Текстові блоки")
    txt_intro = st.text_area("Вступ", "Відповідно до наданих даних пропонуємо наступне:")
    c1, c2, c3 = st.columns(3)
    l1 = c1.text_input("Пункт 1", "Організація автономного живлення ліфтів")
    l2 = c2.text_input("Пункт 2", "Організація автономного живлення насосної")
    l3 = c3.text_input("Пункт 3", "Аварійне освітлення та відеонагляд")

    st.subheader("📦 Специфікація")
    tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
    for i, cat in enumerate(EQUIPMENT_BASE.keys()):
        with tabs[i]:
            selected_names = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
            for name in selected_names:
                key = f"{cat}_{name}"
                b_price = float(EQUIPMENT_BASE[cat].get(name, 0))
                # Розрахунок ціни за од. з націнкою 1.6
                display_p = precise_round(b_price * 1.6)
                
                col_n, col_q, col_p, col_s = st.columns([4, 1, 1.5, 1.5])
                col_n.write(name)
                qty = col_q.number_input("К-сть", 1, 500, 1, key=f"q_{key}")
                p = col_p.number_input("Ціна за од.", 0.0, 1000000.0, float(display_p), key=f"p_{key}")
                row_sum = precise_round(p * qty)
                col_s.write(f"**{format_num(row_sum)} грн**")
                st.session_state.selected_items[key] = {"name": name, "qty": qty, "p": p, "sum": row_sum, "cat": cat}

    # Очистка видалених зі списку
    current_keys = [f"{cat}_{n}" for cat in EQUIPMENT_BASE for n in st.session_state.get(f"ms_{cat}", [])]
    st.session_state.selected_items = {k: v for k, v in st.session_state.selected_items.items() if k in current_keys}
    all_items = list(st.session_state.selected_items.values())

    if all_items:
        total_final = sum(it["sum"] for it in all_items)
        st.info(f"💰 ЗАГАЛЬНА СУМА: {format_num(total_final)} грн")

        if st.button("🚀 ЗГЕНЕРУВАТИ ДОКУМЕНТИ", type="primary"):
            save_to_google_sheets([date_str, kp_num, customer, address, vendor_choice, total_final, manager])
            
            # Словник для заміни тексту
            reps = {
                "vendor_name": v["full"], "vendor_short_name": v["short"], "vendor_address": v["adr"],
                "vendor_inn": v["inn"], "vendor_iban": v["iban"], "vendor_bank": v["bank"],
                "vendor_email": email, "customer": customer, "address": address,
                "kp_num": kp_num, "spec_id_postavka": kp_num, "date": date_str,
                "manager": manager, "phone": phone, "txt_intro": txt_intro,
                "line1": l1, "line2": l2, "line3": l3,
                "total_sum_digits": format_num(total_final),
                "total_sum_words": amount_to_text_uk(total_final)
            }

            templates = {
                "kp": ("template.docx", f"КП_{kp_num}.docx"),
                "spec": ("template_postavka.docx", f"Специфікація_{kp_num}.docx")
            }
            
            results = {}
            for k, (t_file, out_name) in templates.items():
                if os.path.exists(t_file):
                    doc = Document(t_file)
                    replace_text_globally(doc, reps)
                    
                    target_table = find_main_table(doc)
                    if target_table:
                        # Фільтрація для специфікації (лише товари)
                        it_list = all_items
                        if k == "spec": it_list = [i for i in all_items if "роботи" not in i["cat"].lower()]
                        
                        fill_document_table(target_table, it_list, v['tax_label'], v['tax_rate'], is_fop=is_fop_selected)
                    
                    buf = BytesIO()
                    doc.save(buf)
                    buf.seek(0)
                    results[k] = {"name": out_name, "data": buf}
            
            st.session_state.generated_files = results
            st.rerun()

    if st.session_state.generated_files:
        st.divider()
        c = st.columns(len(st.session_state.generated_files))
        for i, (k, info) in enumerate(st.session_state.generated_files.items()):
            c[i].download_button(f"💾 Завантажити {info['name']}", info['data'], info['name'])
