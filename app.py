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

# Функція для точного заокруглення (2 знаки після коми)
def precise_round(number):
    return float(Decimal(str(number)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def load_full_database_from_gsheets():
    """Зчитує всі вкладки з таблиці База_Товарів і формує EQUIPMENT_BASE"""
    try:
        credentials_info = st.secrets["gcp_service_account"]
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=scope)
        gc = gspread.authorize(creds)
        
        sh = gc.open("База_Товарів")
        all_sheets = sh.worksheets()
        
        full_base = {}
        for sheet in all_sheets:
            category_name = sheet.title
            data = sheet.get_all_records()
            items_in_cat = {}
            for row in data:
                name = str(row.get('Назва', '')).strip()
                price_raw = str(row.get('Ціна', 0)).replace(" ", "").replace(",", ".")
                try:
                    price = float(price_raw) if price_raw else 0.0
                except:
                    price = 0.0
                if name:
                    items_in_cat[name] = price
            if items_in_cat:
                full_base[category_name] = items_in_cat
        return full_base
    except Exception as e:
        st.error(f"⚠️ Помилка завантаження бази: {e}")
        return {}

# Завантажуємо актуальну базу
EQUIPMENT_BASE = load_full_database_from_gsheets()

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
        "tax_label": "Податкове навантаження (6%)",
        "tax_rate": 0.06
    },
    "ФОП Шилова Ксенія Вікторівна": {
        "full": "ФОП Шилова Ксенія Вікторівна",
        "short": "Ксенія ШИЛОВА",
        "inn": "3237308989",
        "adr": "20901 м. Чигирин, вул. Миру 4, кв. 43",
        "iban": "UA433220010000026007350102344",
        "bank": "АТ УНІВЕРСАЛ БАНК",
        "tax_label": "Податкове навантаження (6%)",
        "tax_rate": 0.06
    }
}

# ================== ДОПОМІЖНІ ФУНКЦІЇ ==================
def format_num(n):
    return f"{precise_round(n):,.2f}".replace(",", " ").replace(".", ",")

def amount_to_text_uk(amount):
    val = precise_round(amount)
    if num2words is None: return f"{format_num(val)} грн."
    try:
        integer_part = int(val)
        words = num2words(integer_part, lang='uk').capitalize()
        return f"{words} гривень 00 копійок"
    except: return f"{format_num(val)} грн."

def set_cell_style(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

def replace_headers_styled(doc, reps):
    bold_labels = ["Комерційна пропозиція:", "Дата:", "Замовник:", "Адреса:", "Виконавець:", "Контактний телефон:", "Відповідальний:", "E-mail:"]
    all_paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)

    for p in all_paragraphs:
        for key, val in reps.items():
            if f"{{{{{key}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{key}}}}}", str(val))
        
        for label in bold_labels:
            if label in p.text:
                full_text = p.text
                p.clear()
                parts = full_text.split(label, 1)
                run_l = p.add_run(label)
                run_l.bold = True
                run_l.font.name = 'Times New Roman'
                run_l.font.size = Pt(12)
                if len(parts) > 1:
                    run_v = p.add_run(parts[1])
                    run_v.bold = False
                    run_v.font.name = 'Times New Roman'
                    run_v.font.size = Pt(12)
                break

def fill_document_table(tbl, items, tax_label, tax_rate, is_fop):
    def get_category_name(item_cat):
        c = item_cat.lower()
        if "роботи" in c or "послуги" in c: return "РОБОТИ"
        # Комплектуючі та щити тепер об'єднані в МАТЕРІАЛИ
        if any(x in c for x in ["комплект", "щит", "кріплення", "матеріал", "кабель", "провід"]): 
            return "МАТЕРІАЛИ"
        return "ОБЛАДНАННЯ"

    grouped_items = {"ОБЛАДНАННЯ": [], "МАТЕРІАЛИ": [], "РОБОТИ": []}
    grand_total = 0
    for it in items:
        cat_key = get_category_name(it['cat'])
        grouped_items[cat_key].append(it)
        grand_total += it['sum']

    sections_order = ["ОБЛАДНАННЯ", "МАТЕРІАЛИ", "РОБОТИ"]
    col_count = len(tbl.columns)

    for section in sections_order:
        sec_items = grouped_items[section]
        if not sec_items: continue
        row_h = tbl.add_row().cells
        if col_count >= 4: row_h[0].merge(row_h[col_count-1])
        p = row_h[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(section.upper())
        run.italic = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        
        for it in sec_items:
            r = tbl.add_row().cells
            set_cell_style(r[0], it['name'])
            if col_count >= 4:
                set_cell_style(r[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(r[2], format_num(it['p']), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(r[3], format_num(it['sum']), WD_ALIGN_PARAGRAPH.RIGHT)

    if is_fop:
        footer_rows = [("ЗАГАЛЬНА ВАРТІСТЬ, грн:", grand_total, True)]
    else:
        pure_sum = precise_round(grand_total / (1 + tax_rate))
        tax_val = precise_round(grand_total - pure_sum)
        footer_rows = [
            ("РАЗОМ (без ПДВ), грн:", pure_sum, False), 
            (f"{tax_label}:", tax_val, False), 
            ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", grand_total, True)
        ]

    for label, val, is_bold in footer_rows:
        row = tbl.add_row().cells
        if col_count >= 4:
            row[0].merge(row[2])
            set_cell_style(row[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            set_cell_style(row[3], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)
    return grand_total

# ================== ТЕХНІЧНІ ФУНКЦІЇ ==================
def send_to_telegram(file_data, file_name):
    try:
        token = st.secrets["telegram_bot_token"]
        chat_ids = st.secrets["telegram_chat_id"]
        if isinstance(chat_ids, str): chat_ids = [chat_ids]
        url = f"https://api.telegram.org/bot{token}/sendDocument"
        for chat_id in chat_ids:
            file_data.seek(0)
            files = {'document': (file_name, file_data)}
            requests.post(url, data={'chat_id': chat_id, 'caption': f"📄 {file_name}"}, files=files)
        st.success("✅ Файл надіслано в Telegram!")
    except Exception as e:
        st.error(f"❌ Telegram Error: {e}")

def save_to_google_sheets(row_data):
    try:
        credentials_info = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        gc = gspread.authorize(creds)
        sh = gc.open("Реєстр КП Talo")
        sh.get_worksheet(0).append_row(row_data)
        return True
    except: return False

def convert_docx_to_pdf(docx_data):
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, "temp.docx")
            with open(docx_path, "wb") as f: f.write(docx_data.getvalue())
            subprocess.run(['lowriter', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path], check=True)
            with open(os.path.join(tmpdir, "temp.pdf"), "rb") as f: return BytesIO(f.read())
    except: return None

# ================== STREAMLIT INTERFACE ==================
st.set_page_config(page_title="Talo Generator", layout="wide")
st.title("⚡ Генератор КП та Специфікацій")

if "generated_files" not in st.session_state: st.session_state.generated_files = None
if "selected_items" not in st.session_state: st.session_state.selected_items = {}

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_choice = col1.selectbox("Виконавець:", list(VENDORS.keys()))
    is_fop = "ФОП" in vendor_choice
    v = VENDORS[vendor_choice]
    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса об'єкта")
    kp_num = col2.text_input("Номер КП/Договору", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_str = col2.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
    phone = col2.text_input("Телефон", "+380 (67) 477-17-18")
    email = col2.text_input("E-mail", "o.kramarenko@talo.com.ua")

st.subheader("📝 Текст для КП")
txt_intro = st.text_area("Вступний текст", "Відповідно до наданих даних пропонуємо наступне:")
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
            base_p = float(EQUIPMENT_BASE[cat].get(name, 0))
            
            # Розрахунок при виборі ФОП: ціна * 1.06
            display_p = precise_round(base_p * 1.06) if is_fop else precise_round(base_p)
            
            col_n, col_q, col_p, col_s = st.columns([4.5, 1, 1.5, 1.5])
            col_n.markdown(f"<div style='padding-top:10px;'>{name}</div>", unsafe_allow_html=True)
            edit_qty = col_q.number_input("К-сть", 1, 1000, 1, key=f"q_{key}")
            edit_p = col_p.number_input("Ціна за од.", 0.0, 1000000.0, display_p, key=f"p_{key}")
            
            # Сума рядка
            row_sum = precise_round(edit_p * edit_qty)
            col_s.markdown(f"<div style='padding-top:10px; font-weight:bold; text-align:right;'>{format_num(row_sum)} грн</div>", unsafe_allow_html=True)
            st.session_state.selected_items[key] = {"name": name, "qty": edit_qty, "p": edit_p, "sum": row_sum, "cat": cat}

# Видалення не вибраних
current_keys = [f"{cat}_{n}" for cat in EQUIPMENT_BASE for n in st.session_state.get(f"ms_{cat}", [])]
st.session_state.selected_items = {k: v for k, v in st.session_state.selected_items.items() if k in current_keys}
all_items = list(st.session_state.selected_items.values())

if all_items:
    total_val = sum(it["sum"] for it in all_items)
    st.info(f"🚀 **ЗАГАЛЬНА СУМА: {format_num(total_val)} грн**")

    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        reps = {
            "vendor_name": v["full"], "vendor_address": v["adr"], "vendor_inn": v["inn"],
            "vendor_iban": v["iban"], "vendor_bank": v["bank"], "vendor_email": email, "vendor_short_name": v["short"],
            "customer": customer, "address": address, "kp_num": kp_num, "date": date_str,
            "manager": manager, "phone": phone, "email": email, "txt_intro": txt_intro,
            "line1": l1, "line2": l2, "line3": l3, "spec_id_postavka": kp_num,
            "total_sum_digits": format_num(total_val), "total_sum_words": amount_to_text_uk(total_val)
        }
        
        save_to_google_sheets([date_str, kp_num, customer, address, vendor_choice, total_val, manager])
        
        results = {}
        templates = {"kp": "template.docx", "p": "template_postavka.docx", "w": "template_roboti.docx"}
        for k, t_file in templates.items():
            if os.path.exists(t_file):
                doc = Document(t_file)
                replace_headers_styled(doc, reps)
                
                # Фільтруємо товари для різних специфікацій
                it_list = all_items
                if k == "p": it_list = [i for i in all_items if "роботи" not in i["cat"].lower()]
                if k == "w": it_list = [i for i in all_items if "роботи" in i["cat"].lower()]
                
                if it_list:
                    fill_document_table(doc.tables[0], it_list, v['tax_label'], v['tax_rate'], is_fop)
                    buf = BytesIO(); doc.save(buf); buf.seek(0)
                    results[k] = {"name": f"{k.upper()}_{kp_num}.docx", "data": buf}
        
        st.session_state.generated_files = results
        st.rerun()

if st.session_state.generated_files:
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(f"💾 {info['name']}", info['data'], info['name'], key=f"dl_{k}")
    
    if st.button("🚀 Надіслати КП у PDF керівнику", use_container_width=True):
        pdf = convert_docx_to_pdf(st.session_state.generated_files["kp"]["data"])
        if pdf: send_to_telegram(pdf, f"КП_{kp_num}.pdf")
