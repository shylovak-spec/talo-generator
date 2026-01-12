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

# ==============================================================================
# 1. ТЕХНІЧНІ НАЛАШТУВАННЯ ТА БЕЗПЕЧНЕ ЗАВАНТАЖЕННЯ БАЗИ
# ==============================================================================

def precise_round(number):
    return float(Decimal(str(number)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

@st.cache_data(ttl=3600)
def load_full_database_from_gsheets():
    """Завантаження бази з захистом від порожніх цінових клітинок"""
    try:
        if "gcp_service_account" not in st.secrets: return {}
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], 
               scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        gc = gspread.authorize(creds)
        sh = gc.open("База_Товарів")
        full_base = {}
        for sheet in sh.worksheets():
            category_name = sheet.title
            data = sheet.get_all_records()
            items_in_cat = {}
            for row in data:
                name = str(row.get('Назва', '')).strip()
                # Очищення ціни від пробілів та заміна коми на крапку
                price_raw = str(row.get('Ціна', '0')).replace(" ", "").replace(",", ".")
                
                # ЗАХИСТ ВІД ПОМИЛКИ 'could not convert string to float'
                try:
                    price = float(price_raw) if (price_raw and price_raw.strip() != "") else 0.0
                except ValueError:
                    price = 0.0
                
                if name: items_in_cat[name] = price
            if items_in_cat: full_base[category_name] = items_in_cat
        return full_base
    except Exception as e:
        st.sidebar.error(f"⚠️ Помилка бази: {e}")
        return {}

EQUIPMENT_BASE = load_full_database_from_gsheets()

VENDORS = {
    "ТОВ «ТАЛО»": {"full": "ТОВ «ТАЛО»", "short": "Олексій КРАМАРЕНКО", "inn": "32670939", "adr": "03113, м. Київ, проспект Перемоги, будинок 68/1 офіс 62", "iban": "_________", "bank": "АТ «УКРСИББАНК»", "tax_label": "ПДВ (20%)", "tax_rate": 0.20},
    "ФОП Крамаренко Олексій Сергійович": {"full": "ФОП Крамаренко Олексій Сергійович", "short": "Олексій КРАМАРЕНКО", "inn": "3048920896", "adr": "02156 м. Київ, вул. Кіото 9, кв. 40", "iban": "UA423348510000000026009261015", "bank": "АТ «ПУМБ»", "tax_label": "6%", "tax_rate": 0.06},
    "ФОП Шилова Ксенія Вікторівна": {"full": "ФОП Шилова Ксенія Вікторівна", "short": "Ксенія ШИЛОВА", "inn": "3237308989", "adr": "20901 м. Чигирин, вул. Миру 4, кв. 43", "iban": "UA433220010000026007350102344", "bank": "АТ УНІВЕРСАЛ БАНК", "tax_label": "6%", "tax_rate": 0.06}
}

try: from num2words import num2words
except: num2words = None

# ==============================================================================
# 2. PDF ТА ТЕЛЕГРАМ
# ==============================================================================

def docx_to_pdf_libreoffice(docx_bytes):
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_path = os.path.join(tmp_dir, "temp.docx")
        with open(input_path, "wb") as f: f.write(docx_bytes)
        try:
            subprocess.run(['lowriter', '--headless', '--convert-to', 'pdf', '--outdir', tmp_dir, input_path], check=True)
            pdf_path = os.path.join(tmp_dir, "temp.pdf")
            with open(pdf_path, "rb") as f: return f.read()
        except: return None

def send_telegram_file(file_bytes, file_name):
    token = st.secrets.get("telegram_bot_token")
    chat_id = st.secrets.get("telegram_chat_id")
    if not token or not chat_id: return
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    try:
        files = {'document': (file_name, file_bytes)}
        requests.post(url, files=files, data={'chat_id': chat_id})
        st.toast(f"✅ Telegram: {file_name}")
    except: pass

# ==============================================================================
# 3. ШРИФТИ ТА ТАБЛИЦІ
# ==============================================================================

def set_cell_style(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]; p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)
    # Захист кирилиці
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

def fill_document_table(doc, items, tax_label, tax_rate, is_fop):
    target_table = None
    for tbl in doc.tables:
        if any("Найменування" in cell.text for cell in tbl.rows[0].cells):
            target_table = tbl
            break
    if not target_table: return

    def get_cat(c):
        c = c.lower()
        if any(x in c for x in ["роботи", "послуги"]): return "РОБОТИ"
        if any(x in c for x in ["комплект", "щит", "кріплення", "матеріал", "кабель", "провід"]): return "МАТЕРІАЛИ"
        return "ОБЛАДНАННЯ"

    grouped = {"ОБЛАДНАННЯ": [], "МАТЕРІАЛИ": [], "РОБОТИ": []}
    grand_total = 0
    for it in items:
        grouped[get_cat(it['cat'])].append(it)
        grand_total += it['sum']

    cols = len(target_table.columns)
    for section in ["ОБЛАДНАННЯ", "МАТЕРІАЛИ", "РОБОТИ"]:
        if not grouped[section]: continue
        row_h = target_table.add_row()
        row_h.allow_break_across_pages = False
        set_cell_style(row_h.cells[0].merge(row_h.cells[cols-1]), section, WD_ALIGN_PARAGRAPH.CENTER, True)
        for it in grouped[section]:
            r = target_table.add_row()
            r.allow_break_across_pages = False
            cells = r.cells
            set_cell_style(cells[0], it['name'])
            if cols >= 4:
                set_cell_style(cells[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(cells[2], f"{it['p']:,.2f}".replace(",", " ").replace(".", ","), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(cells[3], f"{it['sum']:,.2f}".replace(",", " ").replace(".", ","), WD_ALIGN_PARAGRAPH.RIGHT)

    # Підсумки (ПДВ або один рядок для ФОП)
    if is_fop:
        f_row = target_table.add_row(); f_row.allow_break_across_pages = False
        f_row.cells[0].merge(f_row.cells[cols-2])
        set_cell_style(f_row.cells[0], "ЗАГАЛЬНА СУМА, грн:", WD_ALIGN_PARAGRAPH.LEFT, True)
        set_cell_style(f_row.cells[cols-1], f"{grand_total:,.2f}".replace(",", " ").replace(".", ","), WD_ALIGN_PARAGRAPH.RIGHT, True)
    else:
        pure = precise_round(grand_total / (1 + tax_rate))
        f_rows = [("РАЗОМ (без ПДВ), грн:", pure, False), (f"{tax_label}:", grand_total-pure, False), ("ЗАГАЛЬНА СУМА, грн:", grand_total, True)]
        for label, val, is_bold in f_rows:
            r = target_table.add_row(); r.allow_break_across_pages = False
            r.cells[0].merge(r.cells[cols-2])
            set_cell_style(r.cells[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            set_cell_style(r.cells[cols-1], f"{val:,.2f}".replace(",", " ").replace(".", ","), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)

# ==============================================================================
# 4. ІНТЕРФЕЙС
# ==============================================================================

st.set_page_config(page_title="Talo Generator", layout="wide")
st.title("⚡ Генератор КП")

with st.sidebar:
    st.header("⚙️ Керування")
    if st.button("🔄 Оновити базу"):
        st.cache_data.clear()
        st.rerun()

if "selected_items" not in st.session_state: st.session_state.selected_items = {}
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

st.subheader("📝 Текст для КП")
txt_intro = st.text_area("Вступний текст", "Відповідно до наданих даних пропонуємо наступне:")
tc1, tc2, tc3 = st.columns(3)
l1 = tc1.text_input("Пункт 1", "Організація автономного живлення ліфтів")
l2 = tc2.text_input("Пункт 2", "Організація автономного живлення насосної")
l3 = tc3.text_input("Пункт 3", "Аварійне освітлення та відеонагляд")

st.subheader("📦 Специфікація")
if EQUIPMENT_BASE:
    tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
    for i, cat in enumerate(EQUIPMENT_BASE.keys()):
        with tabs[i]:
            sel = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
            for name in sel:
                key = f"{cat}_{name}"
                bp = EQUIPMENT_BASE[cat][name]
                dp = precise_round(bp * 1.06) if is_fop else bp
                cn, cq, cp, cs = st.columns([4, 1, 1.5, 1.5])
                cn.write(name)
                q = cq.number_input("К-сть", 1, 500, 1, key=f"q_{key}")
                p = cp.number_input("Ціна за од.", 0.0, 1000000.0, dp, key=f"p_{key}")
                st.session_state.selected_items[key] = {"name": name, "qty": q, "p": p, "sum": precise_round(p*q), "cat": cat}

active_keys = [f"{cat}_{n}" for cat in EQUIPMENT_BASE for n in st.session_state.get(f"ms_{cat}", [])]
st.session_state.selected_items = {k: v for k, v in st.session_state.selected_items.items() if k in active_keys}
items = list(st.session_state.selected_items.values())

if items:
    total = sum(i['sum'] for i in items)
    if st.button("🚀 ЗГЕНЕРУВАТИ ТА ВІДПРАВИТИ", type="primary", use_container_width=True):
        reps = {"vendor_name": v["full"], "vendor_address": v["adr"], "vendor_inn": v["inn"], "vendor_iban": v["iban"], 
                "vendor_bank": v["bank"], "vendor_email": email, "vendor_short_name": v["short"], "customer": customer, 
                "address": address, "kp_num": kp_num, "date": date_str, "manager": manager, "phone": phone, "email": email,
                "txt_intro": txt_intro, "line1": l1, "line2": l2, "line3": l3, "spec_id_postavka": kp_num, "spec_id_roboti": kp_num,
                "total_sum_digits": f"{total:,.2f}".replace(",", " ").replace(".", ","), 
                "total_sum_words": amount_to_text_uk(total)}
        
        # Реєстр
        try:
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
            gspread.authorize(creds).open("Реєстр КП Talo").get_worksheet(0).append_row([date_str, kp_num, customer, address, vendor_choice, total, manager])
        except: pass

        results = {}
        file_map = {"КП": "template.docx", "Специфікація_ОБЛ": "template_postavka.docx", "Специфікація_РОБ": "template_roboti.docx"}
        clean_addr = re.sub(r'[^\w\s-]', '', address).replace(' ', '_')[:30]

        for label, t_file in file_map.items():
            if os.path.exists(t_file):
                doc = Document(t_file)
                for item in list(doc.paragraphs) + [cell for tbl in doc.tables for row in tbl.rows for cell in row.cells]:
                    for k, val in reps.items():
                        if f"{{{{{k}}}}}" in item.text: item.text = item.text.replace(f"{{{{{k}}}}}", str(val))
                
                it_fill = items
                if "ОБЛ" in label: it_fill = [i for i in items if "роботи" not in i["cat"].lower()]
                if "РОБ" in label: it_fill = [i for i in items if "роботи" in i["cat"].lower()]
                
                if it_fill:
                    fill_document_table(doc, it_fill, v['tax_label'], v['tax_rate'], is_fop)
                    buf = BytesIO(); doc.save(buf); buf.seek(0)
                    docx_name = f"{label}_{kp_num}_{clean_addr}.docx"
                    
                    pdf_data = docx_to_pdf_libreoffice(buf.getvalue())
                    if pdf_data: send_telegram_file(pdf_data, docx_name.replace(".docx", ".pdf"))
                    results[label] = {"name": docx_name, "data": buf}
        
        st.session_state.generated_files = results
        st.rerun()

if st.session_state.generated_files:
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(f"💾 {info['name']}", info['data'], info['name'])
