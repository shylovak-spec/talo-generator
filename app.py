import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import datetime
import os
from decimal import Decimal, ROUND_HALF_UP

# ================== 1. ТЕХНІЧНІ НАЛАШТУВАННЯ ТА КЕШУВАННЯ ==================

def precise_round(number):
    return float(Decimal(str(number)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

@st.cache_data(ttl=3600)
def load_full_database_from_gsheets():
    """Безпечне завантаження бази з кешуванням на 1 годину"""
    try:
        if "gcp_service_account" not in st.secrets: return {}
        credentials_info = st.secrets["gcp_service_account"]
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=scope)
        gc = gspread.authorize(creds)
        sh = gc.open("База_Товарів")
        full_base = {}
        for sheet in sh.worksheets():
            category_name = sheet.title
            data = sheet.get_all_records()
            items_in_cat = {}
            for row in data:
                name = str(row.get('Назва', '')).strip()
                price_raw = str(row.get('Ціна', 0)).replace(" ", "").replace(",", ".")
                try: price = float(price_raw) if price_raw else 0.0
                except: price = 0.0
                if name: items_in_cat[name] = price
            if items_in_cat: full_base[category_name] = items_in_cat
        return full_base
    except Exception as e:
        st.sidebar.warning(f"⚠️ База завантажена з помилкою (API): {e}")
        return {}

EQUIPMENT_BASE = load_full_database_from_gsheets()

try:
    from num2words import num2words
except ImportError:
    num2words = None

VENDORS = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВ «ТАЛО»", "short": "Олексій КРАМАРЕНКО", "inn": "32670939",
        "adr": "03113, м. Київ, проспект Перемоги, будинок 68/1 офіс 62",
        "iban": "_________", "bank": "АТ «УКРСИББАНК»", "tax_label": "ПДВ (20%)", "tax_rate": 0.20
    },
    "ФОП Крамаренко Олексій Сергійович": {
        "full": "ФОП Крамаренко Олексій Сергійович", "short": "Олексій КРАМАРЕНКО", "inn": "3048920896",
        "adr": "02156 м. Київ, вул. Кіото 9, кв. 40",
        "iban": "UA423348510000000026009261015", "bank": "АТ «ПУМБ»", "tax_label": "6%", "tax_rate": 0.06
    },
    "ФОП Шилова Ксенія Вікторівна": {
        "full": "ФОП Шилова Ксенія Вікторівна", "short": "Ксенія ШИЛОВА", "inn": "3237308989",
        "adr": "20901 м. Чигирин, вул. Миру 4, кв. 43",
        "iban": "UA433220010000026007350102344", "bank": "АТ УНІВЕРСАЛ БАНК", "tax_label": "6%", "tax_rate": 0.06
    }
}

# ================== 2. ДОПОМІЖНІ ФУНКЦІЇ (ФОРМАТУВАННЯ) ==================

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
    p = cell.paragraphs[0]; p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold; run.font.name = 'Times New Roman'; run.font.size = Pt(11)

def fill_document_table(doc, items, tax_label, tax_rate, is_fop):
    """Шукає правильну таблицю за словом 'Найменування' і заповнює її"""
    target_table = None
    for tbl in doc.tables:
        if "Найменування" in tbl.cell(0, 0).text:
            target_table = tbl
            break
    
    if not target_table:
        return # Якщо таблицю не знайдено, нічого не робимо, щоб не псувати файл

    def get_category_name(item_cat):
        c = item_cat.lower()
        if "роботи" in c or "послуги" in c: return "РОБОТИ"
        if any(x in c for x in ["комплект", "щит", "кріплення", "матеріал", "кабель", "провід"]): 
            return "МАТЕРІАЛИ"
        return "ОБЛАДНАННЯ"

    grouped = {"ОБЛАДНАННЯ": [], "МАТЕРІАЛИ": [], "РОБОТИ": []}
    grand_total = 0
    for it in items:
        cat_key = get_category_name(it['cat'])
        grouped[cat_key].append(it)
        grand_total += it['sum']

    col_count = len(target_table.columns)
    
    for section in ["ОБЛАДНАННЯ", "МАТЕРІАЛИ", "РОБОТИ"]:
        if not grouped[section]: continue
        
        # Додаємо заголовок секції
        new_row = target_table.add_row()
        new_row.allow_break_across_pages = False
        row_h = new_row.cells
        if col_count >= 4: row_h[0].merge(row_h[col_count-1])
        p = row_h[0].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(section); run.italic = True; run.font.name = 'Times New Roman'; run.font.size = Pt(12)
        
        for it in grouped[section]:
            r_row = target_table.add_row()
            r_row.allow_break_across_pages = False
            r = r_row.cells
            set_cell_style(r[0], it['name'])
            if col_count >= 4:
                set_cell_style(r[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(r[2], format_num(it['p']), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(r[3], format_num(it['sum']), WD_ALIGN_PARAGRAPH.RIGHT)

    # Підсумок
    if is_fop:
        footer = [("ЗАГАЛЬНА СУМА, грн:", grand_total, True)]
    else:
        pure = precise_round(grand_total / (1 + tax_rate))
        footer = [("РАЗОМ (без ПДВ), грн:", pure, False), (f"{tax_label}:", grand_total-pure, False), ("ЗАГАЛЬНА СУМА, грн:", grand_total, True)]

    for label, val, is_bold in footer:
        f_row = target_table.add_row()
        f_row.allow_break_across_pages = False
        row = f_row.cells
        if col_count >= 4:
            row[0].merge(row[2])
            set_cell_style(row[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            set_cell_style(row[3], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)

    # ЛОГІКА ПІДСУМКУ (ТІЛЬКИ ЗАГАЛЬНА СУМА ДЛЯ ФОП)
    if is_fop:
        footer = [("ЗАГАЛЬНА СУМА, грн:", grand_total, True)]
    else:
        pure = precise_round(grand_total / (1 + tax_rate))
        footer = [("РАЗОМ (без ПДВ), грн:", pure, False), (f"{tax_label}:", grand_total-pure, False), ("ЗАГАЛЬНА СУМА, грн:", grand_total, True)]

    for label, val, is_bold in footer:
        row = tbl.add_row().cells
        if col_count >= 4:
            row[0].merge(row[2])
            set_cell_style(row[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            set_cell_style(row[3], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)

# ================== 3. ІНТЕРФЕЙС STREAMLIT ==================

st.set_page_config(page_title="Talo Generator", layout="wide")
st.title("⚡ Генератор КП та Специфікацій")

if "selected_items" not in st.session_state: st.session_state.selected_items = {}
if "generated_files" not in st.session_state: st.session_state.generated_files = None

with st.sidebar:
    st.write("🔧 Керування")
    if st.button("🔄 Оновити базу з Google"):
        st.cache_data.clear()
        st.rerun()

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_choice = col1.selectbox("Виконавець:", list(VENDORS.keys()))
    is_fop = "ФОП" in vendor_choice
    v = VENDORS[vendor_choice]
    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса об'єкта")
    kp_num = col2.text_input("Номер КП", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_str = col2.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
    phone = col2.text_input("Телефон", "+380 (67) 477-17-18")
    email = col2.text_input("E-mail", "o.kramarenko@talo.com.ua")

# ТЕКСТОВІ БЛОКИ КП (ВІДНОВЛЕНО)
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
        selected = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
        for name in selected:
            key = f"{cat}_{name}"
            base_p = float(EQUIPMENT_BASE[cat].get(name, 0))
            # Автоматичний коефіцієнт 1.06 для ФОП
            def_p = precise_round(base_p * 1.06) if is_fop else precise_round(base_p)
            
            cn, cq, cp, cs = st.columns([4.5, 1, 1.5, 1.5])
            cn.markdown(f"<div style='padding-top:10px;'>{name}</div>", unsafe_allow_html=True)
            q = cq.number_input("К-сть", 1, 1000, 1, key=f"q_{key}")
            p = cp.number_input("Ціна", 0.0, 1000000.0, def_p, key=f"p_{key}")
            row_sum = precise_round(p * q)
            cs.markdown(f"<div style='padding-top:10px; font-weight:bold; text-align:right;'>{format_num(row_sum)} грн</div>", unsafe_allow_html=True)
            st.session_state.selected_items[key] = {"name": name, "qty": q, "p": p, "sum": row_sum, "cat": cat}

# Видалення неактивних
active_keys = [f"{cat}_{n}" for cat in EQUIPMENT_BASE for n in st.session_state.get(f"ms_{cat}", [])]
st.session_state.selected_items = {k: v for k, v in st.session_state.selected_items.items() if k in active_keys}
items_list = list(st.session_state.selected_items.values())

if items_list:
    total_val = sum(it["sum"] for it in items_list)
    st.info(f"🚀 **ЗАГАЛЬНА СУМА: {format_num(total_val)} грн**")

    if st.button("🚀 ЗГЕНЕРУВАТИ ДОКУМЕНТИ", type="primary", use_container_width=True):
        reps = {
            "vendor_name": v["full"], "vendor_address": v["adr"], "vendor_inn": v["inn"],
            "vendor_iban": v["iban"], "vendor_bank": v["bank"], "vendor_email": email, "vendor_short_name": v["short"],
            "customer": customer, "address": address, "kp_num": kp_num, "date": date_str,
            "manager": manager, "phone": phone, "email": email, "spec_id_postavka": kp_num, "spec_id_roboti": kp_num,
            "txt_intro": txt_intro, "line1": l1, "line2": l2, "line3": l3,
            "total_sum_digits": format_num(total_val), "total_sum_words": amount_to_text_uk(total_val)
        }
        
        # Безпечний запис у Реєстр (через try/except)
        try:
            credentials_info = st.secrets.get("gcp_service_account")
            if credentials_info:
                creds = Credentials.from_service_account_info(credentials_info, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
                gc = gspread.authorize(creds)
                sh = gc.open("Реєстр КП Talo")
                sh.get_worksheet(0).append_row([date_str, kp_num, customer, address, vendor_choice, total_val, manager])
        except: pass

        results = {}
        for k, t_file in {"kp": "template.docx", "p": "template_postavka.docx", "w": "template_roboti.docx"}.items():
            if os.path.exists(t_file):
                doc = Document(t_file)
                # Заміна тегів
                for p in list(doc.paragraphs):
                    for tag, val in reps.items():
                        if f"{{{{{tag}}}}}" in p.text: p.text = p.text.replace(f"{{{{{tag}}}}}", str(val))
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                for tag, val in reps.items():
                                    if f"{{{{{tag}}}}}" in p.text: p.text = p.text.replace(f"{{{{{tag}}}}}", str(val))
                
                # Фільтр для специфікацій
                it_to_fill = items_list
                if k == "p": it_to_fill = [i for i in items_list if "роботи" not in i["cat"].lower()]
                if k == "w": it_to_fill = [i for i in items_list if "роботи" in i["cat"].lower()]
                
                if it_to_fill:
                    fill_document_table(doc.tables[0], it_to_fill, v['tax_label'], v['tax_rate'], is_fop)
                    buf = BytesIO(); doc.save(buf); buf.seek(0)
                    results[k] = {"name": f"{k.upper()}_{kp_num}.docx", "data": buf}
        
        st.session_state.generated_files = results
        st.rerun()

if st.session_state.generated_files:
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(f"💾 {info['name']}", info['data'], info['name'], key=f"dl_{k}")
