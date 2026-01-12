import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import datetime
import os
import re
from decimal import Decimal, ROUND_HALF_UP

# ==============================================================================
# 1. ТЕХНІЧНІ ФУНКЦІЇ ТА РОБОТА З ДАНИМИ
# ==============================================================================

def precise_round(number):
    """Точне округлення до 2 знаків після коми (бухгалтерське)"""
    return float(Decimal(str(number)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

@st.cache_data(ttl=3600)
def load_full_database_from_gsheets():
    """Завантаження бази товарів з Google Sheets з кешуванням"""
    try:
        if "gcp_service_account" not in st.secrets:
            st.sidebar.error("❌ Відсутні секрети gcp_service_account в Streamlit Cloud")
            return {}
        
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
        st.sidebar.error(f"⚠️ Помилка завантаження бази: {e}")
        return {}

# Константи та словники
VENDORS = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВ «ТАЛО»", "short": "Олексій КРАМАРЕНКО", "inn": "32670939",
        "adr": "03113, м. Київ, проспект Перемоги, будинок 68/1 офіс 62",
        "iban": "UA_________________________", "bank": "АТ «УКРСИББАНК»", 
        "tax_label": "ПДВ (20%)", "tax_rate": 0.20
    },
    "ФОП Крамаренко Олексій Сергійович": {
        "full": "ФОП Крамаренко Олексій Сергійович", "short": "Олексій КРАМАРЕНКО", "inn": "3048920896",
        "adr": "02156 м. Київ, вул. Кіото 9, кв. 40",
        "iban": "UA423348510000000026009261015", "bank": "АТ «ПУМБ»", 
        "tax_label": "6%", "tax_rate": 0.06
    },
    "ФОП Шилова Ксенія Вікторівна": {
        "full": "ФОП Шилова Ксенія Вікторівна", "short": "Ксенія ШИЛОВА", "inn": "3237308989",
        "adr": "20901 м. Чигирин, вул. Миру 4, кв. 43",
        "iban": "UA433220010000026007350102344", "bank": "АТ УНІВЕРСАЛ БАНК", 
        "tax_label": "6%", "tax_rate": 0.06
    }
}

try:
    from num2words import num2words
except ImportError:
    num2words = None

def format_num(n):
    """Форматування чисел: 10 000,00"""
    return f"{precise_round(n):,.2f}".replace(",", " ").replace(".", ",")

def amount_to_text_uk(amount):
    """Сума прописом"""
    val = precise_round(amount)
    if num2words is None: return f"{format_num(val)} грн."
    try:
        integer_part = int(val)
        words = num2words(integer_part, lang='uk').capitalize()
        return f"{words} гривень 00 копійок"
    except:
        return f"{format_num(val)} грн."

# ==============================================================================
# 2. ФУНКЦІЇ РОБОТИ З ДОКУМЕНТАМИ WORD
# ==============================================================================

def set_cell_style(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    """Стилізація тексту в комірці"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

def fill_document_table(doc, items, tax_label, tax_rate, is_fop):
    """Пошук правильної таблиці за словом 'Найменування' та її заповнення"""
    target_table = None
    for tbl in doc.tables:
        # Перевіряємо перший рядок на наявність ключового слова
        first_row_text = "".join([cell.text for cell in tbl.rows[0].cells])
        if "Найменування" in first_row_text:
            target_table = tbl
            break
    
    if not target_table:
        return

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
        
        # Заголовок секції
        row_h = target_table.add_row()
        row_h.allow_break_across_pages = False
        cells_h = row_h.cells
        cells_h[0].merge(cells_h[col_count-1])
        set_cell_style(cells_h[0], section, WD_ALIGN_PARAGRAPH.CENTER, True)
        
        for it in grouped[section]:
            r_row = target_table.add_row()
            r_row.allow_break_across_pages = False
            r = r_row.cells
            set_cell_style(r[0], it['name'])
            if col_count >= 4:
                set_cell_style(r[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(r[2], format_num(it['p']), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(r[3], format_num(it['sum']), WD_ALIGN_PARAGRAPH.RIGHT)

    # Підсумок (footer)
    if is_fop:
        footer_rows = [("ЗАГАЛЬНА СУМА, грн:", grand_total, True)]
    else:
        pure = precise_round(grand_total / (1 + tax_rate))
        footer_rows = [
            ("РАЗОМ (без ПДВ), грн:", pure, False),
            (f"{tax_label}:", grand_total - pure, False),
            ("ЗАГАЛЬНА СУМА, грн:", grand_total, True)
        ]

    for label, val, is_bold in footer_rows:
        f_row = target_table.add_row()
        f_row.allow_break_across_pages = False
        cells_f = f_row.cells
        cells_f[0].merge(cells_f[col_count-2])
        set_cell_style(cells_f[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
        set_cell_style(cells_f[col_count-1], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)

# ==============================================================================
# 3. ІНТЕРФЕЙС STREAMLIT
# ==============================================================================

st.set_page_config(page_title="Talo Generator v2.5", layout="wide")
st.title("⚡ Генератор КП та Специфікацій")

# Бічна панель
with st.sidebar:
    st.header("⚙️ Керування")
    if st.button("🔄 Оновити базу з Google Sheets"):
        st.cache_data.clear()
        st.rerun()
    st.write("---")
    st.info("Комплектуючі та Щити автоматично потрапляють у розділ 'МАТЕРІАЛИ'")

# Завантаження бази
EQUIPMENT_BASE = load_full_database_from_gsheets()

if "selected_items" not in st.session_state: st.session_state.selected_items = {}
if "generated_files" not in st.session_state: st.session_state.generated_files = None

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_choice = col1.selectbox("Виконавець:", list(VENDORS.keys()))
    is_fop = "ФОП" in vendor_choice
    v = VENDORS[vendor_choice]
    
    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса об'єкта", "м. Київ")
    
    kp_num = col2.text_input("Номер КП", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_val = col2.date_input("Дата", datetime.date.today())
    date_str = date_val.strftime("%d.%m.%Y")
    
    phone = col2.text_input("Телефон", "+380 (67) 477-17-18")
    email = col2.text_input("E-mail", "o.kramarenko@talo.com.ua")

# Текстові блоки для КП
st.subheader("📝 Текст для КП")
txt_intro = st.text_area("Вступний текст", "Відповідно до наданих даних пропонуємо наступне:")
c1, c2, c3 = st.columns(3)
l1 = c1.text_input("Пункт 1", "Організація автономного живлення ліфтів")
l2 = c2.text_input("Пункт 2", "Організація автономного живлення насосної")
l3 = c3.text_input("Пункт 3", "Аварійне освітлення та відеонагляд")

st.subheader("📦 Вибір обладнання")
if not EQUIPMENT_BASE:
    st.warning("База товарів порожня. Перевірте з'єднання з Google Sheets.")
else:
    tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
    for i, cat in enumerate(EQUIPMENT_BASE.keys()):
        with tabs[i]:
            selected = st.multiselect(f"Додати позиції з '{cat}':", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
            for name in selected:
                key = f"{cat}_{name}"
                base_p = float(EQUIPMENT_BASE[cat].get(name, 0))
                # Автоматична націнка 6% для ФОП
                default_p = precise_round(base_p * 1.06) if is_fop else precise_round(base_p)
                
                cn, cq, cp, cs = st.columns([4.5, 1, 1.5, 1.5])
                cn.markdown(f"<div style='padding-top:10px;'>{name}</div>", unsafe_allow_html=True)
                qty = cq.number_input("К-сть", 1, 1000, 1, key=f"qty_{key}")
                price = cp.number_input("Ціна за од.", 0.0, 1000000.0, default_p, key=f"price_{key}")
                row_total = precise_round(qty * price)
                cs.markdown(f"<div style='padding-top:10px; font-weight:bold; text-align:right;'>{format_num(row_total)} грн</div>", unsafe_allow_html=True)
                
                st.session_state.selected_items[key] = {
                    "name": name, "qty": qty, "p": price, "sum": row_total, "cat": cat
                }

# Очищення від тих, що видалили з мультиселекту
all_active_keys = []
for cat in EQUIPMENT_BASE.keys():
    for name in st.session_state.get(f"ms_{cat}", []):
        all_active_keys.append(f"{cat}_{name}")

st.session_state.selected_items = {k: v for k, v in st.session_state.selected_items.items() if k in all_active_keys}
final_items = list(st.session_state.selected_items.values())

if final_items:
    grand_total_sum = sum(it['sum'] for it in final_items)
    st.success(f"💰 ЗАГАЛЬНА СУМА: {format_num(grand_total_sum)} грн")

    if st.button("🚀 ЗГЕНЕРУВАТИ ПАКЕТ ДОКУМЕНТІВ", type="primary", use_container_width=True):
        replacements = {
            "vendor_name": v["full"], "vendor_address": v["adr"], "vendor_inn": v["inn"],
            "vendor_iban": v["iban"], "vendor_bank": v["bank"], "vendor_email": email, 
            "vendor_short_name": v["short"], "customer": customer, "address": address, 
            "kp_num": kp_num, "date": date_str, "manager": manager, "phone": phone, "email": email,
            "txt_intro": txt_intro, "line1": l1, "line2": l2, "line3": l3,
            "spec_id_postavka": kp_num, "spec_id_roboti": kp_num,
            "total_sum_digits": format_num(grand_total_sum),
            "total_sum_words": amount_to_text_uk(grand_total_sum)
        }
        
        # Логіка формування назви файлу
        safe_addr = re.sub(r'[^\w\s-]', '', address).replace(' ', '_')[:30]
        
        # Запис у реєстр
        try:
            creds_reg = Credentials.from_service_account_info(st.secrets["gcp_service_account"], 
                        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
            gc_reg = gspread.authorize(creds_reg)
            sh_reg = gc_reg.open("Реєстр КП Talo")
            sh_reg.get_worksheet(0).append_row([date_str, kp_num, customer, address, vendor_choice, grand_total_sum, manager])
        except:
            pass

        generated = {}
        file_templates = {
            "КП": "template.docx", 
            "Специфікація_ОБЛ": "template_postavka.docx", 
            "Специфікація_РОБ": "template_roboti.docx"
        }

        for label, t_name in file_templates.items():
            if os.path.exists(t_name):
                doc = Document(t_name)
                
                # Заміна тегів у параграфах
                for p in doc.paragraphs:
                    for tag, val in replacements.items():
                        if f"{{{{{tag}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{tag}}}}}", str(val))
                
                # Заміна тегів у всіх таблицях (реквізити тощо)
                for tbl in doc.tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                for tag, val in replacements.items():
                                    if f"{{{{{tag}}}}}" in p.text:
                                        p.text = p.text.replace(f"{{{{{tag}}}}}", str(val))
                
                # Фільтрація для специфікацій поставки та робіт
                items_to_use = final_items
                if "ОБЛ" in label:
                    items_to_use = [i for i in final_items if "роботи" not in i["cat"].lower()]
                elif "РОБ" in label:
                    items_to_use = [i for i in final_items if "роботи" in i["cat"].lower()]
                
                # Заповнення основної таблиці товарів
                if items_to_use:
                    fill_document_table(doc, items_to_use, v['tax_label'], v['tax_rate'], is_fop)
                    
                    output = BytesIO()
                    doc.save(output)
                    output.seek(0)
                    
                    file_name = f"{label}_{kp_num}_{safe_addr}.docx"
                    generated[label] = {"name": file_name, "data": output}
        
        st.session_state.generated_files = generated
        st.rerun()

# Блок завантаження
if st.session_state.generated_files:
    st.write("---")
    st.subheader("📥 Готові файли для завантаження:")
    cols = st.columns(len(st.session_state.generated_files))
    for i, (key, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(
            label=f"💾 {info['name']}",
            data=info['data'],
            file_name=info['name'],
            key=f"dl_{key}"
        )
