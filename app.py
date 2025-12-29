import streamlit as st
import datetime
import re
import gspread
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2.service_account import Credentials
from num2words import num2words
from database import EQUIPMENT_BASE  # Впевнений, що цей файл у вас є

# Налаштування версії та сторінки
FORM_VERSION = "v_spec_final"
st.set_page_config(page_title="Talo КП Generator", layout="wide", page_icon="⚡")

# ================== БАЗА РЕКВІЗИТІВ ==================
VENDORS_DATA = {
    "ФОП Крамаренко Олексій Сергійович": {
        "short_name": "Олексій КРАМАРЕНКО",
        "email": "oleksii.kramarenko.fop@gmail.com",
        "inn": "3048920896",
        "address": "02156 м. Київ, вул. Кіото 9, кв. 40",
        "iban": "UA423348510000000026009261015",
        "bank": "в АТ «ПУМБ» м. Київ"
    },
    "ФОП Шилова Ксенія Вікторівна": {
        "short_name": "Ксенія ШИЛОВА",
        "email": "shilova.ksenia.fop@gmail.com",
        "inn": "1234567890", # ЗАМІНІТЬ НА РЕАЛЬНИЙ
        "address": "м. Київ, вул. Прикладна 1", # ЗАМІНІТЬ НА РЕАЛЬНУ
        "iban": "UA000000000000000000000000000", # ЗАМІНІТЬ НА РЕАЛЬНИЙ
        "bank": "в АТ «ПРИВАТБАНК»"
    },
    "ТОВ «ТАЛО»": {
        "short_name": "Олексій КРАМАРЕНКО",
        "email": "talo.energy@gmail.com",
        "inn": "45274534",
        "address": "03115, м. Київ, вул. Крамського Івана, 9",
        "iban": "UA443052990000026004046815601",
        "bank": "в АТ КБ «ПРИВАТБАНК»"
    }
}

# ================== ДОПОМІЖНІ ФУНКЦІЇ ==================

def amount_to_text(amount):
    """Перетворює число у суму прописом українською"""
    units = int(amount)
    cents = int(round((amount - units) * 100))
    words = num2words(units, lang='uk').capitalize()
    return f"{words} гривень {cents:02d} копійок"

def get_ukr_date(date_obj):
    """Форматує дату: 22 грудня 2025 року"""
    months = {
        1: "січня", 2: "лютого", 3: "березня", 4: "квітня", 5: "травня", 6: "червня",
        7: "липня", 8: "серпня", 9: "вересня", 10: "жовтня", 11: "листопада", 12: "грудня"
    }
    return f"{date_obj.day} {months[date_obj.month]} {date_obj.year} року"

def save_to_google_sheets(row_data):
    try:
        if "gcp_service_account" not in st.secrets:
            return False
        credentials_info = st.secrets["gcp_service_account"]
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=scope)
        gc = gspread.authorize(creds)
        sh = gc.open("Реєстр КП Talo")
        worksheet = sh.get_worksheet(0)
        worksheet.append_row(row_data)
        return True
    except Exception as e:
        st.error(f"Помилка Google Sheets: {e}")
        return False

def replace_placeholders(doc, replacements):
    """Універсальна функція заміни тегів у Word"""
    def process_paragraph(p):
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))

    for p in doc.paragraphs: process_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: process_paragraph(p)

# ================== ІНТЕРФЕЙС STREAMLIT ==================
st.title("⚡ Generator КП та Специфікацій")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    
    vendor_choice = col1.selectbox("Виконавець (для КП):", ["ТОВ «ТАЛО»", "ФОП Крамаренко Олексій Сергійович"])
    
    # Логіка податків та даних виконавця
    if vendor_choice == "ТОВ «ТАЛО»":
        v_display = "ТОВ «ТАЛО»"
        tax_rate, tax_label = 0.20, "ПДВ (20%)"
    else:
        v_display = "ФОП Крамаренко О.С."
        tax_rate, tax_label = 0.00, "без ПДВ"

    customer = col1.text_input("Замовник", "ОСББ Назва")
    address = col1.text_input("Адреса об'єкта", "м. Київ, вул...")
    
    kp_num = col2.text_input("Номер договору/КП", "1212-25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_val = col2.date_input("Дата документів", datetime.date.today())
    
    date_str = date_val.strftime("%d.%m.%Y")
    short_year_date = date_val.strftime("%d.%m.%y")

st.subheader("📦 Специфікація товарів")
if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Додати з: {cat}", list(EQUIPMENT_BASE[cat].keys()), key=f"sel_{cat}")
        for item in selected:
            key = f"{cat}_{item}"
            with st.container():
                cA, cB, cC, cD = st.columns([3, 0.8, 1.2, 1])
                cA.write(f"**{item}**")
                qty = cB.number_input("К-сть", min_value=1, value=1, key=f"qty_{key}")
                price = cC.number_input("Ціна", min_value=0, value=int(EQUIPMENT_BASE[cat][item]), key=f"pr_{key}")
                subtotal = qty * price
                cD.write(f"**{subtotal:,}** грн")
                st.session_state.selected_items[key] = {
                    "Найменування": item, "Кількість": qty, "Ціна": price, "Сума": subtotal, "Категорія": cat
                }

# Видалення не вибраних
all_selected_data = [v for k, v in st.session_state.selected_items.items() if any(k.endswith(x) for x in [s for s in selected])]

# ================== ГЕНЕРАЦІЯ ==================
if st.session_state.selected_items:
    st.divider()
    raw_total = sum(item["Сума"] for item in st.session_state.selected_items.values())
    tax_val = int(raw_total * tax_rate)
    final_total = raw_total + tax_val
    st.info(f"Загальна сума: **{final_total:,}** грн ({tax_label})")

    # СЕКЦІЯ СПЕЦИФІКАЦІЙ
    st.subheader("📝 Налаштування специфікацій")
    col_s1, col_s2 = st.columns(2)
    
    # Вибір постачальника обладнання
    supplier_hw_name = v_display
    if vendor_choice == "ФОП Крамаренко Олексій Сергійович":
        supplier_hw_name = col_s1.selectbox("Постачальник обладнання:", ["ФОП Крамаренко Олексій Сергійович", "ФОП Шилова Ксенія Вікторівна"])
    
    if st.button("🚀 Згенерувати КП та Специфікації", type="primary", use_container_width=True):
        
        # 1. Дані для заміни (спільні)
        full_date_ukr = get_ukr_date(date_val)
        spec_id_p = f"№1 від {full_date_ukr} до Договору поставки №П{kp_num} від {short_year_date}"
        spec_id_r = f"№1 від {full_date_ukr} до Договору підряду №Р{kp_num} від {short_year_date}"
        
        # Реквізити
        hw_v_info = VENDORS_DATA.get(supplier_hw_name, VENDORS_DATA["ТОВ «ТАЛО»"])
        work_v_info = VENDORS_DATA.get(vendor_choice, VENDORS_DATA["ТОВ «ТАЛО»"])

        # ГЕНЕРАЦІЯ ПОСТАВКИ
        hw_items = [x for x in st.session_state.selected_items.values() if x["Категорія"] != "4. Послуги та Роботи"]
        if hw_items:
            doc_p = Document("template_postavka.docx")
            p_total = sum(i["Сума"] for i in hw_items)
            p_final = p_total + int(p_total * tax_rate)
            
            replace_placeholders(doc_p, {
                "spec_id_postavka": spec_id_p,
                "customer": customer, "address": address,
                "vendor_name": supplier_hw_name,
                "vendor_address": hw_v_info["address"],
                "vendor_inn": hw_v_info["inn"],
                "vendor_iban": hw_v_info["iban"],
                "vendor_bank": hw_v_info["bank"],
                "vendor_email": hw_v_info["email"],
                "vendor_short_name": hw_v_info["short_name"],
                "total_sum_digits": f"{p_final:,}".replace(",", " "),
                "total_sum_words": amount_to_text(p_final)
            })
            # (Тут додати заповнення таблиці hw_items аналогічно вашому коду)
            
            buf_p = BytesIO()
            doc_p.save(buf_p)
            st.download_button(f"📥 Специфікація Поставки ({supplier_hw_name})", buf_p.getvalue(), f"Spec_Postavka_{customer}.docx")

        # ГЕНЕРАЦІЯ РОБІТ
        sw_items = [x for x in st.session_state.selected_items.values() if x["Категорія"] == "4. Послуги та Роботи"]
        if sw_items:
            doc_r = Document("template_roboti.docx")
            r_total = sum(i["Сума"] for i in sw_items)
            r_final = r_total + int(r_total * tax_rate)
            
            replace_placeholders(doc_r, {
                "spec_id_roboti": spec_id_r,
                "customer": customer, "address": address,
                "vendor_name": vendor_choice,
                "vendor_short_name": work_v_info["short_name"],
                "total_sum_words": amount_to_text(r_final)
            })
            # (Тут додати заповнення таблиці sw_items)
            
            buf_r = BytesIO()
            doc_r.save(buf_r)
            st.download_button(f"📥 Специфікація Робіт ({vendor_choice})", buf_r.getvalue(), f"Spec_Roboti_{customer}.docx")

        # ЗАПИС В ТАБЛИЦЮ
        save_to_google_sheets([date_str, kp_num, customer, address, final_total, manager])
        st.success("✅ Реєстр оновлено!")
