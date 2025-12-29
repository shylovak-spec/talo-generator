import streamlit as st
import datetime
import re
import gspread
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2.service_account import Credentials
from num2words import num2words
from database import EQUIPMENT_BASE  

# Налаштування сторінки
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
        "inn": "1234567890", 
        "address": "м. Київ, вул. Прикладна 1", 
        "iban": "UA000000000000000000000000000", 
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

# ================== ФУНКЦІЇ ==================
def amount_to_text(amount):
    units = int(amount)
    cents = int(round((amount - units) * 100))
    words = num2words(units, lang='uk').capitalize()
    return f"{words} гривень {cents:02d} копійок"

def get_ukr_date(date_obj):
    months = {1:"січня", 2:"лютого", 3:"березня", 4:"квітня", 5:"травня", 6:"червня",
              7:"липня", 8:"серпня", 9:"вересня", 10:"жовтня", 11:"листопада", 12:"грудня"}
    return f"{date_obj.day} {months[date_obj.month]} {date_obj.year} року"

def replace_placeholders(doc, replacements):
    # Заміна в параграфах
    for p in doc.paragraphs:
        for key, value in replacements.items():
            if f"{{{{{key}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{key}}}}}", str(value))
    # Заміна в таблицях
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in replacements.items():
                        if f"{{{{{key}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{key}}}}}", str(value))

# ================== ІНТЕРФЕЙС ==================
st.title("⚡ Генератор КП та Специфікацій")

col1, col2 = st.columns(2)
vendor_choice = col1.selectbox("Виконавець КП:", list(VENDORS_DATA.keys()))
customer = col1.text_input("Замовник", "ОСББ")
address = col1.text_input("Адреса об'єкта")
kp_num = col2.text_input("Номер договору/КП", "1212-25")
date_val = col2.date_input("Дата документів", datetime.date.today())
manager = col2.text_input("Відповідальний", "Олексій Крамаренко")

if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

st.subheader("📦 Вибір товарів")
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Товари в {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
        for item in selected:
            key = f"{cat}_{item}"
            col_q, col_p = st.columns(2)
            qty = col_q.number_input(f"К-сть {item}", 1, 100, 1, key=f"q_{key}")
            price = col_p.number_input(f"Ціна {item}", 0, 1000000, int(EQUIPMENT_BASE[cat][item]), key=f"p_{key}")
            st.session_state.selected_items[key] = {
                "Найменування": item, "Кількість": qty, "Ціна": price, "Сума": qty * price, "Категорія": cat
            }

# ================== ГЕНЕРАЦІЯ ==================
if st.session_state.selected_items:
    st.divider()
    supplier_hw_name = vendor_choice
    if vendor_choice == "ФОП Крамаренко Олексій Сергійович":
        supplier_hw_name = st.selectbox("Хто постачає ОБЛАДНАННЯ?", ["ФОП Крамаренко Олексій Сергійович", "ФОП Шилова Ксенія Вікторівна"])

    if st.button("🚀 ЗГЕНЕРУВАТИ ДОКУМЕНТИ", type="primary", use_container_width=True):
        full_date_ukr = get_ukr_date(date_val)
        short_date = date_val.strftime("%d.%m.%y")

        # РОЗДІЛЯЄМО ТОВАРИ
        hw_items = [v for k, v in st.session_state.selected_items.items() if v["Категорія"] != "4. Послуги та Роботи"]
        work_items = [v for k, v in st.session_state.selected_items.items() if v["Категорія"] == "4. Послуги та Роботи"]

        # 1. ПОСТАВКА
        if hw_items:
            try:
                doc_p = Document("template_postavka.docx")
                info = VENDORS_DATA[supplier_hw_name]
                total = sum(i["Сума"] for i in hw_items)
                
                replace_placeholders(doc_p, {
                    "spec_id_postavka": f"№1 від {full_date_ukr} до Договору поставки №П{kp_num} від {short_date}",
                    "customer": customer, "address": address, "vendor_name": supplier_hw_name,
                    "vendor_address": info["address"], "vendor_inn": info["inn"], "vendor_iban": info["iban"],
                    "vendor_bank": info["bank"], "vendor_email": info["email"], "vendor_short_name": info["short_name"],
                    "total_sum_digits": f"{total:,}".replace(",", " "), "total_sum_words": amount_to_text(total)
                })
                
                table = doc_p.tables[0]
                for it in hw_items:
                    row = table.add_row().cells
                    row[0].text = it['Найменування']
                    row[1].text = str(it['Кількість'])
                    row[2].text = f"{it['Ціна']:,}".replace(",", " ")
                    row[3].text = f"{it['Сума']:,}".replace(",", " ")
                
                buf_p = BytesIO(); doc_p.save(buf_p)
                st.download_button(f"📥 Скачати Поставку", buf_p.getvalue(), f"Spec_Postavka_{customer}.docx")
            except Exception as e: st.error(f"Помилка Поставки: {e}")

        # 2. РОБОТИ
        if work_items:
            try:
                doc_r = Document("template_roboti.docx")
                info = VENDORS_DATA[vendor_choice]
                total = sum(i["Сума"] for i in work_items)
                
                replace_placeholders(doc_r, {
                    "spec_id_roboti": f"№1 від {full_date_ukr} до Договору підряду №Р{kp_num} від {short_date}",
                    "customer": customer, "address": address, "vendor_name": vendor_choice,
                    "vendor_short_name": info["short_name"], "total_sum_words": amount_to_text(total)
                })
                
                table = doc_r.tables[0]
                for it in work_items:
                    row = table.add_row().cells
                    row[0].text = it['Найменування']
                    row[1].text = str(it['Кількість'])
                    row[2].text = f"{it['Ціна']:,}".replace(",", " ")
                    row[3].text = f"{it['Сума']:,}".replace(",", " ")
                
                buf_r = BytesIO(); doc_r.save(buf_r)
                st.download_button(f"📥 Скачати Роботи", buf_r.getvalue(), f"Spec_Roboti_{customer}.docx")
            except Exception as e: st.error(f"Помилка Робіт: {e}")
