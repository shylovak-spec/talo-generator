import streamlit as st
import datetime
import re
import os
from docx import Document
from io import BytesIO
from num2words import num2words
from database import EQUIPMENT_BASE  

# Налаштування сторінки
st.set_page_config(page_title="Talo КП Generator", layout="wide", page_icon="⚡")

# ================== ФУНКЦІЇ СИНХРОНІЗАЦІЇ ТА ОБРОБКИ ==================

def amount_to_text(amount):
    units = int(amount)
    cents = int(round((amount - units) * 100))
    try:
        words = num2words(units, lang='uk').capitalize()
    except:
        words = str(units)
    return f"{words} гривень {cents:02d} копійок"

def replace_placeholders(doc, replacements):
    """Заміна тексту зі збереженням жирного шрифту та стилів"""
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in paragraph.text:
                for run in paragraph.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders(cell, replacements)

# ================== ДАНІ ВИКОНАВЦІВ ==================
VENDORS_DATA = {
    "ТОВ «ТАЛО»": {"short_name": "Олексій КРАМАРЕНКО", "email": "talo.energy@gmail.com", "inn": "45274534", "address": "03115, м. Київ, вул. Крамського Івана, 9", "iban": "UA443052990000026004046815601", "bank": "в АТ КБ «ПРИВАТБАНК»"},
    "ФОП Крамаренко Олексій Сергійович": {"short_name": "Олексій КРАМАРЕНКО", "email": "oleksii.kramarenko.fop@gmail.com", "inn": "3048920896", "address": "02156 м. Київ, вул. Кіото 9, кв. 40", "iban": "UA423348510000000026009261015", "bank": "в АТ «ПУМБ» м. Київ"},
    "ФОП Шилова Ксенія Вікторівна": {"short_name": "Ксенія ШИЛОВА", "email": "shilova.ksenia.fop@gmail.com", "inn": "1234567890", "address": "м. Київ, вул. Прикладна 1", "iban": "UA000000000000000000000000000", "bank": "в АТ «ПРИВАТБАНК»"}
}

# ================== ІНТЕРФЕЙС ==================
st.title("⚡ Генератор КП та Специфікацій")

if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

col1, col2 = st.columns(2)
vendor_choice = col1.selectbox("Виконавець КП:", list(VENDORS_DATA.keys()))
customer = col1.text_input("Замовник", "ОСББ")
address = col1.text_input("Адреса об'єкта")
kp_num = col2.text_input("Номер договору/КП", "1212-25")
date_val = col2.date_input("Дата документів", datetime.date.today())

st.subheader("📦 Вибір товарів")
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))

for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        # Створюємо список обраних товарів
        selected = st.multiselect(f"Додати з: {cat}", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
        
        # --- КЛЮЧОВЕ ВИПРАВЛЕННЯ: СИНХРОНІЗАЦІЯ ---
        current_cat_keys = set(f"{cat}_{item}" for item in selected)
        
        # 1. Видаляємо те, що зняли в цій категорії
        for k in list(st.session_state.selected_items.keys()):
            if k.startswith(f"{cat}_") and k not in current_cat_keys:
                del st.session_state.selected_items[k]
        
        # 2. Додаємо/оновлюємо те, що обрано
        for item in selected:
            key = f"{cat}_{item}"
            col_q, col_p = st.columns([1, 1])
            qty = col_q.number_input(f"К-сть: {item}", 1, 100, 1, key=f"q_{key}")
            price = col_p.number_input(f"Ціна: {item}", 0, 1000000, int(EQUIPMENT_BASE[cat][item]), key=f"p_{key}")
            
            st.session_state.selected_items[key] = {
                "Найменування": item, "Кількість": qty, "Ціна": price, "Сума": qty * price, "Категорія": cat
            }

# --- DEBUG ПАНЕЛЬ (видалити після перевірки) ---
with st.expander("🔍 Діагностика (перевірка вибраних товарів)"):
    st.write(st.session_state.selected_items)

# ================== БЛОК ГЕНЕРАЦІЇ ==================
# Перевіряємо, чи є хоч один запис у вибраному
if len(st.session_state.selected_items) > 0:
    st.divider()
    
    # Вибір постачальника заліза (тільки якщо КП від Крамаренко)
    supplier_hw_name = vendor_choice
    if vendor_choice == "ФОП Крамаренко Олексій Сергійович":
        supplier_hw_name = st.selectbox("Хто постачає ОБЛАДНАННЯ?", ["ФОП Крамаренко Олексій Сергійович", "ФОП Шилова Ксенія Вікторівна"])

    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        
        # Розподіл (безпечний)
        hw_items = [v for v in st.session_state.selected_items.values() if "послуги" not in v["Категорія"].lower() and "роботи" not in v["Категорія"].lower()]
        work_items = [v for v in st.session_state.selected_items.values() if "послуги" in v["Категорія"].lower() or "роботи" in v["Категорія"].lower()]

        # Дати
        full_date = f"{date_val.day} { {1:'січня',2:'лютого',3:'березня',4:'квітня',5:'травня',6:'червня',7:'липня',8:'серпня',9:'вересня',10:'жовтня',11:'листопада',12:'грудня'}[date_val.month]} {date_val.year} року"
        short_date = date_val.strftime("%d.%m.%y")
        safe_cust = re.sub(r'[\\/*?:"<>|]', "", customer)

        # 1. ПОСТАВКА
        if hw_items:
            try:
                doc = Document("template_postavka.docx")
                total = sum(i["Сума"] for i in hw_items)
                info = VENDORS_DATA[supplier_hw_name]
                
                replace_placeholders(doc, {
                    "spec_id_postavka": f"№1 від {full_date} до Договору №П{kp_num} від {short_date}",
                    "customer": customer, "address": address, "vendor_name": supplier_hw_name,
                    "vendor_address": info["address"], "vendor_inn": info["inn"], "vendor_iban": info["iban"],
                    "total_sum_digits": f"{total:,}".replace(",", " "), "total_sum_words": amount_to_text(total),
                    "vendor_short_name": info["short_name"], "vendor_email": info["email"]
                })
                
                table = doc.tables[0]
                for it in hw_items:
                    row = table.add_row().cells
                    row[0].text, row[1].text = it['Найменування'], str(it['Кількість'])
                    row[2].text, row[3].text = f"{it['Ціна']:,}".replace(",", " "), f"{it['Сума']:,}".replace(",", " ")
                
                buf = BytesIO(); doc.save(buf)
                st.download_button(f"📥 Скачати Поставку", buf.getvalue(), f"Postavka_{safe_cust}.docx")
            except Exception as e: st.error(f"Помилка Поставки: {e}")

        # 2. РОБОТИ
        if work_items:
            try:
                doc = Document("template_roboti.docx")
                total = sum(i["Сума"] for i in work_items)
                info = VENDORS_DATA[vendor_choice]
                
                replace_placeholders(doc, {
                    "spec_id_roboti": f"№1 від {full_date} до Договору №Р{kp_num} від {short_date}",
                    "customer": customer, "address": address, "vendor_name": vendor_choice,
                    "total_sum_words": amount_to_text(total), "vendor_short_name": info["short_name"]
                })
                
                table = doc.tables[0]
                for it in work_items:
                    row = table.add_row().cells
                    row[0].text, row[1].text = it['Найменування'], str(it['Кількість'])
                    row[2].text, row[3].text = f"{it['Ціна']:,}".replace(",", " "), f"{it['Сума']:,}".replace(",", " ")
                
                buf = BytesIO(); doc.save(buf)
                st.download_button(f"📥 Скачати Роботи", buf.getvalue(), f"Roboti_{safe_cust}.docx")
            except Exception as e: st.error(f"Помилка Робіт: {e}")
