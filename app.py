import streamlit as st
import datetime
import re
import os
from docx import Document
from io import BytesIO

# Безпечний імпорт бібліотеки перетворення чисел у текст
try:
    from num2words import num2words
except ImportError:
    st.error("Будь ласка, додайте 'num2words' у requirements.txt")

# Імпорт вашої бази
from database import EQUIPMENT_BASE  

st.set_page_config(page_title="Talo Generator", layout="wide")

# --- ДОПОМІЖНІ ФУНКЦІЇ ---
def amount_to_text_uk(amount):
    units = int(amount)
    cents = int(round((amount - units) * 100))
    try:
        words = num2words(units, lang='uk').capitalize()
    except:
        words = str(units)
    return f"{words} гривень {cents:02d} копійок"

def replace_placeholders(doc, replacements):
    """Заміна без втрати жирного шрифту (runs) [cite: 18, 19]"""
    for p in doc.paragraphs:
        for k, v in replacements.items():
            if f"{{{{{k}}}}}" in p.text:
                for run in p.runs:
                    if f"{{{{{k}}}}}" in run.text:
                        run.text = run.text.replace(f"{{{{{k}}}}}" , str(v))
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                replace_placeholders(cell, replacements)

# --- РЕКВІЗИТИ ---
VENDORS = {
    "ТОВ «ТАЛО»": {"short": "О. КРАМАРЕНКО", "inn": "45274534", "adr": "Київ, вул. Крамського, 9", "iban": "UA443052990000026004046815601"},
    "ФОП Крамаренко О.С.": {"short": "О. КРАМАРЕНКО", "inn": "3048920896", "adr": "Київ, вул. Кіото, 9", "iban": "UA423348510000000026009261015"}
}

# --- ІНТЕРФЕЙС ---
st.title("⚡ Генератор Специфікацій")

if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

col1, col2 = st.columns(2)
vendor_name = col1.selectbox("Виконавець:", list(VENDORS.keys()))
customer = col1.text_input("Замовник", "ОСББ")
address = col1.text_input("Адреса об'єкта")
kp_num = col2.text_input("№ КП/Договору", "1212-25")
date_val = col2.date_input("Дата", datetime.date.today())

st.subheader("📦 Вибір товарів")
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))

for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Товари в {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
        
        # Синхронізація (видалення неактуальних)
        current_keys = set(f"{cat}_{item}" for item in selected)
        for k in list(st.session_state.selected_items.keys()):
            if k.startswith(f"{cat}_") and k not in current_keys:
                del st.session_state.selected_items[k]
        
        for item in selected:
            key = f"{cat}_{item}"
            q_col, p_col = st.columns(2)
            qty = q_col.number_input(f"К-сть {item}", 1, 100, 1, key=f"q_{key}")
            price = p_col.number_input(f"Ціна {item}", 0, 1000000, int(EQUIPMENT_BASE[cat][item]), key=f"p_{key}")
            st.session_state.selected_items[key] = {
                "name": item, "qty": qty, "price": price, "sum": qty * price, "cat": cat
            }

# --- ГЕНЕРАЦІЯ ---
if st.session_state.selected_items:
    st.divider()
    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        
        full_date = f"{date_val.day} { {1:'січня',2:'лютого',3:'березня',4:'квітня',5:'травня',6:'червня',7:'липня',8:'серпня',9:'вересня',10:'жовтня',11:'листопада',12:'грудня'}[date_val.month]} {date_val.year} року"
        safe_cust = re.sub(r'[\\/*?:"<>|]', "", customer)
        
        # Розподіл
        hw_list = [v for v in st.session_state.selected_items.values() if "послуги" not in v["cat"].lower() and "роботи" not in v["cat"].lower()]
        sw_list = [v for v in st.session_state.selected_items.values() if v not in hw_list]

        # 1. Специфікація Поставки
        if hw_list and os.path.exists("template_postavka.docx"):
            doc_p = Document("template_postavka.docx")
            total_p = sum(i["sum"] for i in hw_list)
            info = VENDORS[vendor_name]
            
            replace_placeholders(doc_p, {
                "spec_id_postavka": f"№1 від {full_date}", "customer": customer, "address": address,
                "vendor_name": vendor_name, "total_sum_digits": f"{total_p:,}".replace(",", " "),
                "total_sum_words": amount_to_text_uk(total_p), "vendor_short_name": info["short"]
            })
            
            table = doc_p.tables[0] # Використовуємо першу таблицю [cite: 12]
            for it in hw_list:
                row = table.add_row().cells
                row[0].text, row[1].text = it['name'], str(it['qty'])
                row[2].text, row[3].text = f"{it['price']:,}", f"{it['sum']:,}"
            
            buf_p = BytesIO(); doc_p.save(buf_p)
            st.download_button("📥 Скачати Поставку", buf_p.getvalue(), f"Postavka_{safe_cust}.docx")

        # 2. Специфікація Робіт
        if sw_list and os.path.exists("template_roboti.docx"):
            doc_r = Document("template_roboti.docx")
            total_r = sum(i["sum"] for i in sw_list)
            
            replace_placeholders(doc_r, {
                "spec_id_roboti": f"№1 від {full_date}", "customer": customer,
                "total_sum_words": amount_to_text_uk(total_r), "vendor_name": vendor_name
            })
            # Виправлення тегу адреси з подвійними пробілами 
            for p in doc_r.paragraphs:
                if "{{  address }}" in p.text:
                    p.text = p.text.replace("{{  address }}", address)

            table = doc_r.tables[0] # Використовуємо першу таблицю [cite: 4]
            for it in sw_list:
                row = table.add_row().cells
                row[0].text, row[1].text = it['name'], str(it['qty'])
                row[2].text, row[3].text = f"{it['price']:,}", f"{it['sum']:,}"
            
            buf_r = BytesIO(); doc_r.save(buf_r)
            st.download_button("📥 Скачати Роботи", buf_r.getvalue(), f"Roboti_{safe_cust}.docx")
