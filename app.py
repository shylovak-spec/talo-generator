import streamlit as st
import datetime
import re
import os
from docx import Document
from io import BytesIO

# Спробуємо імпортувати num2words, якщо ні - зробимо «заглушку»
try:
    from num2words import num2words
except ModuleNotFoundError:
    st.error("Помилка: Бібліотека 'num2words' не встановлена. Додайте її в requirements.txt")

from database import EQUIPMENT_BASE  

# Налаштування сторінки
st.set_page_config(page_title="Talo КП Generator", layout="wide", page_icon="⚡")

# ================== ДОПОМІЖНІ ФУНКЦІЇ ==================

def amount_to_text(amount):
    """Перетворення суми в текст з обробкою помилок мови"""
    units = int(amount)
    cents = int(round((amount - units) * 100))
    try:
        words = num2words(units, lang='uk').capitalize()
    except Exception:
        words = str(units) # Запасний варіант, якщо укр. мова не підтримується
    return f"{words} гривень {cents:02d} копійок"

def replace_placeholders(doc, replacements):
    """Заміна тексту без втрати жирного шрифту (через runs)"""
    def process_element(element):
        for paragraph in element.paragraphs:
            for key, value in replacements.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in paragraph.text:
                    # Шукаємо тег усередині runs
                    for run in paragraph.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, str(value))
    
    process_element(doc)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_element(cell)

# ================== БАЗА РЕКВІЗИТІВ ==================
VENDORS_DATA = {
    "ТОВ «ТАЛО»": {"short_name": "Олексій КРАМАРЕНКО", "email": "talo.energy@gmail.com", "inn": "45274534", "address": "03115, м. Київ, вул. Крамського Івана, 9", "iban": "UA443052990000026004046815601", "bank": "в АТ КБ «ПРИВАТБАНК»"},
    "ФОП Крамаренко Олексій Сергійович": {"short_name": "Олексій КРАМАРЕНКО", "email": "oleksii.kramarenko.fop@gmail.com", "inn": "3048920896", "address": "02156 м. Київ, вул. Кіото 9, кв. 40", "iban": "UA423348510000000026009261015", "bank": "в АТ «ПУМБ» м. Київ"},
    "ФОП Шилова Ксенія Вікторівна": {"short_name": "Ксенія ШИЛОВА", "email": "shilova.ksenia.fop@gmail.com", "inn": "1234567890", "address": "м. Київ, вул. Прикладна 1", "iban": "UA000000000000000000000000000", "bank": "в АТ «ПРИВАТБАНК»"}
}

# ================== ІНТЕРФЕЙС ==================
st.title("⚡ Генератор документів Talo")

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
        selected = st.multiselect(f"Товари в категорії {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
        
        # СИНХРОНІЗАЦІЯ: Очищення видалених
        current_cat_keys = set(f"{cat}_{item}" for item in selected)
        for k in list(st.session_state.selected_items.keys()):
            if k.startswith(f"{cat}_") and k not in current_cat_keys:
                del st.session_state.selected_items[k]
        
        # Додавання обраних
        for item in selected:
            key = f"{cat}_{item}"
            c_q, c_p = st.columns(2)
            qty = c_q.number_input(f"К-сть {item}", 1, 100, 1, key=f"q_{key}")
            price = c_p.number_input(f"Ціна {item}", 0, 1000000, int(EQUIPMENT_BASE[cat][item]), key=f"p_{key}")
            st.session_state.selected_items[key] = {
                "Наименування": item, "Кількість": qty, "Ціна": price, "Сума": qty * price, "Категорія": cat
            }

# Перевірочна панель (для вас)
with st.expander("🔍 Перевірка обраних товарів"):
    st.write(st.session_state.selected_items)

# ================== ГЕНЕРАЦІЯ ==================
if len(st.session_state.selected_items) > 0:
    st.divider()
    
    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        # Перевірка шаблонів [cite: 1, 10, 18]
        templates = ["template_postavka.docx", "template_roboti.docx"]
        missing = [t for t in templates if not os.path.exists(t)]
        if missing:
            st.error(f"Відсутні шаблони: {', '.join(missing)}")
            st.stop()

        safe_cust = re.sub(r'[\\/*?:"<>|]', "", customer)
        full_date = f"{date_val.day} { {1:'січня',2:'лютого',3:'березня',4:'квітня',5:'травня',6:'червня',7:'липня',8:'серпня',9:'вересня',10:'жовтня',11:'листопада',12:'грудня'}[date_val.month]} {date_val.year} року"
        
        # Розподіл товарів
        work_items = [v for v in st.session_state.selected_items.values() if "послуги" in v["Категорія"].lower() or "роботи" in v["Категорія"].lower()]
        hw_items = [v for v in st.session_state.selected_items.values() if v not in work_items]

        # 1. ГЕНЕРУЄМО ПОСТАВКУ [cite: 10, 12, 13, 17]
        if hw_items:
            doc_p = Document("template_postavka.docx")
            total_p = sum(i["Сума"] for i in hw_items)
            info = VENDORS_DATA[vendor_choice]
            
            replace_placeholders(doc_p, {
                "spec_id_postavka": f"№1 від {full_date}", "customer": customer, "address": address,
                "vendor_name": vendor_choice, "vendor_address": info["address"], "vendor_inn": info["inn"],
                "total_sum_digits": f"{total_p:,}".replace(",", " "), "total_sum_words": amount_to_text(total_p),
                "vendor_short_name": info["short_name"], "vendor_iban": info["iban"]
            })
            
            table = doc_p.tables[0] [cite: 12]
            for it in hw_items:
                row = table.add_row().cells
                row[0].text, row[1].text = it['Наименування'], str(it['Кількість'])
                row[2].text, row[3].text = f"{it['Ціна']:,}".replace(",", " "), f"{it['Сума']:,}".replace(",", " ")
            
            buf_p = BytesIO(); doc_p.save(buf_p)
            st.download_button("📥 Скачати Поставку", buf_p.getvalue(), f"Spec_Postavka_{safe_cust}.docx")

        # 2. ГЕНЕРУЄМО РОБОТИ [cite: 1, 2, 4, 5, 9]
        if work_items:
            doc_r = Document("template_roboti.docx")
            total_r = sum(i["Сума"] for i in work_items)
            info = VENDORS_DATA[vendor_choice]
            
            # Спеціальна обробка тегу з подвійними пробілами {{  address }} зі скриншота 
            replace_placeholders(doc_r, {
                "spec_id_roboti": f"№1 від {full_date}", "customer": customer, 
                "total_sum_words": amount_to_text(total_r), "vendor_name": vendor_choice,
                "vendor_short_name": info["short_name"]
            })
            for p in doc_r.paragraphs:
                if "{{  address }}" in p.text:
                    p.text = p.text.replace("{{  address }}", address)

            table = doc_r.tables[0] [cite: 4]
            for it in work_items:
                row = table.add_row().cells
                row[0].text, row[1].text = it['Наименування'], str(it['Кількість'])
                row[2].text, row[3].text = f"{it['Ціна']:,}".replace(",", " "), f"{it['Сума']:,}".replace(",", " ")
            
            buf_r = BytesIO(); doc_r.save(buf_r)
            st.download_button("📥 Скачати Роботи", buf_r.getvalue(), f"Spec_Roboti_{safe_cust}.docx")
