import streamlit as st
from database import EQUIPMENT_BASE
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import math

# Спробуємо імпортувати num2words
try:
    from num2words import num2words
except ImportError:
    num2words = None

# ================== НАЛАШТУВАННЯ ТА ДАНІ ==================
VENDORS = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВ «ТАЛО»",
        "inn": "45274534",
        "adr": "03115, м. Київ, вул. Крамського Івана, 9",
        "iban": "UA443052990000026004046815601",
        "tax_label": "ПДВ (20%)",
        "tax_rate": 0.20
    }
}

# ================== ГЛОБАЛЬНА ЗАМІНА ТЕГІВ ==================
def global_replace(doc, replacements):
    """Шукає теги {{tag}} всюди: в параграфах і в усіх таблицях"""
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if f"{{{{{key}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{key}}}}}", str(val))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in replacements.items():
                        if f"{{{{{key}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{key}}}}}", str(val))

# ================== ЗАПОВНЕННЯ ТАБЛИЦІ (З ГРУПУВАННЯМ) ==================
def fill_smart_table(tbl, items, tax_label, tax_rate):
    # Визначаємо категорії (як на скріні)
    groups = {
        "ОБЛАДНАННЯ": ["Інвертори Deye", "Акумулятори (АКБ)"],
        "МАТЕРІАЛИ ТА КОМПЛЕКТУЮЧІ": ["Комплектуючі та щити"],
        "ПОСЛУГИ ТА РОБОТИ": ["Послуги та Роботи"]
    }
    
    grand_pure = 0
    col_count = len(tbl.columns)

    for g_name, g_cats in groups.items():
        g_items = [it for it in items if it['cat'] in g_cats]
        if not g_items: continue

        # Рядок-заголовок групи
        row = tbl.add_row().cells
        row[0].merge(row[col_count - 1])
        row[0].text = g_name
        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[0].paragraphs[0].runs[0].bold = True

        for it in g_items:
            r = tbl.add_row().cells
            r[0].text = f"- {it['name']}"
            if col_count >= 4:
                r[1].text = str(it['qty'])
                r[2].text = f"{it['p']:,}".replace(",", " ")
                r[3].text = f"{it['sum']:,}".replace(",", " ")
            grand_pure += it['sum']

    # Підсумки
    tax_val = math.ceil(grand_pure * tax_rate)
    total = grand_pure + tax_val

    for label, val in [("РАЗОМ, грн:", grand_pure), (f"{tax_label}:", tax_val), ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", total)]:
        r = tbl.add_row().cells
        if col_count >= 4:
            r[0].merge(r[2])
            r[0].text = label
            r[3].text = f"{val:,}".replace(",", " ")
            r[3].paragraphs[0].runs[0].bold = True
        else:
            r[0].text = f"{label} {val:,}".replace(",", " ")
            
    return total

# ================== STREAMLIT ==================
st.title("⚡ ТАЛО: Генератор")

if "selected_items" not in st.session_state: st.session_state.selected_items = {}

# (Тут ваш блок вибору товарів як раніше...)

items = list(st.session_state.selected_items.values())

if items and st.button("🚀 ЗГЕНЕРУВАТИ"):
    reps = {
        "kp_num": "1223.25", 
        "customer": "ОСББ", 
        "vendor_name": "ТОВ «ТАЛО»",
        "date": "29.12.2025"
    }

    if os.path.exists("template.docx"):
        doc = Document("template.docx")
        
        # 1. Сначала заменяем теги {{kp_num}} и т.д.
        global_replace(doc, reps)
        
        # 2. Потом заполняем таблицу
        tbl = doc.tables[0] # або пошук за назвою "Найменування"
        fill_smart_table(tbl, items, "ПДВ (20%)", 0.20)
        
        buf = BytesIO()
        doc.save(buf)
        st.download_button("Завантажити КП", buf.getvalue(), "KP.docx")
