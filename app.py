import streamlit as st
import pandas as pd
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Talo КП Generator", page_icon="⚡", layout="wide")

# --- ФУНКЦІЯ ЗАМІНИ ТЕКСТУ ---
def replace_text_in_docx(doc, replacements):
    # Шукаємо в параграфах
    for p in doc.paragraphs:
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))
    
    # Шукаємо в таблицях (шапка КП часто в таблиці)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in replacements.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(value))

# --- ІНТЕРФЕЙС ---
st.title("⚡ Генератор КП ТОВ «Тало»")

with st.expander("📌 Дані для заповнення", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        customer = st.text_input("Замовник", "ОСББ Назва")
        address = st.text_input("Адреса", "вул. Прикладна, 1")
        kp_num = st.text_input("Номер КП", "001-2025")
    with col2:
        manager = st.text_input("Менеджер", "Олексій Крамаренко")
        date_str = st.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
        line1 = st.text_input("Пункт 1 (ліфти)", "Організація живлення до 8 годин...")

# Специфікація (спрощена для тесту)
all_selected_data = []
selected = st.multiselect("Оберіть обладнання:", list(EQUIPMENT_BASE["Інвертори"].keys()) + list(EQUIPMENT_BASE["Акумулятори"].keys()))

for item in selected:
    all_selected_data.append({"Найменування": item, "Кількість": 1, "Сума": 100})

# --- КНОПКА ГЕНЕРАЦІЇ ---
if st.button("🚀 Сформувати Word"):
    try:
        doc = Document("template.docx")
        
        replacements = {
            "customer": customer,
            "address": address,
            "kp_num": kp_num,
            "manager": manager,
            "date": date_str,
            "line1": line1
        }
        
        replace_text_in_docx(doc, replacements)
        
        # Зберігання
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 СКАЧАТИ ГОТОВИЙ ФАЙЛ",
            data=output,
            file_name=f"KP_{customer}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.success("Файл сформовано! Перевірте завантаження.")
        
    except Exception as e:
        st.error(f"Помилка: {e}. Переконайтеся, що файл template.docx завантажений на GitHub.")
