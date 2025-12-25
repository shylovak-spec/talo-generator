import streamlit as st
import pandas as pd
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Talo КП Generator", page_icon="⚡", layout="wide")

# --- ФУНКЦІЯ СКЛЕЮВАННЯ ТА ЗАМІНИ МІТОК ---
def replace_placeholders(doc, replacements):
    def process_element(element):
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in element.text:
                full_text = "".join([run.text for run in element.runs])
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(value))
                    for i, run in enumerate(element.runs):
                        if i == 0:
                            run.text = new_text
                            run.bold = False
                        else:
                            run.text = ""

    for p in doc.paragraphs:
        process_element(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_element(p)

# --- ІНТЕРФЕЙС ---
st.title("⚡ Генератор КП ТОВ «Тало»")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        customer = st.text_input("Замовник", "ОСББ Вишгородська 45")
        address = st.text_input("Адреса об'єкта", "м. Київ, вул. Вишгородська 45")
        kp_num = st.text_input("Номер КП", "1223.25POW-B")
    with col2:
        manager = st.text_input("Відповідальний", "Олексій Крамаренко")
        date_str = st.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
        phone = st.text_input("Телефон", "+380 (67) 477-17-18")

st.subheader("📝 Технічні умови")
col_l, col_r = st.columns(2)
with col_l:
    txt_intro = st.text_area("Вступ", "Відповідно до наданих даних...")
    line1 = st.text_input("Пункт 1", "Організація автономного живлення ліфтів...")
with col_r:
    line2 = st.text_input("Пункт 2", "Організація автономного живлення насосної...")
    line3 = st.text_input("Пункт 3", "Аварійне освітлення та відеонагляд;")

st.subheader("📦 Специфікація обладнання")
all_selected_data = []
categories = list(EQUIPMENT_BASE.keys())
tabs = st.tabs(categories)

for i, cat in enumerate(categories):
    with tabs[i]:
        available_items = EQUIPMENT_BASE[cat]
        selected_for_cat = st.multiselect(f"Оберіть {cat}:", list(available_items.keys()), key=f"s_{cat}")
        for item in selected_for_cat:
            c_name, c_qty, c_price, c_total = st.columns([4, 1, 2, 2])
            with c_name: st.write(f"**{item}**")
            with c_qty: qty = st.number_input(f"К-сть", min_value=1, value=1, key=f"q_{item}")
            with c_price: 
                price = available_items[item]
                st.write(f"{price:,} грн")
            with c_total:
                subtotal = price * qty
                st.write(f"**{subtotal:,} грн**")
                all_selected_data.append({"Найменування": item, "Кількість": qty, "Ціна": price, "Сума": subtotal})

# --- КНОПКА ТА ГЕНЕРАЦІЯ ---
if all_selected_data:
    st.write("---")
    total_sum = sum(item["Сума"] for item in all_selected_data)
    st.header(f"Загальна сума: {total_sum:,} грн")
    
    if st.button("🚀 Сформувати та завантажити Word"):
        try:
            doc = Document("template.docx")
            
            # 1. Заміна текстів
            replacements = {
                "customer": customer, "address": address, "kp_num": kp_num,
                "manager": manager, "date": date_str, "phone": phone,
                "txt_intro": txt_intro, "line1": line1, "line2": line2, "line3": line3
            }
            replace_placeholders(doc, replacements)
            
            # 2. Пошук та заповнення таблиці
            target_table = None
            for table in doc.tables:
                # Шукаємо таблицю, де в першому рядку є "Найменування"
                if len(table.rows) > 0 and "Найменування" in table.rows[0].cells[0].text:
                    target_table = table
                    break
            
            if target_table:
                for item in all_selected_data:
                    cells = target_table.add_row().cells
                    cells[0].text = str(item["Найменування"])
                    cells[1].text = str(item["Кількість"])
                    cells[2].text = f"{item['Ціна']:,}".replace(',', ' ')
                    cells[3].text = f"{item['Сума']:,}".replace(',', ' ')
            
            # 3. Збереження
            output = BytesIO()
            doc.save(output)
            output.seek(0)
            
            st.download_button(
                label="📥 ЗАВАНТАЖИТИ ГОТОВЕ КП",
                data=output,
                file_name=f"KP_Talo_{customer}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Помилка: {e}")
