import streamlit as st
import pandas as pd
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Talo КП Generator", page_icon="⚡", layout="wide")

# --- ФУНКЦІЯ ЗАМІНИ ТЕКСТУ В WORD ---
def replace_placeholders(doc, replacements):
    # 1. Заміна в параграфах
    for p in doc.paragraphs:
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))
    
    # 2. Заміна в таблицях (шапка часто там)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in replacements.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(value))

# --- ІНТЕРФЕЙС ПРОГРАМИ ---
st.title("⚡ Генератор комерційних пропозицій ТОВ «Тало»")

# БЛОК 1: ШАПКА
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

# БЛОК 2: ТЕХНІЧНІ УМОВИ (Червона текстовка)
st.subheader("📝 Детальні технічні умови")
col_l, col_r = st.columns(2)
with col_l:
    txt_intro = st.text_area("Вступна фраза", "Відповідно до наданих даних з метою автономного забезпечення роботи ліфта, насосної ХВП, ІТП та освітлення ліфтових холів та фоє пропонуємо наступний комплекс обладнання та робіт.")
    line1 = st.text_input("Пункт 1", "Організація автономного живлення ліфтів в/п 1000 та 630 кг до 8 годин автономної роботи...")
with col_r:
    line2 = st.text_input("Пункт 2", "Організація автономного живлення насосної та ІТП від 6-8 годин автономної роботи...")
    line3 = st.text_input("Пункт 3", "Електрозабезпечення аварійного освітлення, домофона та відеонагляду;")

# БЛОК 3: ВИБІР ОБЛАДНАННЯ (Повертаємо категорії)
st.subheader("📦 Специфікація обладнання")
all_selected_data = []
categories = list(EQUIPMENT_BASE.keys())
tabs = st.tabs(categories)

for i, cat in enumerate(categories):
    with tabs[i]:
        available_items = EQUIPMENT_BASE[cat]
        selected_for_cat = st.multiselect(f"Оберіть товари ({cat}):", list(available_items.keys()), key=f"select_{cat}")
        
        for item in selected_for_cat:
            c_name, c_qty, c_price, c_total = st.columns([4, 1, 2, 2])
            with c_name:
                st.write(f"**{item}**")
            with c_qty:
                qty = st.number_input(f"Кількість", min_value=1, value=1, key=f"qty_{item}")
            with c_price:
                price = available_items[item]
                st.write(f"{price:,} грн")
            with c_total:
                subtotal = price * qty
                st.write(f"**{subtotal:,} грн**")
                all_selected_data.append({"Найменування": item, "Кількість": qty, "Ціна": price, "Сума": subtotal})

# БЛОК 4: ФІНАЛ ТА ЗАВАНТАЖЕННЯ
if all_selected_data:
    st.write("---")
    total_sum = sum(item["Сума"] for item in all_selected_data)
    st.header(f"Загальна сума: {total_sum:,} грн")
    
    if st.button("🚀 Сформувати та завантажити Word"):
        try:
            doc = Document("template.docx")
            
            # Словник для заміни
            replacements = {
                "customer": customer,
                "address": address,
                "kp_num": kp_num,
                "manager": manager,
                "date": date_str,
                "phone": phone,
                "txt_intro": txt_intro,
                "line1": line1,
                "line2": line2,
                "line3": line3
            }
            
            replace_placeholders(doc, replacements)
            
            # Збереження у файл
            target_stream = BytesIO()
            doc.save(target_stream)
            target_stream.seek(0)
            
            st.download_button(
                label="📥 ЗАВАНТАЖИТИ ГОТОВЕ КП",
                data=target_stream,
                file_name=f"KP_Talo_{customer}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("Документ готовий!")
        except Exception as e:
            st.error(f"Помилка: {e}")
