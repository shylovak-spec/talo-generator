import streamlit as st
import pandas as pd
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Talo КП Generator", page_icon="⚡", layout="wide")

st.title("⚡ Генератор комерційних пропозицій ТОВ «Тало»")

# --- БЛОК 1: ШАПКА КП ---
with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        customer = st.text_input("Замовник", "ОСББ Вишгородська 45")
        address = st.text_input("Адреса об'єкта", "м. Київ, вул. Вишгородська 45")
        kp_num = st.text_input("Номер КП", "1223.25POW-B")
    with col2:
        manager = st.text_input("Відповідальний", "Олексій Крамаренко")
        date_val = st.date_input("Дата", datetime.date.today())
        date_str = date_val.strftime("%d.%m.%Y")
        phone = st.text_input("Телефон", "+380 (67) 477-17-18")

# --- БЛОК 2: ТЕХНІЧНЕ ЗАВДАННЯ (Червона текстовка) ---
st.subheader("📝 Технічні умови (Червона частина)")
col_l, col_r = st.columns(2)
with col_l:
    txt_intro = st.text_area("Вступна фраза", "Відповідно до наданих даних з метою автономного забезпечення роботи ліфта, насосної ХВП, ІТП та освітлення ліфтових холів та фоє пропонуємо наступний комплекс обладнання та робіт.")
    line1 = st.text_input("Пункт 1 (Ліфти)", "Організація автономного живлення ліфтів в/п 1000 та 630 кг до 8 годин автономної роботи, 2 години від мережі загального користування з повним зарядом батарей;")
with col_r:
    line2 = st.text_input("Пункт 2 (Насоси)", "Організація автономного живлення насосної та ІТП від 6-8 годин автономної роботи, 4 години від мережі загального користування з повним зарядом батарей;")
    line3 = st.text_input("Пункт 3 (Безпека)", "Електрозабезпечення аварійного освітлення, домофона та відеонагляду;")

# --- БЛОК 3: ВИБІР ОБЛАДНАННЯ ---
st.subheader("📦 Специфікація обладнання")
all_selected_data = []
categories = list(EQUIPMENT_BASE.keys())
tabs = st.tabs(categories)

for i, cat in enumerate(categories):
    with tabs[i]:
        available_items = EQUIPMENT_BASE[cat]
        selected_for_cat = st.multiselect(f"Оберіть {cat}:", list(available_items.keys()), key=cat)
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

# --- ФУНКЦІЯ ГЕНЕРАЦІЇ ---
def generate_docx(info):
    try:
        doc = Document("template.docx")
        # Заміна міток у всіх параграфах
        for p in doc.paragraphs:
            for key, value in info.items():
                if f"{{{{{key}}}}}" in p.text:
                    p.text = p.text.replace(f"{{{{{key}}}}}", str(value))
        
        # Заміна міток у всіх таблицях (для шапки)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in info.items():
                        if f"{{{{{key}}}}}" in cell.text:
                            cell.text = cell.text.replace(f"{{{{{key}}}}}", str(value))
        
        target_stream = BytesIO()
        doc.save(target_stream)
        target_stream.seek(0)
        return target_stream
    except Exception as e:
        st.error(f"Помилка при читанні шаблону: {e}")
        return None

# --- ФІНАЛ ---
if all_selected_data:
    st.write("---")
    total_sum = sum(item["Сума"] for item in all_selected_data)
    st.header(f"Підсумок: {total_sum:,} грн".replace(',', ' '))
    
    data_to_fill = {
        "customer": customer, "address": address, "kp_num": kp_num,
        "manager": manager, "date": date_str, "phone": phone,
        "txt_intro": txt_intro, "line1": line1, "line2": line2, "line3": line3
    }
    
    if st.button("🚀 Сформувати та завантажити Word"):
        file_data = generate_docx(data_to_fill)
        if file_data:
            st.download_button(
                label="📥 Натисніть тут для скачування",
                data=file_data,
                file_name=f"КП_Тало_{customer}_{kp_num}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
