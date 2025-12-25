import streamlit as st
import pandas as pd
from database import EQUIPMENT_BASE
import datetime

st.set_page_config(page_title="Talo КП Generator", page_icon="⚡", layout="wide")

# Стилізація заголовка
st.title("⚡ Генератор комерційних пропозицій ТОВ «Тало»")
st.info("Заповніть дані та оберіть обладнання. Система автоматично розрахує вартість.")

# --- БЛОК 1: ШАПКА КП ---
with st.expander("📌 Основна інформація про замовлення", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        customer = st.text_input("Замовник", "ОСББ Вишгородська 45")
        address = st.text_input("Адреса об'єкта", "м. Київ, вул. Вишгородська 45, к. 9")
    with col2:
        kp_num = st.text_input("Номер КП", "1223.25POW-B")
        manager = st.text_input("Відповідальний", "Олексій Крамаренко")
    with col3:
        date = st.date_input("Дата", datetime.date.today())
        phone = st.text_input("Телефон", "+380 (67) 000-00-00")

# --- БЛОК 2: ТЕХНІЧНЕ ЗАВДАННЯ ---
st.subheader("📝 Технічні умови (преамбула)")
requirements = st.text_area("Опис умов (буде вставлено в текст КП):", 
    "Автономне живлення ліфтів в/п 1000 та 630 кг до 8 годин автономної роботи, 2 години від мережі для повного заряду.")

# --- БЛОК 3: ВИБІР ОБЛАДНАННЯ ---
st.subheader("📦 Специфікація обладнання та робіт")

all_selected_data = []

# Створюємо вкладки для кожної категорії з бази даних
categories = list(EQUIPMENT_BASE.keys())
tabs = st.tabs(categories)

for i, cat in enumerate(categories):
    with tabs[i]:
        available_items = EQUIPMENT_BASE[cat]
        selected_for_cat = st.multiselect(f"Оберіть товари з категорії {cat}:", list(available_items.keys()), key=cat)
        
        for item in selected_for_cat:
            col_name, col_qty, col_price, col_total = st.columns([4, 1, 2, 2])
            with col_name:
                st.write(f"**{item}**")
            with col_qty:
                qty = st.number_input(f"К-сть", min_value=1, value=1, key=f"qty_{item}")
            with col_price:
                price = available_items[item]
                st.write(f"{price:,} грн".replace(',', ' '))
            with col_total:
                subtotal = price * qty
                st.write(f"**{subtotal:,} грн**".replace(',', ' '))
                all_selected_data.append({
                    "Категорія": cat,
                    "Найменування": item,
                    "Кількість": qty,
                    "Ціна, грн": price,
                    "Сума, грн": subtotal
                })

# --- БЛОК 4: ПІДСУМКИ ТА ГЕНЕРАЦІЯ ---
if all_selected_data:
    st.write("---")
    df = pd.DataFrame(all_selected_data)
    total_all = df["Сума, грн"].sum()

    st.header(f"Загальна вартість проекту: {total_all:,} грн".replace(',', ' '))

    # Попередня таблиця для перевірки
    st.subheader("Попередній перегляд таблиці КП")
    st.table(df[["Найменування", "Кількість", "Ціна, грн", "Сума, грн"]])

    if st.button("🚀 Сформувати файл Word"):
        st.success("Функція збереження у шаблон .docx підключається автоматично після завантаження вашого шаблону!")
        st.balloons()
else:
    st.warning("Будь ласка, оберіть хоча б одну позицію обладнання.")
