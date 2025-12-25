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
                # Збираємо текст з усіх фрагментів (runs)
                full_text = "".join([run.text for run in element.runs])
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(value))
                    # Очищаємо фрагменти і записуємо результат у перший
                    for i, run in enumerate(element.runs):
                        if i == 0:
                            run.text = new_text
                            run.bold = False  # Текст даних завжди звичайний
                        else:
                            run.text = ""

    for p in doc.paragraphs:
        process_element(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_element(p)

# --- ІНТЕРФЕЙС STREAMLIT ---
st.title("⚡ Генератор КП ТОВ «Тало»")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        customer = st.text_input("Замовник", "ОСББ Вишгородська 45")
        address = st.text_input("Адреса об'єкта", "м. Київ, вул. Вишгородська 45")
        kp_num = st.text_input("Номер КП", "1223.25POW-B")
        # ВИБІР ПОДАТКУ
        tax_type = st.radio(
            "Оберіть систему оподаткування:",
            ["ПДВ (20%)", "Податкове навантаження (6%)", "Без податку"],
            horizontal=True
        )
    with col2:
        manager = st.text_input("Відповідальний", "Олексій Крамаренко")
        date_str = st.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
        phone = st.text_input("Телефон", "+380 (67) 477-17-18")

st.subheader("📝 Технічні умови")
col_l, col_r = st.columns(2)
with col_l:
    txt_intro = st.text_area("Вступ", "Відповідно до наданих даних з метою автономного забезпечення роботи ліфта, насосної ХВП, ІТП та освітлення ліфтових холів та фоє пропонуємо наступний комплекс обладнання та робіт.")
    line1 = st.text_input("Пункт 1", "Організація автономного живлення ліфтів в/п 1000 та 630 кг до 8 годин автономної роботи...")
with col_r:
    line2 = st.text_input("Пункт 2", "Організація автономного живлення насосної та ІТП від 6-8 годин автономної роботи...")
    line3 = st.text_input("Пункт 3", "Електрозабезпечення аварійного освітлення, домофона та відеонагляду;")

st.subheader("📦 Специфікація обладнання")
all_selected_data = []
categories = list(EQUIPMENT_BASE.keys())
tabs = st.tabs(categories)

for i, cat in enumerate(categories):
    with tabs[i]:
        available_items = EQUIPMENT_BASE[cat]
        selected_for_cat = st.multiselect(f"Оберіть товари ({cat}):", list(available_items.keys()), key=f"s_{cat}")
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

# --- ЛОГІКА ГЕНЕРАЦІЇ ПРИ НАТИСКАННІ КНОПКИ ---
if all_selected_data:
    st.write("---")
    
    # Попередній розрахунок для відображення на сайті
    raw_total = sum(item["Сума"] for item in all_selected_data)
    
    if tax_type == "ПДВ (20%)":
        t_rate, t_label = 0.20, "ПДВ (20%)"
    elif tax_type == "Податкове навантаження (6%)":
        t_rate, t_label = 0.06, "Податкове навантаження (6%)"
    else:
        t_rate, t_label = 0, "Без податку"
    
    tax_val = raw_total * t_rate
    final_total = raw_total + tax_val
    
    st.write(f"Сума без податку: {raw_total:,} грн")
    st.write(f"{t_label}: {tax_val:,} грн")
    st.header(f"Усього до сплати: {final_total:,} грн")
    
    if st.button("🚀 Сформувати та завантажити Word"):
        try:
            doc = Document("template.docx")
            
            # 1. Заміна статичних текстів
            info = {
                "customer": customer, "address": address, "kp_num": kp_num,
                "manager": manager, "date": date_str, "phone": phone,
                "txt_intro": txt_intro, "line1": line1, "line2": line2, "line3": line3
            }
            replace_placeholders(doc, info)
            
            # 2. Пошук таблиці та заповнення
            target_table = None
            for table in doc.tables:
                if len(table.rows) > 0 and "Найменування" in table.rows[0].cells[0].text:
                    target_table = table
                    break
            
            if target_table:
                # Визначаємо категорії для розділення
                sections = {
                    "Обладнання": ["1. Інвертори Deye", "2. Акумулятори (АКБ)"],
                    "Матеріали": ["3. Комплектуючі та щити"],
                    "Роботи": ["4. Послуги та Роботи"]
                }

                for section_name, base_cats in sections.items():
                    # Відфільтровуємо товари, що належать до поточної секції
                    section_items = [item for item in all_selected_data if any(cat in item["Найменування"] or cat in EQUIPMENT_BASE and item["Найменування"] in EQUIPMENT_BASE[cat] for cat in base_cats)]
                    
                    if section_items:
                        # Додаємо заголовок розділу (наприклад, "ОБЛАДНАННЯ")
                        row_head = target_table.add_row().cells
                        row_head[0].text = section_name.upper()
                        row_head[0].paragraphs[0].runs[0].bold = True
                        
                        # Додаємо товари цього розділу
                        for item in section_items:
                            cells = target_table.add_row().cells
                            cells[0].text = f" - {item['Найменування']}"
                            cells[1].text = str(item["Кількість"])
                            cells[2].text = f"{item['Ціна']:,}".replace(',', ' ')
                            cells[3].text = f"{item['Сума']:,}".replace(',', ' ')

                # --- ПІДСУМКИ (нижче таблиці з товарами) ---
                # Додаємо порожній рядок для візуального розділення
                target_table.add_row()

                # Рядок РАЗОМ
                row_sum = target_table.add_row().cells
                row_sum[0].text = "РАЗОМ (без податку):"
                row_sum[3].text = f"{raw_total:,}".replace(',', ' ')
                row_sum[0].paragraphs[0].runs[0].bold = True

                # Рядок ПОДАТКУ
                if t_rate > 0:
                    row_tax = target_table.add_row().cells
                    row_tax[0].text = f"{t_label}:"
                    row_tax[3].text = f"{tax_val:,}".replace(',', ' ')

                # Рядок ЗАГАЛЬНА ВАРТІСТЬ
                row_final = target_table.add_row().cells
                row_final[0].text = "ЗАГАЛЬНА ВАРТІСТЬ З ПОДАТКОМ:"
                row_final[3].text = f"{final_total:,}".replace(',', ' ')
                for cell in row_final:
                    if cell.text:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.bold = True
            
            # 3. Збереження файлу
            output = BytesIO()
            doc.save(output)
            output.seek(0)
            
            st.download_button(
                label="📥 ЗАВАНТАЖИТИ ГОТОВЕ КП",
                data=output,
                file_name=f"KP_Talo_{customer}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("Документ сформовано успішно!")
            
        except Exception as e:
            st.error(f"Виникла помилка: {e}")
