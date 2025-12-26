import streamlit as st
import pandas as pd
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Talo КП Generator", page_icon="⚡", layout="wide")

# --- РОЗУМНА ФУНКЦІЯ ЗАМІНИ (Склеює розірвані мітки) ---
def replace_placeholders(doc, replacements):
    for p in doc.paragraphs:
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                full_text = "".join([run.text for run in p.runs])
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(value))
                    for i, run in enumerate(p.runs):
                        if i == 0:
                            run.text = new_text
                            run.bold = False
                        else:
                            run.text = ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in replacements.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            full_text = "".join([run.text for run in p.runs])
                            if placeholder in full_text:
                                new_text = full_text.replace(placeholder, str(value))
                                for i, run in enumerate(p.runs):
                                    if i == 0:
                                        run.text = new_text
                                        run.bold = False
                                    else:
                                        run.text = ""

# --- ІНТЕРФЕЙС ---
st.title("⚡ Генератор КП")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        vendor_choice = st.selectbox(
            "Оберіть Виконавця:",
            ["ТОВ «ТАЛО»", "ФОП Крамаренко Олексій Сергійович"]
        )
        customer = st.text_input("Замовник", "ОСББ Вишгородська 45")
        address = st.text_input("Адреса об'єкта", "м. Київ, вул. Вишгородська 45")
    
    with col2:
        kp_num = st.text_input("Номер КП", "1223.25POW-B")
        manager = st.text_input("Відповідальний", "Олексій Крамаренко")
        date_str = st.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
        phone = st.text_input("Телефон", "+380 (67) 477-17-18")

# Логіка виконавця та оподаткування
if vendor_choice == "ТОВ «ТАЛО»":
    v_display = "ТОВ «Тало»"
    v_full_name = "Директор ТОВ «ТАЛО»"
    tax_rate = 0.20
    tax_label = "ПДВ (20%)"
else:
    v_display = "ФОП Крамаренко О.С."
    v_full_name = "ФОП Крамаренко Олексій Сергійович"
    tax_rate = 0.06
    tax_label = "Податкове навантаження (6%)"

st.subheader("📝 Технічні умови")
txt_intro = st.text_area("Вступ", "Відповідно до наданих даних пропонуємо наступне...")
l1 = st.text_input("Пункт 1", "Організація автономного живлення ліфтів...")
l2 = st.text_input("Пункт 2", "Організація автономного живлення насосної...")
l3 = st.text_input("Пункт 3", "Аварійне освітлення та відеонагляд;")

st.subheader("📦 Специфікація")
all_selected_data = []
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))

for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"sel_{i}")
        for item in selected:
            c1, c2, c3, c4 = st.columns([3, 1, 2, 2])
            with c1: st.write(f"**{item}**")
            with c2: qty = st.number_input("К-сть", min_value=1, value=1, key=f"q_{item}")
            with c3: price = st.number_input("Ціна, грн", min_value=0, value=int(EQUIPMENT_BASE[cat][item]), key=f"p_{item}")
            with c4:
                # Розрахунок рядка
                subtotal = int(qty * price)
                st.write(f"**{subtotal:,}** грн")
                all_selected_data.append({
                    "Найменування": item, 
                    "Кількість": qty, 
                    "Ціна": price, 
                    "Сума": subtotal, 
                    "Категорія": cat
                })

if all_selected_data:
    st.divider()
    
    # --- МАТЕМАТИЧНЕ ЗАОКРУГЛЕННЯ ---
    raw_total = sum(item["Сума"] for item in all_selected_data)
    tax_val = round(raw_total * tax_rate, 0) # заокруглюємо до цілого для чистоти
    final_total = int(raw_total + tax_val)

    st.write(f"Сума: **{raw_total:,}** грн")
    st.write(f"{tax_label}: **{tax_val:,}** грн")
    st.header(f"Усього: {final_total:,} грн")

    if st.button("🚀 Згенерувати КП"):
        doc = Document("template.docx")
        
        info = {
            "vendor_name": v_display, 
            "vendor_full_name": v_full_name,
            "customer": customer, 
            "address": address, 
            "kp_num": kp_num, 
            "manager": manager, 
            "date": date_str, 
            "phone": phone,
            "txt_intro": txt_intro, 
            "line1": l1, 
            "line2": l2, 
            "line3": l3
        }
        replace_placeholders(doc, info)

        target_table = next((t for t in doc.tables if "Найменування" in t.rows[0].cells[0].text), None)
        if target_table:
            sections = {
                "ОБЛАДНАННЯ": ["1. Інвертори Deye", "2. Акумулятори (АКБ)"],
                "МАТЕРІАЛИ": ["3. Комплектуючі та щити"],
                "РОБОТИ ТА ПОСЛУГИ": ["4. Послуги та Роботи"]
            }
            for sec_name, cats in sections.items():
                items = [x for x in all_selected_data if x["Категорія"] in cats]
                if items:
                    row_h = target_table.add_row().cells
                    row_h[0].text = sec_name
                    row_h[0].paragraphs[0].runs[0].bold = True
                    for it in items:
                        cells = target_table.add_row().cells
                        cells[0].text = f" - {it['Найменування']}"
                        cells[1].text = str(it['Кількість'])
                        cells[2].text = f"{it['Ціна']:,}".replace(',', ' ')
                        cells[3].text = f"{it['Сума']:,}".replace(',', ' ')

            # Фінальні рядки в таблиці
            target_table.add_row()
            r1 = target_table.add_row().cells
            r1[0].text, r1[3].text = "РАЗОМ (без податку):", f"{raw_total:,}".replace(',', ' ')
            
            r2 = target_table.add_row().cells
            r2[0].text, r2[3].text = f"{tax_label}:", f"{int(tax_val):,}".replace(',', ' ')
            
            r3 = target_table.add_row().cells
            r3[0].text, r3[3].text = "ЗАГАЛЬНА ВАРТІСТЬ:", f"{final_total:,}".replace(',', ' ')
            for cell in r3:
                if cell.text: cell.paragraphs[0].runs[0].bold = True

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 ЗАВАНТАЖИТИ КП", output, f"KP_{customer}.docx")
