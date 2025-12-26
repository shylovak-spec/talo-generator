import streamlit as st
import pandas as pd
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Talo КП Generator", page_icon="⚡", layout="wide")

# --- ФУНКЦІЯ ЗАМІНИ МІТОК ---
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

    for p in doc.paragraphs: process_element(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: process_element(p)

# --- ІНТЕРФЕЙС ---
st.title("⚡ Генератор КП ТОВ «Тало»")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        customer = st.text_input("Замовник", "ОСББ Вишгородська 45")
        address = st.text_input("Адреса", "м. Київ, вул. Вишгородська 45")
        kp_num = st.text_input("Номер КП", "1223.25POW-B")
        tax_type = st.radio("Податок:", ["ПДВ (20%)", "Податкове навантаження (6%)", "Без податку"], horizontal=True)
    with col2:
        manager = st.text_input("Відповідальний", "Олексій Крамаренко")
        date_str = st.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
        phone = st.text_input("Телефон", "+380 (67) 477-17-18")

st.subheader("📦 Специфікація (Оберіть та вкажіть ціну)")
all_selected_data = []
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))

for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Додати з розділу {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"sel_{i}")
        for item in selected:
            c1, c2, c3, c4 = st.columns([3, 1, 2, 2])
            with c1: st.write(f"**{item}**")
            with c2: qty = st.number_input("К-сть", min_value=1, value=1, key=f"q_{item}")
            with c3: price = st.number_input("Ціна за од, грн", min_value=0, value=int(EQUIPMENT_BASE[cat][item]), key=f"p_{item}")
            with c4:
                subtotal = qty * price
                st.write(f"**{subtotal:,}** грн".replace(',', ' '))
                all_selected_data.append({"Найменування": item, "Кількість": qty, "Ціна": price, "Сума": subtotal, "Категорія": cat})

if all_selected_data:
    st.divider()
    raw_total = sum(item["Сума"] for item in all_selected_data)
    t_rate = 0.20 if tax_type == "ПДВ (20%)" else (0.06 if tax_type == "Податкове навантаження (6%)" else 0)
    tax_val = raw_total * t_rate
    final_total = raw_total + tax_val

    st.write(f"Сума: {raw_total:,} грн | Податок: {tax_val:,} грн")
    st.header(f"Усього: {final_total:,} грн")

    if st.button("🚀 Згенерувати КП"):
        doc = Document("template.docx")
        replace_placeholders(doc, {"customer": customer, "address": address, "kp_num": kp_num, "manager": manager, "date": date_str, "phone": phone})

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
                    row = target_table.add_row().cells
                    row[0].text = sec_name
                    row[0].paragraphs[0].runs[0].bold = True
                    for it in items:
                        cells = target_table.add_row().cells
                        cells[0].text = f" - {it['Найменування']}"
                        cells[1].text = str(it['Кількість'])
                        cells[2].text = f"{it['Ціна']:,}".replace(',', ' ')
                        cells[3].text = f"{it['Сума']:,}".replace(',', ' ')

            # Підсумки
            target_table.add_row()
            r1 = target_table.add_row().cells
            r1[0].text, r1[3].text = "РАЗОМ:", f"{raw_total:,}".replace(',', ' ')
            r1[0].paragraphs[0].runs[0].bold = True
            
            r2 = target_table.add_row().cells
            r2[0].text, r2[3].text = tax_type + ":", f"{tax_val:,}".replace(',', ' ')
            
            r3 = target_table.add_row().cells
            r3[0].text, r3[3].text = "ЗАГАЛЬНА ВАРТІСТЬ:", f"{final_total:,}".replace(',', ' ')
            for cell in r3: 
                if cell.text: cell.paragraphs[0].runs[0].bold = True

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 ЗАВАНТАЖИТИ", output, f"KP_{customer}.docx")
