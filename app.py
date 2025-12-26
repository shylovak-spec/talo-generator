import streamlit as st
import pandas as pd
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO 
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Talo КП Generator", page_icon="⚡", layout="wide")

# --- НАДІЙНА ФУНКЦІЯ ЗАМІНИ ---
def replace_placeholders(doc, replacements):
    # 1. Обробка абзаців (шапка та вступ)
    for p in doc.paragraphs:
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                # Склеюємо текст, щоб знайти мітку, яка розбита Word-ом
                full_text = "".join([run.text for run in p.runs])
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(value))
                    
                    # Якщо це рядок з двократкою (наприклад, "Замовник: ...")
                    if ":" in new_text and key not in ["txt_intro", "line1", "line2", "line3"]:
                        header, data = new_text.split(":", 1)
                        # Очищаємо існуючі runs, не видаляючи сам абзац
                        for run in p.runs:
                            run.text = ""
                        # Додаємо жирний заголовок і звичайні дані
                        r1 = p.add_run(header + ":")
                        r1.bold = True
                        r2 = p.add_run(data)
                        r2.bold = False
                    else:
                        # Для звичайного тексту (вступ, пункти) просто замінюємо текст
                        # зберігаючи існуючий стиль першого прогону
                        first_run = True
                        for run in p.runs:
                            if first_run:
                                run.text = new_text
                                run.bold = False
                                first_run = False
                            else:
                                run.text = ""

    # 2. Обробка таблиць (якщо мітки там є)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in replacements.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            full_text = "".join([run.text for run in p.runs])
                            if placeholder in full_text:
                                final_val = full_text.replace(placeholder, str(value))
                                for run in p.runs: run.text = ""
                                r = p.add_run(final_val)
                                r.bold = False

# --- ІНТЕРФЕЙС ---
st.title("⚡ Генератор КП")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        vendor_choice = st.selectbox("Виконавець:", ["ТОВ «ТАЛО»", "ФОП Крамаренко Олексій Сергійович"])
        customer = st.text_input("Замовник", "ОСББ Вишгородська 45")
        address = st.text_input("Адреса об'єкта", "м. Київ, вул. Вишгородська 45")
    with col2:
        kp_num = st.text_input("Номер КП", "1223.25POW-B")
        manager = st.text_input("Відповідальний", "Олексій Крамаренко")
        date_str = st.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
        phone = st.text_input("Телефон", "+380 (67) 477-17-18")
        email = "o.kramarenko@talo.com.ua"

# Логіка податків
if vendor_choice == "ТОВ «ТАЛО»":
    v_display, v_full = "ТОВ «Тало»", "Директор ТОВ «ТАЛО»"
    tax_rate, tax_label = 0.20, "ПДВ (20%)"
else:
    v_display, v_full = "ФОП Крамаренко О.С.", "ФОП Крамаренко Олексій Сергійович"
    tax_rate, tax_label = 0.06, "Податкове навантаження (6%)"

st.subheader("📝 Технічні умови")
txt_intro = st.text_area("Вступний опис", "Відповідно до наданих даних пропонуємо наступне:")
l1 = st.text_input("Пункт 1", "Організація автономного живлення ліфтів")
l2 = st.text_input("Пункт 2", "Організація автономного живлення насосної")
l3 = st.text_input("Пункт 3", "Аварійне освітлення та відеонагляд")

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
                subtotal = int(qty * price)
                st.write(f"**{subtotal:,}** грн")
                all_selected_data.append({"Найменування": item, "Кількість": qty, "Ціна": price, "Сума": subtotal, "Категорія": cat})

if all_selected_data:
    raw_total = int(sum(item["Сума"] for item in all_selected_data))
    tax_val = int(round(raw_total * tax_rate, 0))
    final_total = int(raw_total + tax_val)

    if st.button("🚀 Згенерувати КП"):
        doc = Document("template.docx")
        info = {
            "vendor_name": v_display, "vendor_full_name": v_full,
            "customer": customer, "address": address, "kp_num": kp_num, 
            "manager": manager, "date": date_str, "phone": phone, "email": email,
            "txt_intro": txt_intro, "line1": l1, "line2": l2, "line3": l3
        }
        replace_placeholders(doc, info)

        target_table = next((t for t in doc.tables if "Найменування" in t.rows[0].cells[0].text), None)
        if target_table:
            sections = {"ОБЛАДНАННЯ": ["1. Інвертори Deye", "2. Акумулятори (АКБ)"], "МАТЕРІАЛИ": ["3. Комплектуючі та щити"], "РОБОТИ ТА ПОСЛУГИ": ["4. Послуги та Роботи"]}
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
                        cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        cells[2].text = f"{it['Ціна']:,}".replace(',', ' ')
                        cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        cells[3].text = f"{it['Сума']:,}".replace(',', ' ')
                        cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

            target_table.add_row()
            for label, val, is_bold in [("РАЗОМ (без податку):", raw_total, False), (f"{tax_label}:", tax_val, False), ("УСЬОГО ДО СПЛАТИ:", final_total, True)]:
                r = target_table.add_row().cells
                r[0].text, r[3].text = label, f"{val:,}".replace(',', ' ')
                r[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if is_bold:
                    for c in r: 
                        if c.text: c.paragraphs[0].runs[0].bold = True

        # Назва файлу без спецсимволів
        clean_name = "".join([c for c in customer if c.isalnum() or c in ' _-']).strip()
        file_name = f"KP_{kp_num}_{clean_name}.docx"
        
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 ЗАВАНТАЖИТИ КП", output, file_name)
