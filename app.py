import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

st.set_page_config(page_title="Talo КП Generator", layout="wide", page_icon="⚡")

# ================== ФУНКЦІЯ ЗАМІНИ ==================
def replace_placeholders(doc, replacements):
    bold_headers = [
        "Виконавець", "Замовник", "Адреса", "Відповідальний",
        "Контактний телефон", "E-mail", "Дата", "Комерційна пропозиція"
    ]
    def process_paragraph(p):
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                new_text = p.text.replace(placeholder, str(value))
                p.clear()
                is_header = False
                for bh in bold_headers:
                    if new_text.strip().startswith(bh + ":"):
                        left, right = new_text.split(":", 1)
                        r1 = p.add_run(left + ":")
                        r1.bold = True
                        r2 = p.add_run(right)
                        r2.bold = False
                        is_header = True
                        break
                if not is_header:
                    p.add_run(new_text).bold = False

    for p in doc.paragraphs: process_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: process_paragraph(p)

# ================== ІНТЕРФЕЙС STREAMLIT ==================
st.title("⚡ Генератор Комерційних Пропозицій")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    
    vendor_choice = col1.selectbox(
    "Виконавець:", 
    ["ТОВ «ТАЛО»", "ФОП Крамаренко Олексій Сергійович"]
)

    if vendor_choice == "ТОВ «ТАЛО»":
    v_display, v_full = "ТОВ «Тало»", "Директор ТОВ «ТАЛО»"
    tax_rate, tax_label = 0.20, "ПДВ (20%)"
    curr_phone = "+380 (67) 477-17-18"
    curr_email = "o.kramarenko@talo.com.ua"
    v_id = "talo_choice" # Унікальний маркер для ТОВ
else:
    v_display, v_full = "ФОП Крамаренко О.С.", "ФОП Крамаренко О.С."
    tax_rate, tax_label = 0.06, "Податкове навантаження (6%)"
    curr_phone = "+380 (67) 477-17-18" # Ваше тимчасово однакове значення
    curr_email = "o.kramarenko@talo.com.ua"
    v_id = "fop_choice"  # Унікальний маркер для ФОП

# ... (інші поля) ...

with col2:
    # 3. ВИПРАВЛЕННЯ: додаємо v_id у key. 
    # Коли ви перемкнете ТОВ на ФОП, ключ зміниться з "..._talo_choice" на "..._fop_choice".
    # Streamlit сприйме це як НОВЕ поле і завантажить актуальний value.
    phone = st.text_input("Телефон", value=curr_phone, key=f"phone_{v_id}")
    email = st.text_input("E-mail", value=curr_email, key=f"email_{v_id}")

st.subheader("📝 Технічне завдання та опис")
txt_intro = st.text_area("Вступний текст", "Відповідно до наданих даних пропонуємо наступне:")
c1, c2, c3 = st.columns(3)
l1 = c1.text_input("Пункт 1", "Організація автономного живлення ліфтів")
l2 = c2.text_input("Пункт 2", "Організація автономного живлення насосної")
l3 = c3.text_input("Пункт 3", "Аварійне освітлення та відеонагляд")

st.divider()

# ================== СПЕЦИФІКАЦІЯ ==================
st.subheader("📦 Специфікація")

if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"sel_{cat}")
        
        current_keys = set(f"{cat}_{item}" for item in selected)
        for key in list(st.session_state.selected_items.keys()):
            if key.startswith(f"{cat}_") and key not in current_keys:
                del st.session_state.selected_items[key]

        for item in selected:
            cA, cB, cC, cD = st.columns([3, 1, 2, 2])
            with cA: st.write(f"**{item}**")
            with cB: qty = st.number_input("К-сть", min_value=1, value=1, key=f"qty_{cat}_{item}")
            with cC: price = st.number_input("Ціна, грн", min_value=0, value=int(EQUIPMENT_BASE[cat][item]), key=f"pr_{cat}_{item}")
            sub = int(qty * price)
            with cD: st.write(f"**{sub:,}** грн".replace(',', ' '))
            
            st.session_state.selected_items[f"{cat}_{item}"] = {
                "Найменування": item, "Кількість": qty, "Ціна": price, "Сума": sub, "Категорія": cat
            }

# ================== ОБРОБКА ТА ВАРТІСТЬ ==================
all_selected_data = list(st.session_state.selected_items.values())

if all_selected_data:
    st.divider()
    raw_total = sum(item["Сума"] for item in all_selected_data)
    tax_val = int(round(raw_total * tax_rate))
    final_total = raw_total + tax_val
    
    st.info(f"Загальна вартість КП: **{final_total:,}** грн".replace(',', ' '))

    if st.button("🚀 Згенерувати КП", type="primary", use_container_width=True):
        doc = Document("template.docx")
        
        replace_placeholders(doc, {
            "vendor_name": v_display, "vendor_full_name": v_full,
            "customer": customer, "address": address, "kp_num": kp_num, 
            "manager": manager, "date": date_str, "phone": phone, "email": email,
            "txt_intro": txt_intro, "line1": l1, "line2": l2, "line3": l3
        })

        target_table = next((t for t in doc.tables if "Найменування" in t.rows[0].cells[0].text), None)
        if target_table:
            sections = {
                "ОБЛАДНАННЯ": ["1. Інвертори Deye", "2. Акумулятори (АКБ)"],
                "МАТЕРІАЛИ": ["3. Комплектуючі та щити"],
                "РОБОТИ ТА ПОСЛУГИ": ["4. Послуги та Роботи"]
            }
            
            for sec, cats in sections.items():
                items = [x for x in all_selected_data if x["Категорія"] in cats]
                if items:
                    row = target_table.add_row()
                    merged = row.cells[0].merge(row.cells[3])
                    p = merged.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(sec)
                    run.bold, run.italic = False, True
                    
                    for it in items:
                        r = target_table.add_row().cells
                        r[0].text = f" - {it['Найменування']}"
                        r[1].text = str(it["Кількість"])
                        r[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r[2].text = f"{it['Ціна']:,}".replace(",", " ")
                        r[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        r[3].text = f"{it['Сума']:,}".replace(",", " ")
                        r[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

            summary = [
                ("РАЗОМ, грн:", raw_total, False), 
                (f"{tax_label}:", tax_val, False), 
                ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", final_total, True)
            ]
            for label, val, is_bold in summary:
                r = target_table.add_row().cells
                r[0].text, r[3].text = label, f"{val:,}".replace(",", " ")
                r[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if is_bold:
                    for c in r:
                        for rn in c.paragraphs[0].runs: rn.bold = True

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button(
            label="✅ Завантажити файл",
            data=output,
            file_name=f"KP_{kp_num}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
