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
    bold_headers = ["Виконавець", "Замовник", "Адреса", "Відповідальний", "Контактний телефон", "E-mail", "Дата", "Комерційна пропозиція"]
    for p in doc.paragraphs:
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                new_text = p.text.replace(placeholder, str(value))
                p.clear()
                is_header = False
                for bh in bold_headers:
                    if new_text.strip().startswith(bh + ":"):
                        left, right = new_text.split(":", 1)
                        p.add_run(left + ":").bold = True
                        p.add_run(right).bold = False
                        is_header = True
                        break
                if not is_header:
                    p.add_run(new_text).bold = False

# ================== ІНТЕРФЕЙС ==================
st.title("⚡ Генератор Комерційних Пропозицій")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    
    vendor_choice = col1.selectbox("Виконавець:", ["ТОВ «ТАЛО»", "ФОП Крамаренко Олексій Сергійович"])

    # ВАЖЛИВО: Правильні відступи тут
    if vendor_choice == "ТОВ «ТАЛО»":
        v_display, v_full = "ТОВ «Тало»", "Директор ТОВ «ТАЛО»"
        tax_rate, tax_label = 0.20, "ПДВ (20%)"
        curr_phone, curr_email, v_id = "+380 (67) 477-17-18", "o.kramarenko@talo.com.ua", "talo"
    else:
        v_display, v_full = "ФОП Крамаренко О.С.", "ФОП Крамаренко О.С."
        tax_rate, tax_label = 0.06, "Податкове навантаження (6%)"
        curr_phone, curr_email, v_id = "+380 (67) 477-17-18", "o.kramarenko@talo.com.ua", "fop"

    customer = col1.text_input("Замовник", "ОСББ Вишгородська 45")
    address = col1.text_input("Адреса об'єкта", "м. Київ, вул. Вишгородська 45")
    
    kp_num = col2.text_input("Номер КП", "1223.25POW-B")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_str = col2.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
    
    # Динамічні ключі змушують Streamlit оновити значення
    phone = col2.text_input("Телефон", value=curr_phone, key=f"p_{v_id}")
    email = col2.text_input("E-mail", value=curr_email, key=f"e_{v_id}")

st.divider()

# ================== СПЕЦИФІКАЦІЯ ==================
if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        sel = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"s_{cat}")
        
        # Очищення старих
        current_keys = set(f"{cat}_{item}" for item in sel)
        for k in list(st.session_state.selected_items.keys()):
            if k.startswith(f"{cat}_") and k not in current_keys:
                del st.session_state.selected_items[k]

        for item in sel:
            cA, cB, cC, cD = st.columns([3, 1, 2, 2])
            qty = cB.number_input("К-сть", 1, key=f"q_{cat}_{item}")
            pr = cC.number_input("Ціна", 0, value=int(EQUIPMENT_BASE[cat][item]), key=f"p_{cat}_{item}")
            sub = qty * pr
            cD.write(f"**{sub:,}** грн")
            st.session_state.selected_items[f"{cat}_{item}"] = {
                "Найменування": item, "Кількість": qty, "Ціна": pr, "Сума": sub, "Категорія": cat
            }

# ================== ФІНАЛ ==================
all_data = list(st.session_state.selected_items.values())
if all_data:
    raw_total = sum(item["Сума"] for item in all_data) # Тут тепер ТІЛЬКИ кирилиця
    tax_val = int(raw_total * tax_rate)
    final_total = raw_total + tax_val
    st.info(f"Разом: {final_total:,} грн")

    if st.button("Генерувати КП"):
        doc = Document("template.docx")
        replace_placeholders(doc, {
            "vendor_name": v_display, "vendor_full_name": v_full, "customer": customer,
            "address": address, "kp_num": kp_num, "manager": manager, "date": date_str,
            "phone": phone, "email": email
        })
        output = BytesIO()
        doc.save(output)
        st.download_button("Завантажити КП", output.getvalue(), f"KP_{kp_num}.docx")
