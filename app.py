FORM_VERSION = "v_final_reset"
import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

st.set_page_config(page_title="Talo КП Generator", layout="wide", page_icon="⚡")

# --- ЛОГІКА АВТОМАТИЧНОГО НОМЕРА ---
if "kp_counter" not in st.session_state:
    st.session_state.kp_counter = 25 # Початкове число

def get_kp_number():
    # Формує номер за шаблоном 1223.[число]POW-B
    return f"1223.{st.session_state.kp_counter}POW-B"

# ================== ФУНКЦІЯ ЗАМІНИ ==================
def replace_placeholders(doc, replacements):
    bold_headers = ["Виконавець", "Замовник", "Адреса", "Відповідальний", "Контактний телефон", "E-mail", "Дата", "Комерційна пропозиція"]
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
                        p.add_run(left + ":").bold = True
                        p.add_run(right).bold = False
                        is_header = True
                        break
                if not is_header:
                    p.add_run(new_text).bold = False

    for p in doc.paragraphs: process_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: process_paragraph(p)

# ================== ІНТЕРФЕЙС ==================
st.title("⚡ Генератор Комерційних Пропозицій")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_choice = col1.selectbox("Виконавець:", ["ТОВ «ТАЛО»", "ФОП Крамаренко Олексій Сергійович"])
    
    if vendor_choice == "ТОВ «ТАЛО»":
        v_display, v_full, tax_rate, tax_label = "ТОВ «Тало»", "Директор ТОВ «ТАЛО»", 0.20, "ПДВ (20%)"
        curr_phone, curr_email, v_id = "+380 (67) 477-17-18", "o.kramarenko@talo.com.ua", "talo"
    else:
        v_display, v_full, tax_rate, tax_label = "ФОП Крамаренко О.С.", "ФОП Крамаренко О.С.", 0.06, "Податкове навантаження (6%)"
        curr_phone, curr_email, v_id = "+380 (67) 477-17-18", "o.kramarenko@talo.com.ua", "fop"

    customer = col1.text_input("Замовник", "ОСББ Вишгородська 45")
    address = col1.text_input("Адреса об'єкта", "м. Київ, вул. Вишгородська 45")
    
    # Номер КП тепер береться з функції
    current_kp = get_kp_number()
    kp_num = col2.text_input("Номер КП (автоматично)", value=current_kp)
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_str = col2.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
    
    phone = col2.text_input("Телефон", value=curr_phone, key=f"ph_{v_id}")
    email = col2.text_input("E-mail", value=curr_email, key=f"em_{v_id}")

# ... (Блок ТЗ та Специфікація залишаються без змін, як у вашому коді) ...
st.subheader("📝 Технічне завдання та опис")
txt_intro = st.text_area("Вступний текст ({{txt_intro}})", "Відповідно до наданих даних пропонуємо наступне:")
c1, c2, c3 = st.columns(3)
l1 = c1.text_input("Пункт 1 ({{line1}})", "Організація автономного живлення ліфтів")
l2 = c2.text_input("Пункт 2 ({{line2}})", "Організація автономного живлення насосної")
l3 = c3.text_input("Пункт 3 ({{line3}})", "Аварійне освітлення та відеонагляд")

st.divider()
st.subheader("📦 Специфікація")
if "selected_items" not in st.session_state: st.session_state.selected_items = {}

tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Обрати з: {cat}", list(EQUIPMENT_BASE[cat].keys()), key=f"sel_{cat}")
        current_keys = set(f"{cat}_{item}" for item in selected)
        for key in list(st.session_state.selected_items.keys()):
            if key.startswith(f"{cat}_") and key not in current_keys: del st.session_state.selected_items[key]
        if selected:
            h1, h2, h3, h4 = st.columns([3, 0.8, 1.2, 1])
            h1.caption("🏷️ Товар"); h2.caption("🔢 К-сть"); h3.caption("💰 Ціна"); h4.caption("📈 Сума")
            for item in selected:
                with st.container():
                    cA, cB, cC, cD = st.columns([3, 0.8, 1.2, 1])
                    with cA: st.markdown(f"<div style='padding-top:5px;'><b>{item}</b></div>", unsafe_allow_html=True)
                    with cB: qty = st.number_input("К-сть", min_value=1, value=1, key=f"qty_{cat}_{item}", label_visibility="collapsed")
                    with cC: price = st.number_input("Ціна", min_value=0, value=int(EQUIPMENT_BASE[cat][item]), key=f"pr_{cat}_{item}", label_visibility="collapsed")
                    subtotal = int(qty * price)
                    with cD: st.markdown(f"<div style='padding-top:5px;'><b>{subtotal:,}</b> грн</div>".replace(',', ' '), unsafe_allow_html=True)
                    st.session_state.selected_items[f"{cat}_{item}"] = {"Найменування": item, "Кількість": qty, "Ціна": price, "Сума": subtotal, "Категорія": cat}

# ================== ГЕНЕРАЦІЯ ==================
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
            "vendor_name": v_display, "vendor_full_name": v_full, "customer": customer, "address": address, 
            "kp_num": kp_num, "manager": manager, "date": date_str, "phone": phone, "email": email,
            "txt_intro": txt_intro, "line1": l1, "line2": l2, "line3": l3
        })
        
        # (Логіка заповнення таблиці - залишаємо як була)
        target_table = next((t for t in doc.tables if "Найменування" in t.rows[0].cells[0].text), None)
        if target_table:
            sections = {"ОБЛАДНАННЯ": ["1. Інвертори Deye", "2. Акумулятори (АКБ)"], "МАТЕРІАЛИ": ["3. Комплектуючі та щити"], "РОБОТИ ТА ПОСЛУГИ": ["4. Послуги та Роботи"]}
            for sec, cats in sections.items():
                items = [x for x in all_selected_data if x["Категорія"] in cats]
                if items:
                    row = target_table.add_row()
                    row.cells[0].merge(row.cells[3]).paragraphs[0].add_run(sec).italic = True
                    for it in items:
                        r = target_table.add_row().cells
                        r[0].text, r[1].text = f" - {it['Найменування']}", str(it["Кількість"])
                        r[2].text = f"{it['Ціна']:,}".replace(",", " ")
                        r[3].text = f"{it['Сума']:,}".replace(",", " ")

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        
        # Формуємо назву файлу
        safe_addr = re.sub(r'[\\/*?:"<>|]', "", address).replace(" ", "_")
        file_name = f"КП_{kp_num}_{safe_addr}.docx"

        st.download_button(
            label="✅ Завантажити файл",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            on_click=lambda: st.session_state.update({"kp_counter": st.session_state.kp_counter + 1})
        )
