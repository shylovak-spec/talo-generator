FORM_VERSION = "v_final_fix_total"
import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Talo КП Generator", layout="wide", page_icon="⚡")

# ================== ФУНКЦІЯ GOOGLE SHEETS ==================
def save_to_google_sheets(row_data):
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Секрети не знайдено в Streamlit Secrets!")
            return False
            
        credentials_info = st.secrets["gcp_service_account"]
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=scope)
        gc = gspread.authorize(creds)
        
        sh = gc.open("Реєстр КП Talo")
        worksheet = sh.get_worksheet(0)
        worksheet.append_row(row_data)
        return True
    except Exception as e:
        st.error(f"❌ Помилка Google Sheets: {e}")
        return False

# ================== ФУНКЦІЯ ЗАМІНИ (Шапка та Текст) ==================
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
                        parts = new_text.split(":", 1)
                        p.add_run(parts[0] + ":").bold = True
                        if len(parts) > 1: p.add_run(parts[1]).bold = False
                        is_header = True
                        break
                if not is_header: p.add_run(new_text).bold = False

    for p in doc.paragraphs: process_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: process_paragraph(p)

# ================== ІНТЕРФЕЙС STREAMLIT ==================
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
    kp_num = col2.text_input("Номер КП", "1223.25POW-B")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_str = col2.date_input("Дата", datetime.date.today()).strftime("%d.%m.%Y")
    phone = col2.text_input("Телефон", value=curr_phone, key=f"ph_{v_id}")
    email = col2.text_input("E-mail", value=curr_email, key=f"em_{v_id}")

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
                cA, cB, cC, cD = st.columns([3, 0.8, 1.2, 1])
                cA.markdown(f"**{item}**")
                qty = cB.number_input("К-сть", min_value=1, value=1, key=f"qty_{cat}_{item}", label_visibility="collapsed")
                price = cC.number_input("Ціна", min_value=0, value=int(EQUIPMENT_BASE[cat][item]), key=f"pr_{cat}_{item}", label_visibility="collapsed")
                subtotal = int(qty * price)
                cD.markdown(f"**{subtotal:,}** грн".replace(',', ' '))
                st.session_state.selected_items[f"{cat}_{item}"] = {"Найменування": item, "Кількість": qty, "Ціна": price, "Сума": subtotal, "Категорія": cat}

all_selected_data = list(st.session_state.selected_items.values())

if all_selected_data:
    st.divider()
    raw_total = sum(item["Сума"] for item in all_selected_data)
    tax_val = int(round(raw_total * tax_rate))
    final_total = raw_total + tax_val
    st.info(f"Загальна вартість КП: **{final_total:,}** грн".replace(',', ' '))

    if st.button("🚀 Згенерувати та завантажити КП", type="primary", use_container_width=True):
        doc = Document("template.docx")
        replace_placeholders(doc, {
            "vendor_name": v_display, "vendor_full_name": v_full, "customer": customer, "address": address, 
            "kp_num": kp_num, "manager": manager, "date": date_str, "phone": phone, "email": email,
            "txt_intro": "Відповідно до наданих даних пропонуємо наступне:", "line1": "Пункт 1", "line2": "Пункт 2", "line3": "Пункт 3"
        })

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
                        r[0].text = f" - {it['Найменування']}"
                        r[1].text = str(it["Кількість"])
                        r[2].text = f"{it['Ціна']:,}".replace(",", " ")
                        r[3].text = f"{it['Сума']:,}".replace(",", " ")

            # --- ОСЬ ТУТ ВИПРАВЛЕНО РОЗРАХУНОК РАЗОМ ---
            summary_data = [
                ("РАЗОМ, грн:", raw_total, False),
                (f"{tax_label}:", tax_val, False),
                ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", final_total, True)
            ]
            for label, val, is_bold in summary_data:
                r = target_table.add_row().cells
                r[0].text = label
                r[3].text = f"{val:,}".replace(",", " ")
                r[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if is_bold:
                    for cell in r:
                        for run in cell.paragraphs[0].runs: run.bold = True

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        
        safe_addr = re.sub(r'[\\/*?:"<>|«»]', "", address).replace(" ", "_")
        if save_to_google_sheets([date_str, kp_num, customer, address, final_total, manager]):
            st.toast("📊 Дані збережено в Google Sheets!")

        st.download_button("✅ Завантажити готовий файл", output, f"КП_{kp_num}_{safe_addr}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
