FORM_VERSION = "v_final_fixed_all"
import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os
import gspread
from google.oauth2.service_account import Credentials

# Спробуємо імпортувати num2words
try:
    from num2words import num2words
except ImportError:
    num2words = None

st.set_page_config(page_title="Talo КП Generator", layout="wide", page_icon="⚡")

# ================== ДОПОМІЖНІ ФУНКЦІЇ ==================
def amount_to_text_uk(amount):
    if num2words is None: return f"{amount} грн."
    units = int(amount)
    cents = int(round((amount - units) * 100))
    try:
        words = num2words(units, lang='uk').capitalize()
        return f"{words} гривень {cents:02d} копійок"
    except: return f"{amount} грн."

def save_to_google_sheets(row_data):
    try:
        if "gcp_service_account" not in st.secrets:
            return False
        credentials_info = st.secrets["gcp_service_account"]
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=scope)
        gc = gspread.authorize(creds)
        sh = gc.open("Реєстр КП Talo")
        worksheet = sh.get_worksheet(0)
        worksheet.append_row(row_data)
        return True
    except:
        return False

def replace_placeholders_stable(doc, replacements):
    """Заміна тексту через runs для збереження форматування"""
    def process_paragraph(p):
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                for run in p.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
    
    for p in doc.paragraphs: process_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: process_paragraph(p)

# ================== БАЗА РЕКВІЗИТІВ ==================
VENDORS_SPEC = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ «ТАЛО»",
        "short": "О. КРАМАРЕНКО", "inn": "45274534", 
        "adr": "03115, м. Київ, вул. Крамського Івана, 9", 
        "iban": "UA443052990000026004046815601", "bank": "АТ КБ «ПРИВАТБАНК»"
    },
    "ФОП Крамаренко Олексій Сергійович": {
        "full": "ФОП Крамаренко Олексій Сергійович",
        "short": "Олексій КРАМАРЕНКО", "inn": "3048920896", 
        "adr": "02156 м. Київ, вул. Кіото 9, кв. 40", 
        "iban": "UA423348510000000026009261015", "bank": "АТ «ПУМБ» м. Київ"
    }
}

# ================== ІНТЕРФЕЙС ==================
st.title("⚡ Генератор КП та Специфікацій")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_choice = col1.selectbox("Виконавець КП:", list(VENDORS_SPEC.keys()))
    
    vendor_info = VENDORS_SPEC[vendor_choice]
    tax_rate = 0.20 if vendor_choice == "ТОВ «ТАЛО»" else 0.06
    tax_label = "ПДВ (20%)" if vendor_choice == "ТОВ «ТАЛО»" else "Податкове навантаження (6%)"

    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса об'єкта")
    kp_num = col2.text_input("Номер КП/Договору", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_val = col2.date_input("Дата", datetime.date.today())
    date_str = date_val.strftime("%d.%m.%Y")
    phone = col2.text_input("Телефон", "+380 (67) 477-17-18")
    email = col2.text_input("E-mail", "o.kramarenko@talo.com.ua")

st.subheader("📝 Опис проекту (для КП)")
txt_intro = st.text_area("Вступний текст", "Відповідно до наданих даних пропонуємо наступне:")
c1, c2, c3 = st.columns(3)
l1 = c1.text_input("Пункт 1", "Організація автономного живлення ліфтів")
l2 = c2.text_input("Пункт 2", "Організація автономного живлення насосної")
l3 = c3.text_input("Пункт 3", "Аварійне освітлення та відеонагляд")

st.subheader("📦 Товари та послуги")
if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Обрати з: {cat}", list(EQUIPMENT_BASE[cat].keys()), key=f"sel_{cat}")
        current_keys = set(f"{cat}_{item}" for item in selected)
        for key in list(st.session_state.selected_items.keys()):
            if key.startswith(f"{cat}_") and key not in current_keys:
                del st.session_state.selected_items[key]

        for item in selected:
            key = f"{cat}_{item}"
            cA, cB, cC, cD = st.columns([3, 0.8, 1.2, 1])
            with cA: st.markdown(f"<div style='padding-top: 5px;'><b>{item}</b></div>", unsafe_allow_html=True)
            with cB: qty = st.number_input("К-сть", 1, 100, 1, key=f"q_{key}", label_visibility="collapsed")
            with cC: price = st.number_input("Ціна", 0, 1000000, int(EQUIPMENT_BASE[cat][item]), key=f"p_{key}", label_visibility="collapsed")
            subtotal = qty * price
            cD.markdown(f"<div style='padding-top: 5px;'><b>{subtotal:,}</b> грн</div>".replace(',', ' '), unsafe_allow_html=True)
            st.session_state.selected_items[key] = {"Найменування": item, "Кількість": qty, "Ціна": price, "Сума": subtotal, "Категорія": cat}

all_selected = list(st.session_state.selected_items.values())

if all_selected:
    st.divider()
    raw_total = sum(i["Сума"] for i in all_selected)
    tax_val = int(round(raw_total * tax_rate))
    final_total = raw_total + tax_val
    st.info(f"Загальна вартість: **{final_total:,}** грн".replace(',', ' '))

    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        save_to_google_sheets([date_str, kp_num, customer, address, final_total, manager])
        full_date_ukr = f"{date_val.day} { {1:'січня',2:'лютого',3:'березня',4:'квітня',5:'травня',6:'червня',7:'липня',8:'серпня',9:'вересня',10:'жовтня',11:'листопада',12:'грудня'}[date_val.month]} {date_val.year} року"
        safe_addr = re.sub(r'[\\/*?:"<>|]', "", address).replace(" ", "_")

        # --- 1. КП ---
        if os.path.exists("template.docx"):
            doc = Document("template.docx")
            replace_placeholders_stable(doc, {
                "vendor_name": vendor_choice, "customer": customer, "address": address, 
                "kp_num": kp_num, "date": date_str, "manager": manager, "phone": phone, "email": email,
                "txt_intro": txt_intro, "line1": l1, "line2": l2, "line3": l3
            })
            table = next((t for t in doc.tables if "Найменування" in t.rows[0].cells[0].text), doc.tables[0])
            for it in all_selected:
                row = table.add_row().cells
                row[0].text, row[1].text = it['Найменування'], str(it['Кількість'])
                row[2].text, row[3].text = f"{it['Ціна']:,}".replace(",", " "), f"{it['Сума']:,}".replace(",", " ")
            
            # Рядок разом у КП
            row_total = table.add_row().cells
            row_total[0].text = "ЗАГАЛЬНА ВАРТІСТЬ З УРАХУВАННЯМ ПОДАТКІВ, грн"
            row_total[0].merge(row_total[2])
            row_total[3].text = f"{final_total:,}".replace(",", " ")
            for run in row_total[0].paragraphs[0].runs + row_total[3].paragraphs[0].runs: run.bold = True

            buf_kp = BytesIO(); doc.save(buf_kp); buf_kp.seek(0)
            st.download_button(f"📥 КП_{kp_num}_{safe_addr}", buf_kp, f"КП_{kp_num}_{safe_addr}.docx")

        # Розподіл на залізо та послуги
        hw_items = [i for i in all_selected if "послуги" not in i["Категорія"].lower() and "роботи" not in i["Категорія"].lower()]
        work_items = [i for i in all_selected if i not in hw_items]

        # --- 2. СПЕЦИФІКАЦІЯ ПОСТАВКИ ---
        if hw_items and os.path.exists("template_postavka.docx"):
            doc_p = Document("template_postavka.docx")
            total_p = sum(i["Сума"] for i in hw_items)
            replace_placeholders_stable(doc_p, {
                "spec_id_postavka": f"№1 від {full_date_ukr}", "customer": customer, "address": address,
                "vendor_name": vendor_info["full"], "vendor_address": vendor_info["adr"], "vendor_inn": vendor_info["inn"],
                "vendor_iban": vendor_info["iban"], "vendor_email": email, "vendor_short_name": vendor_info["short"],
                "total_sum_digits": f"{total_p:,}".replace(",", " "), "total_sum_words": amount_to_text_uk(total_p)
            })
            table_p = doc_p.tables[0]
            for it in hw_items:
                row = table_p.add_row().cells
                row[0].text, row[1].text, row[2].text, row[3].text = it['Найменування'], str(it['Кількість']), f"{it['Ціна']:,}", f"{it['Сума']:,}"
            
            row_res = table_p.add_row().cells
            row_res[0].text, row_res[3].text = "РАЗОМ", f"{total_p:,}".replace(",", " ")
            buf_p = BytesIO(); doc_p.save(buf_p); buf_p.seek(0)
            st.download_button("📥 Специфікація Поставки", buf_p, f"Spec_Postavka_{kp_num}.docx")

        # --- 3. СПЕЦИФІКАЦІЯ РОБІТ ---
        if work_items and os.path.exists("template_roboti.docx"):
            doc_r = Document("template_roboti.docx")
            total_r = sum(i["Сума"] for i in work_items)
            replace_placeholders_stable(doc_r, {
                "spec_id_roboti": f"№1 від {full_date_ukr}", "customer": customer, "vendor_name": vendor_info["full"],
                "vendor_address": vendor_info["adr"], "vendor_inn": vendor_info["inn"], "vendor_iban": vendor_info["iban"],
                "vendor_email": email, "vendor_short_name": vendor_info["short"], "total_sum_words": amount_to_text_uk(total_r)
            })
            # Спеціальна заміна адреси для шаблону робіт
            for p in doc_r.paragraphs:
                if "{{ address }}" in p.text or "{{  address }}" in p.text:
                    p.text = p.text.replace("{{ address }}", address).replace("{{  address }}", address)
            
            table_r = doc_r.tables[0]
            for it in work_items:
                row = table_r.add_row().cells
                row[0].text, row[1].text, row[2].text, row[3].text = it['Найменування'], str(it['Кількість']), f"{it['Ціна']:,}", f"{it['Сума']:,}"
            
            row_res = table_r.add_row().cells
            row_res[0].text, row_res[3].text = "РАЗОМ", f"{total_r:,}".replace(",", " ")
            buf_r = BytesIO(); doc_r.save(buf_r); buf_r.seek(0)
            st.download_button("📥 Специфікація Робіт", buf_r, f"Spec_Roboti_{kp_num}.docx")
