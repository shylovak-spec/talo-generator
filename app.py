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

try:
    from num2words import num2words
except ImportError:
    num2words = None

# ================== НАЛАШТУВАННЯ ТА ДАНІ ==================
VENDORS = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ «ТАЛО»",
        "short": "О. КРАМАРЕНКО",
        "inn": "45274534",
        "adr": "03115, м. Київ, вул. Крамського Івана, 9",
        "iban": "UA443052990000026004046815601",
        "tax_label": "ПДВ (20%)",
        "tax_rate": 0.20
    },
    "ФОП Крамаренко Олексій Сергійович": {
        "full": "ФОП Крамаренко Олексій Сергійович",
        "short": "Олексій КРАМАРЕНКО",
        "inn": "3048920896",
        "adr": "02156 м. Київ, вул. Кіото 9, кв. 40",
        "iban": "UA423348510000000026009261015",
        "tax_label": "Податкове навантаження (6%)",
        "tax_rate": 0.06
    }
}

# ================== ДОПОМІЖНІ ФУНКЦІЇ ==================
def amount_to_text_uk(amount):
    if num2words is None: return f"{amount} грн."
    units, cents = divmod(int(round(amount * 100)), 100)
    try:
        words = num2words(units, lang='uk').capitalize()
        return f"{words} гривень {cents:02d} копійок"
    except: return f"{amount} грн."

def save_to_google_sheets(row_data):
    try:
        if "gcp_service_account" not in st.secrets: return False
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], 
                                                     scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        gc = gspread.authorize(creds)
        sh = gc.open("Реєстр КП Talo")
        sh.get_worksheet(0).append_row(row_data)
        return True
    except: return False

def replace_placeholders_stable(doc, replacements):
    """Агресивна заміна тегів у всьому документі"""
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if f"{{{{{key}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{key}}}}}", str(val))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in replacements.items():
                        if f"{{{{{key}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{key}}}}}", str(val))

# ================== ІНТЕРФЕЙС ==================
st.set_page_config(page_title="Talo Generator", layout="wide")
st.title("⚡ Генератор КП та Специфікацій")

if "generated_files" not in st.session_state:
    st.session_state.generated_files = None

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_choice = col1.selectbox("Виконавець:", list(VENDORS.keys()))
    v = VENDORS[vendor_choice]
    
    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса об'єкта")
    
    kp_num = col2.text_input("Номер КП/Договору", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_val = col2.date_input("Дата", datetime.date.today())
    date_str = date_val.strftime("%d.%m.%Y")
    phone = col2.text_input("Телефон", "+380 (67) 477-17-18")
    email = col2.text_input("E-mail", "o.kramarenko@talo.com.ua")

st.subheader("📝 Текст для КП")
txt_intro = st.text_area("Вступний текст", "Відповідно до наданих даних пропонуємо наступне:")
c1, c2, c3 = st.columns(3)
l1 = c1.text_input("Пункт 1", "Організація автономного живлення ліфтів")
l2 = c2.text_input("Пункт 2", "Організація автономного живлення насосної")
l3 = c3.text_input("Пункт 3", "Аварійне освітлення та відеонагляд")

st.subheader("📦 Специфікація")
if "selected_items" not in st.session_state: st.session_state.selected_items = {}

tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"s_{cat}")
        for item in selected:
            key = f"{cat}_{item}"
            cA, cB, cC, cD = st.columns([3, 0.8, 1.2, 1])
            qty = cB.number_input("К-сть", 1, 100, 1, key=f"q_{key}")
            price = cC.number_input("Ціна", 0, 1000000, int(EQUIPMENT_BASE[cat][item]), key=f"p_{key}")
            sub = qty * price
            cD.write(f"**{sub:,}** грн")
            st.session_state.selected_items[key] = {"name": item, "qty": qty, "p": price, "sum": sub, "cat": cat}

# ================== ГЕНЕРАЦІЯ ==================
all_items = [v_it for k_it, v_it in st.session_state.selected_items.items() if any(k_it.startswith(c) for c in list(EQUIPMENT_BASE.keys()))]

if all_items:
    st.divider()
    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        full_date_ukr = f"{date_val.day} { {1:'січня',2:'лютого',3:'березня',4:'квітня',5:'травня',6:'червня',7:'липня',8:'серпня',9:'вересня',10:'жовтня',11:'листопада',12:'грудня'}[date_val.month]} {date_val.year} року"
        safe_addr = re.sub(r'[\\/*?:"<>|]', "", address).replace(" ", "_")
        
        # Спільні реквізити для всіх файлів
        base_reps = {
            "vendor_name": v["full"], "vendor_address": v["adr"], "vendor_inn": v["inn"],
            "vendor_iban": v["iban"], "vendor_email": email, "vendor_short_name": v["short"],
            "customer": customer, "address": address, "kp_num": kp_num, "date": date_str,
            "manager": manager, "phone": phone, "email": email, "txt_intro": txt_intro,
            "line1": l1, "line2": l2, "line3": l3
        }
        
        files_results = {}

        # 1. КП (template.docx)
        if os.path.exists("template.docx"):
            doc_kp = Document("template.docx")
            replace_placeholders_stable(doc_kp, base_reps)
            tbl = next((t for t in doc_kp.tables if "Найменування" in t.rows[0].cells[0].text), doc_kp.tables[0])
            total_raw = 0
            for it in all_items:
                row = tbl.add_row().cells
                row[0].text, row[1].text = it['name'], str(it['qty'])
                row[2].text, row[3].text = f"{it['p']:,}".replace(",", " "), f"{it['sum']:,}".replace(",", " ")
                total_raw += it['sum']
            
            tax_val = int(total_raw * v['tax_rate'])
            # Рядок податку
            r_tax = tbl.add_row().cells
            r_tax[0].text = v['tax_label']
            r_tax[0].merge(r_tax[2]); r_tax[3].text = f"{tax_val:,}".replace(",", " ")
            # Рядок Разом
            r_total = tbl.add_row().cells
            r_total[0].text = "ЗАГАЛЬНА ВАРТІСТЬ З УРАХУВАННЯМ ПОДАТКІВ, грн"
            r_total[0].merge(r_total[2]); r_total[3].text = f"{total_raw + tax_val:,}".replace(",", " ")
            
            buf_kp = BytesIO(); doc_kp.save(buf_kp); buf_kp.seek(0)
            files_results["kp"] = {"name": f"КП_{kp_num}_{safe_addr}.docx", "data": buf_kp}

        # 2. Специфікація Поставки (товари)
        hw = [i for i in all_items if "роботи" not in i["cat"].lower()]
        if hw and os.path.exists("template_postavka.docx"):
            doc_p = Document("template_postavka.docx")
            sum_p = sum(i['sum'] for i in hw)
            tax_p = int(sum_p * v['tax_rate'])
            final_p = sum_p + tax_p
            
            reps_p = base_reps.copy()
            reps_p.update({"spec_id_postavka": f"№1 від {full_date_ukr}", "total_sum_digits": f"{final_p:,}", "total_sum_words": amount_to_text_uk(final_p)})
            replace_placeholders_stable(doc_p, reps_p)
            
            tbl_p = doc_p.tables[0]
            for it in hw:
                r = tbl_p.add_row().cells
                r[0].text, r[1].text, r[2].text, r[3].text = it['name'], str(it['qty']), f"{it['p']:,}", f"{it['sum']:,}"
            
            # Рядки підсумків у специфікації
            rt = tbl_p.add_row().cells
            rt[0].text = v['tax_label']
            rt[0].merge(rt[2]); rt[3].text = f"{tax_p:,}".replace(",", " ")
            rf = tbl_p.add_row().cells
            rf[0].text = "РАЗОМ"
            rf[0].merge(rf[2]); rf[3].text = f"{final_p:,}".replace(",", " ")

            buf_p = BytesIO(); doc_p.save(buf_p); buf_p.seek(0)
            files_results["p"] = {"name": f"Spec_Postavka_{kp_num}.docx", "data": buf_p}

        # 3. Специфікація Робіт (послуги)
        wrk = [i for i in all_items if "роботи" in i["cat"].lower()]
        if wrk and os.path.exists("template_roboti.docx"):
            doc_w = Document("template_roboti.docx")
            sum_w = sum(i['sum'] for i in wrk)
            tax_w = int(sum_w * v['tax_rate'])
            final_w = sum_w + tax_w
            
            reps_w = base_reps.copy()
            reps_w.update({"spec_id_roboti": f"№1 від {full_date_ukr}", "total_sum_words": amount_to_text_uk(final_w)})
            replace_placeholders_stable(doc_w, reps_w)
            # Фікс тегу адреси (іноді там зайві пробіли)
            for p in doc_w.paragraphs:
                if "{{ address }}" in p.text or "{{  address }}" in p.text:
                    p.text = p.text.replace("{{ address }}", address).replace("{{  address }}", address)

            tbl_w = doc_w.tables[0]
            for it in wrk:
                r = tbl_w.add_row().cells
                r[0].text, r[1].text, r[2].text, r[3].text = it['name'], str(it['qty']), f"{it['p']:,}", f"{it['sum']:,}"
            
            rt = tbl_w.add_row().cells
            rt[0].text = v['tax_label']
            rt[0].merge(rt[2]); rt[3].text = f"{tax_w:,}".replace(",", " ")
            rf = tbl_w.add_row().cells
            rf[0].text = "РАЗОМ"
            rf[0].merge(rf[2]); rf[3].text = f"{final_w:,}".replace(",", " ")

            buf_w = BytesIO(); doc_w.save(buf_w); buf_w.seek(0)
            files_results["w"] = {"name": f"Spec_Roboti_{kp_num}.docx", "data": buf_w}

        # Запис у Google Sheets
        save_to_google_sheets([date_str, kp_num, customer, address, (sum(i['sum'] for i in all_items) + int(sum(i['sum'] for i in all_items)*v['tax_rate'])), manager])
        
        st.session_state.generated_files = files_results
        st.rerun()

# Відображення кнопок
if st.session_state.generated_files:
    st.write("### 📂 Завантажити документи:")
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(label=f"💾 {info['name']}", data=info['data'], file_name=info['name'], key=f"dl_{k}")
