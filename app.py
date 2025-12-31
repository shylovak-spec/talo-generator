import streamlit as st
import gspread
import requests
import subprocess
import tempfile
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import datetime
import re
import os
import math
from database import EQUIPMENT_BASE

try:
    from num2words import num2words
except ImportError:
    num2words = None

# ================== НАЛАШТУВАННЯ ТА ДАНІ ==================
VENDORS = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВ «ТАЛО»",
        "short": "Олексій КРАМАРЕНКО",
        "inn": "32670939",
        "adr": "03113, м. Київ, проспект Перемоги, будинок 68/1 офіс 62",
        "iban": "_________",
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
def set_document_font(doc):
    """Встановлює базовий шрифт для всього документа"""
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

def format_num(n):
    return f"{math.ceil(n):,}".replace(",", " ")

def amount_to_text_uk(amount):
    val = math.ceil(amount)
    if num2words is None: return f"{format_num(val)} грн."
    try:
        words = num2words(val, lang='uk').capitalize()
        return f"{words} гривень 00 копійок"
    except: return f"{format_num(val)} грн."

def set_cell_style(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    """Стилізація тексту в комірках таблиці"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

def replace_headers_styled(doc, reps):
    """Заміна тегів з дотриманням жирного шрифту для заголовків"""
    bold_labels = [
        "Комерційна пропозиція:", "Дата:", "Замовник:", 
        "Адреса:", "Виконавець:", "Контактний телефон:", 
        "Відповідальний:", "E-mail:"
    ]
    
    all_paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)

    for p in all_paragraphs:
        for key, val in reps.items():
            if f"{{{{{key}}}}}" in p.text:
                full_text = p.text.replace(f"{{{{{key}}}}}", str(val))
                p.clear()
                run = p.add_run(full_text)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
        
        for label in bold_labels:
            if label in p.text:
                full_text = p.text
                p.clear()
                parts = full_text.split(label, 1)
                run_l = p.add_run(label)
                run_l.bold = True
                run_l.font.name = 'Times New Roman'
                run_l.font.size = Pt(12)
                if len(parts) > 1:
                    run_v = p.add_run(parts[1])
                    run_v.bold = False
                    run_v.font.name = 'Times New Roman'
                    run_v.font.size = Pt(12)
                break

def fill_document_table(tbl, items, tax_label, tax_rate):
    def get_category_name(item_cat):
        c = item_cat.lower()
        if "роботи" in c or "послуги" in c: return "РОБОТИ"
        if "комплект" in c or "щит" in c or "кріплення" in c: return "КОМПЛЕКТУЮЧІ"
        if "матеріал" in c or "кабель" in c or "провід" in c: return "МАТЕРІАЛИ"
        return "ОБЛАДНАННЯ"

    grouped_items = {"ОБЛАДНАННЯ": [], "МАТЕРІАЛИ": [], "КОМПЛЕКТУЮЧІ": [], "РОБОТИ": []}
    grand_pure = 0
    for it in items:
        cat_key = get_category_name(it['cat'])
        grouped_items[cat_key].append(it)
        grand_pure += it['sum']

    sections_order = ["ОБЛАДНАННЯ", "МАТЕРІАЛИ", "КОМПЛЕКТУЮЧІ", "РОБОТИ"]
    col_count = len(tbl.columns)

    for section in sections_order:
        sec_items = grouped_items[section]
        if not sec_items: continue
        row_h = tbl.add_row().cells
        if col_count >= 4: row_h[0].merge(row_h[col_count-1])
        row_h[0].text = "" 
        p = row_h[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(section.upper()) 
        run.italic = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        
        for it in sec_items:
            r = tbl.add_row().cells
            set_cell_style(r[0], it['name'], WD_ALIGN_PARAGRAPH.LEFT)
            if col_count >= 4:
                set_cell_style(r[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(r[2], format_num(it['p']), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(r[3], format_num(it['sum']), WD_ALIGN_PARAGRAPH.RIGHT)

    tax_val = math.ceil(grand_pure * tax_rate)
    total_val = grand_pure + tax_val
    footer_rows = [
        ("РАЗОМ, грн:", grand_pure, False),
        (f"{tax_label}:", tax_val, False),
        ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", total_val, True)
    ]
    for label, val, is_bold in footer_rows:
        row = tbl.add_row().cells
        if col_count >= 4:
            row[0].merge(row[2])
            set_cell_style(row[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            set_cell_style(row[3], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)
    return total_val

def send_to_telegram(file_data, file_name):
    """Відправка файлу керівнику в Telegram"""
    try:
        token = st.secrets["telegram_bot_token"]
        chat_id = st.secrets["telegram_chat_id"]
        url = f"https://api.telegram.org/bot{token}/sendDocument"
        
        # Підготовка файлу
        files = {'document': (file_name, file_data)}
        data = {'chat_id': chat_id, 'caption': f"🚀 Нова комерційна пропозиція!\n📄 Файл: {file_name}"}
        
        response = requests.post(url, data=data, files=files)
        if response.status_code == 200:
            st.success("✅ КП успішно надіслано керівнику в Telegram!")
        else:
            st.error(f"❌ Помилка Telegram: {response.text}")
    except Exception as e:
        st.error(f"❌ Не вдалося відправити файл: {e}")

# ================== GOOGLE SHEETS ФУНКЦІЯ ==================
def save_to_google_sheets(row_data):
    """Підключення до Google Sheets та запис рядка даних через Secrets"""
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Секрети 'gcp_service_account' не знайдено!")
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
        st.error(f"❌ Помилка запису в Google Sheets: {e}")
        return False

def convert_docx_to_pdf(docx_data):
    """Конвертує docx (BytesIO) у pdf (BytesIO) за допомогою LibreOffice"""
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, "temp.docx")
            with open(docx_path, "wb") as f:
                f.write(docx_data.getvalue())
            
            # Команда для конвертації (працює на Linux/Streamlit Cloud)
            subprocess.run([
                'lowriter', '--headless', '--convert-to', 'pdf', 
                '--outdir', tmpdir, docx_path
            ], check=True)
            
            pdf_path = os.path.join(tmpdir, "temp.pdf")
            with open(pdf_path, "rb") as f:
                return BytesIO(f.read())
    except Exception as e:
        st.error(f"❌ Помилка конвертації: {e}")
        return None



# ================== ІНТЕРФЕЙС STREAMLIT ==================
st.set_page_config(page_title="Talo Generator", layout="wide")
st.title("⚡ Генератор КП та Специфікацій")

if "generated_files" not in st.session_state:
    st.session_state.generated_files = None
if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

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

# ================== ВИБІР ТОВАРІВ ==================
st.subheader("📦 Специфікація та редагування")
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected_names = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
        current_cat_keys = [f"{cat}_{name}" for name in selected_names]
        for key in list(st.session_state.selected_items.keys()):
            if key.startswith(f"{cat}_") and key not in current_cat_keys:
                del st.session_state.selected_items[key]
        if selected_names:
            for name in selected_names:
                key = f"{cat}_{name}"
                base_price = int(EQUIPMENT_BASE[cat][name])
                col_n, col_q, col_p, col_s = st.columns([3, 1, 1.2, 1])
                col_n.markdown(f"<div style='padding-top: 5px;'>{name}</div>", unsafe_allow_html=True)
                edit_qty = col_q.number_input("К-сть", 1, 100, 1, key=f"q_in_{key}", label_visibility="collapsed")
                edit_price = col_p.number_input("Ціна", 0, 1000000, base_price, key=f"p_in_{key}", label_visibility="collapsed")
                current_sum = edit_qty * edit_price
                col_s.markdown(f"**{format_num(current_sum)}** грн")
                st.session_state.selected_items[key] = {"name": name, "qty": edit_qty, "p": edit_price, "sum": current_sum, "cat": cat}

all_items = list(st.session_state.selected_items.values())

if all_items:
    st.divider()
    total_pure = sum(it["sum"] for it in all_items)
    tax_amount = math.ceil(total_pure * v['tax_rate'])
    total_with_tax = total_pure + tax_amount
    st.info(f"🚀 **РАЗОМ: {format_num(total_with_tax)} грн**")

    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        safe_addr = re.sub(r'[\\/*?:"<>|]', "", address).replace(" ", "_")
        
        # 1. СТВОРЮЄМО base_reps (Це важливо!)
        base_reps = {
            "vendor_name": v["full"], "vendor_address": v["adr"], "vendor_inn": v["inn"],
            "vendor_iban": v["iban"], "vendor_email": email, "vendor_short_name": v["short"],
            "customer": customer, "address": address, "kp_num": kp_num, "date": date_str,
            "manager": manager, "phone": phone, "email": email, "txt_intro": txt_intro,
            "line1": l1, "line2": l2, "line3": l3,
            "total_sum_digits": format_num(total_with_tax),
            "total_sum_words": amount_to_text_uk(total_with_tax),
            "tax_label": v['tax_label'], "tax_amount_val": format_num(tax_amount)
        }
        
        # 2. ЗАПИС В РЕЄСТР (Ваш новий шматок)
        log_row = [date_str, kp_num, customer, address, vendor_choice, total_with_tax, manager]
        with st.spinner("Записую дані в реєстр..."):
            success = save_to_google_sheets(log_row)
            if success:
                st.toast("✅ Дані додано в Google Sheets!")
            else:
                st.error("❌ Не вдалося записати в таблицю. Перевірте формат ключа в Secrets!")

        files_results = {}
        # 1. КП
        if os.path.exists("template.docx"):
            doc_kp = Document("template.docx")
            set_document_font(doc_kp)
            replace_headers_styled(doc_kp, base_reps)
            tbl = next((t for t in doc_kp.tables if len(t.rows)>0 and "Найменування" in t.rows[0].cells[0].text), doc_kp.tables[0])
            fill_document_table(tbl, all_items, v['tax_label'], v['tax_rate'])
            buf = BytesIO(); doc_kp.save(buf); buf.seek(0)
            files_results["kp"] = {"name": f"КП_{kp_num}_{safe_addr}.docx", "data": buf}

        # 2. Поставка
        hw = [i for i in all_items if "роботи" not in i["cat"].lower()]
        if hw and os.path.exists("template_postavka.docx"):
            doc_p = Document("template_postavka.docx")
            set_document_font(doc_p)
            l_sum = sum(i['sum'] for i in hw)
            l_total = l_sum + math.ceil(l_sum * v['tax_rate'])
            reps_p = base_reps.copy()
            reps_p.update({"spec_id_postavka": f"№1 від {date_str}", "total_sum_digits": format_num(l_total), "total_sum_words": amount_to_text_uk(l_total)})
            replace_headers_styled(doc_p, reps_p)
            fill_document_table(doc_p.tables[0], hw, v['tax_label'], v['tax_rate'])
            buf = BytesIO(); doc_p.save(buf); buf.seek(0)
            files_results["p"] = {"name": f"Spec_Postavka_{kp_num}.docx", "data": buf}

        # 3. Роботи
        wrk = [i for i in all_items if "роботи" in i["cat"].lower()]
        if wrk and os.path.exists("template_roboti.docx"):
            doc_w = Document("template_roboti.docx")
            set_document_font(doc_w)
            l_sum = sum(i['sum'] for i in wrk)
            l_total = l_sum + math.ceil(l_sum * v['tax_rate'])
            reps_w = base_reps.copy()
            reps_w.update({"spec_id_roboti": f"№1 від {date_str}", "total_sum_words": amount_to_text_uk(l_total)})
            replace_headers_styled(doc_w, reps_w)
            fill_document_table(doc_w.tables[0], wrk, v['tax_label'], v['tax_rate'])
            buf = BytesIO(); doc_w.save(buf); buf.seek(0)
            files_results["w"] = {"name": f"Spec_Roboti_{kp_num}.docx", "data": buf}

        st.session_state.generated_files = files_results
        st.rerun()

if st.session_state.generated_files:
    st.write("### 📂 Дії з документами:")
    
    # Кнопки завантаження Word-файлів
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(
            label=f"💾 {info['name']}", 
            data=info['data'], 
            file_name=info['name'], 
            key=f"dl_{k}"
        )

    st.divider()
    
    if "kp" in st.session_state.generated_files:
        st.write("### ✈️ Швидка відправка керівнику (PDF):")
        kp_info = st.session_state.generated_files["kp"]
        
        # Кнопка для PDF-відправки
        if st.button("🚀 Надіслати КП у форматі PDF", use_container_width=True):
            with st.spinner("⏳ Конвертуємо у PDF..."):
                kp_info['data'].seek(0)
                pdf_buffer = convert_docx_to_pdf(kp_info['data'])
                
                if pdf_buffer:
                    # Формуємо ідентичну назву, але з .pdf
                    pdf_name = kp_info['name'].replace(".docx", ".pdf")
                    send_to_telegram(pdf_buffer, pdf_name)
