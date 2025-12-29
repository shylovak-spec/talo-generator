import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH # Для вирівнювання
import re
import os
import math

# Спробуємо імпортувати num2words для суми словами
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
def format_number(n):
    """Форматує число: 10500 -> 10 500 (без ком, з пробілом)"""
    return f"{math.ceil(n):,}".replace(",", " ")

def set_cell_align(cell, align):
    """Допоміжна функція для вирівнювання тексту в клітинці"""
    for paragraph in cell.paragraphs:
        paragraph.alignment = align

def amount_to_text_uk(amount):
    """Перетворює ціле число у суму словами"""
    total_int = math.ceil(amount)
    if num2words is None: 
        return f"{format_number(total_int)} грн."
    try:
        words = num2words(total_int, lang='uk').capitalize()
        return f"{words} гривень 00 копійок"
    except: 
        return f"{format_number(total_int)} грн."

def replace_placeholders_stable(doc, replacements):
    """Заміна тегів у тексті та таблицях"""
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

# ================== ІНТЕРФЕЙС STREAMLIT ==================
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

# ================== ВИБІР ТА РЕДАГУВАННЯ ТОВАРІВ ==================
st.subheader("📦 Специфікація та редагування")

if "selected_items" not in st.session_state:
    st.session_state.selected_items = {}

tabs = st.tabs(list(EQUIPMENT_BASE.keys()))

for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected_names = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
        
        current_cat_keys = [f"{cat}_{name}" for name in selected_names]
        for key in list(st.session_state.selected_items.keys()):
            if key.startswith(f"{cat}_") and key not in current_cat_keys:
                del st.session_state.selected_items[key]

        if selected_names:
            h1, h2, h3, h4 = st.columns([3, 1, 1.2, 1])
            h1.caption("🏷️ Найменування")
            h2.caption("🔢 Кількість")
            h3.caption("💰 Ціна за од.")
            h4.caption("📈 Сума")

            for name in selected_names:
                key = f"{cat}_{name}"
                base_price = int(EQUIPMENT_BASE[cat][name])
                
                col_n, col_q, col_p, col_s = st.columns([3, 1, 1.2, 1])
                col_n.write(name)
                edit_qty = col_q.number_input("К-сть", 1, 100, 1, key=f"q_{key}", label_visibility="collapsed")
                edit_price = col_p.number_input("Ціна", 0, 1000000, base_price, key=f"p_{key}", label_visibility="collapsed")
                
                row_sum = edit_qty * edit_price
                col_s.markdown(f"**{format_number(row_sum)}** грн")
                
                st.session_state.selected_items[key] = {
                    "name": name, "qty": edit_qty, "p": edit_price, "sum": row_sum, "cat": cat
                }

# ================== ГЕНЕРАЦІЯ ==================
all_items = list(st.session_state.selected_items.values())

if all_items:
    st.divider()
    total_pure = sum(it["sum"] for it in all_items)
    tax_amount = math.ceil(total_pure * v['tax_rate'])
    total_with_tax = total_pure + tax_amount
    
    st.info(f"💵 Сума: {format_number(total_pure)} грн | 📑 {v['tax_label']}: {format_number(tax_amount)} грн | 🚀 **РАЗОМ: {format_number(total_with_tax)} грн**")

    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        full_date_ukr = f"{date_val.day} { {1:'січня',2:'лютого',3:'березня',4:'квітня',5:'травня',6:'червня',7:'липня',8:'серпня',9:'вересня',10:'жовтня',11:'листопада',12:'грудня'}[date_val.month]} {date_val.year} року"
        safe_addr = re.sub(r'[\\/*?:"<>|]', "", address).replace(" ", "_")
        
        base_reps = {
            "vendor_name": v["full"], "vendor_address": v["adr"], "vendor_inn": v["inn"],
            "vendor_iban": v["iban"], "vendor_email": email, "vendor_short_name": v["short"],
            "customer": customer, "address": address, "kp_num": kp_num, "date": date_str,
            "manager": manager, "phone": phone, "email": email, "txt_intro": txt_intro,
            "line1": l1, "line2": l2, "line3": l3,
            "total_sum_digits": format_number(total_with_tax),
            "total_sum_words": amount_to_text_uk(total_with_tax),
            "tax_label": v['tax_label'],
            "tax_amount_val": format_number(tax_amount)
        }
        
        files_results = {}

        # 1. КП (template.docx)
        if os.path.exists("template.docx"):
            doc_kp = Document("template.docx")
            replace_placeholders_stable(doc_kp, base_reps)
            tbl = next((t for t in doc_kp.tables if "Найменування" in t.rows[0].cells[0].text), doc_kp.tables[0])
            for it in all_items:
                row = tbl.add_row().cells
                row[0].text, row[1].text = it['name'], str(it['qty'])
                row[2].text, row[3].text = format_number(it['p']), format_number(it['sum'])
                # Форматування клітинок
                set_cell_align(row[1], WD_ALIGN_PARAGRAPH.CENTER) # Кількість
                set_cell_align(row[2], WD_ALIGN_PARAGRAPH.RIGHT)  # Ціна
                set_cell_align(row[3], WD_ALIGN_PARAGRAPH.RIGHT)  # Сума
            
            r_tax = tbl.add_row().cells
            r_tax[0].text = v['tax_label']
            r_tax[0].merge(r_tax[2]); r_tax[3].text = format_number(tax_amount)
            set_cell_align(r_tax[3], WD_ALIGN_PARAGRAPH.RIGHT)
            
            r_total = tbl.add_row().cells
            r_total[0].text = "ЗАГАЛЬНА ВАРТІСТЬ З УРАХУВАННЯМ ПОДАТКІВ, грн"
            r_total[0].merge(r_total[2]); r_total[3].text = format_number(total_with_tax)
            set_cell_align(r_total[3], WD_ALIGN_PARAGRAPH.RIGHT)
            
            buf = BytesIO(); doc_kp.save(buf); buf.seek(0)
            files_results["kp"] = {"name": f"КП_{kp_num}_{safe_addr}.docx", "data": buf}

        # 2. Поставка (template_postavka.docx)
        hw = [i for i in all_items if "роботи" not in i["cat"].lower()]
        if hw and os.path.exists("template_postavka.docx"):
            doc_p = Document("template_postavka.docx")
            s_p = sum(i['sum'] for i in hw)
            t_p = math.ceil(s_p * v['tax_rate'])
            f_p = s_p + t_p
            reps_p = base_reps.copy()
            reps_p.update({"spec_id_postavka": f"№1 від {full_date_ukr}", "total_sum_digits": format_number(f_p), "total_sum_words": amount_to_text_uk(f_p)})
            replace_placeholders_stable(doc_p, reps_p)
            tbl_p = doc_p.tables[0]
            for it in hw:
                r = tbl_p.add_row().cells
                r[0].text, r[1].text, r[2].text, r[3].text = it['name'], str(it['qty']), format_number(it['p']), format_number(it['sum'])
                set_cell_align(r[1], WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_align(r[2], WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_align(r[3], WD_ALIGN_PARAGRAPH.RIGHT)
            buf_p = BytesIO(); doc_p.save(buf_p); buf_p.seek(0)
            files_results["p"] = {"name": f"Spec_Postavka_{kp_num}.docx", "data": buf_p}

        # 3. Роботи (template_roboti.docx)
        wrk = [i for i in all_items if "роботи" in i["cat"].lower()]
        if wrk and os.path.exists("template_roboti.docx"):
            doc_w = Document("template_roboti.docx")
            s_w = sum(i['sum'] for i in wrk)
            t_w = math.ceil(s_w * v['tax_rate'])
            f_w = s_w + t_w
            reps_w = base_reps.copy()
            reps_w.update({"spec_id_roboti": f"№1 від {full_date_ukr}", "total_sum_digits": format_number(f_w), "total_sum_words": amount_to_text_uk(f_w)})
            replace_placeholders_stable(doc_w, reps_w)
            tbl_w = doc_w.tables[0]
            for it in wrk:
                r = tbl_w.add_row().cells
                r[0].text, r[1].text, r[2].text, r[3].text = it['name'], str(it['qty']), format_number(it['p']), format_number(it['sum'])
                set_cell_align(r[1], WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_align(r[2], WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_align(r[3], WD_ALIGN_PARAGRAPH.RIGHT)
            buf_w = BytesIO(); doc_w.save(buf_w); buf_w.seek(0)
            files_results["w"] = {"name": f"Spec_Roboti_{kp_num}.docx", "data": buf_w}

        st.session_state.generated_files = files_results
        st.rerun()

if st.session_state.generated_files:
    st.write("### 📂 Завантажити документи:")
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(label=f"💾 {info['name']}", data=info['data'], file_name=info['name'], key=f"dl_{k}")
