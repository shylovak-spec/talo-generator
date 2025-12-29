import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import re
import os
import math

# Спробуємо імпортувати num2words
try:
    from num2words import num2words
except ImportError:
    num2words = None

# ================== НАЛАШТУВАННЯ ТА ДАНІ ==================
VENDORS = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВ «ТАЛО»",
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
def format_num(n):
    """Форматування числа: 1000 -> 1 000"""
    return f"{math.ceil(n):,}".replace(",", " ")

def set_cell_style(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    """Налаштування стилю тексту в клітинці таблиці"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.bold = bold

def amount_to_text_uk(amount):
    """Сума словами (ціле число)"""
    val = math.ceil(amount)
    if num2words is None: return f"{format_num(val)} грн."
    try:
        words = num2words(val, lang='uk').capitalize()
        return f"{words} гривень 00 копійок"
    except: return f"{format_num(val)} грн."

def replace_headers_styled(doc, reps):
    """Заміна заголовків у шапці: Жирний ключ + Звичайне значення"""
    mapping = {
        "Комерційна пропозиція:": reps.get("kp_num", ""),
        "Дата:": reps.get("date", ""),
        "Замовник:": reps.get("customer", ""),
        "Адреса:": reps.get("address", ""),
        "Виконавець:": reps.get("vendor_name", ""),
        "Відповідальний:": reps.get("manager", ""),
        "Контактний телефон:": reps.get("phone", ""),
        "E-mail:": reps.get("vendor_email", "")
    }
    for p in doc.paragraphs:
        for key, val in mapping.items():
            if key in p.text:
                p.text = "" # Очищення параграфа
                run_key = p.add_run(f"{key} ")
                run_key.bold = True
                run_val = p.add_run(str(val))
                run_val.bold = False

def fill_spec_table(tbl, items, tax_label, tax_rate):
    """Групування товарів та заповнення таблиці (як на скріншоті)"""
    # Відповідність ваших категорій до заголовків у таблиці
    groups = {
        "ОБЛАДНАННЯ": ["Інвертори Deye", "Акумулятори (АКБ)"],
        "МАТЕРІАЛИ ТА КОМПЛЕКТУЮЧІ": ["Комплектуючі та щити"],
        "ПОСЛУГИ ТА РОБОТИ": ["Послуги та Роботи"]
    }
    
    grand_pure = 0
    
    for g_name, g_cats in groups.items():
        # Фільтруємо товари, що належать до поточної групи
        g_items = [it for it in items if it['cat'] in g_cats]
        if not g_items: continue
        
        # Рядок-заголовок групи (наприклад, ОБЛАДНАННЯ)
        row = tbl.add_row().cells
        row[0].merge(row[3])
        set_cell_style(row[0], g_name, WD_ALIGN_PARAGRAPH.CENTER, True)
        
        for it in g_items:
            r = tbl.add_row().cells
            set_cell_style(r[0], f"- {it['name']}")
            set_cell_style(r[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_style(r[2], format_num(it['p']), WD_ALIGN_PARAGRAPH.RIGHT)
            set_cell_style(r[3], format_num(it['sum']), WD_ALIGN_PARAGRAPH.RIGHT)
            grand_pure += it['sum']

    # Фінальні розрахунки
    tax_val = math.ceil(grand_pure * tax_rate)
    total_total = grand_pure + tax_val
    
    # Додаємо підсумкові рядки
    for label, value in [
        ("РАЗОМ, грн:", grand_pure),
        (f"{tax_label}:", tax_val),
        ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", total_total)
    ]:
        row = tbl.add_row().cells
        row[0].merge(row[2])
        set_cell_style(row[0], label, bold=True)
        set_cell_style(row[3], format_num(value), WD_ALIGN_PARAGRAPH.RIGHT, True)
    
    return total_total

# ================== ІНТЕРФЕЙС STREAMLIT ==================
st.set_page_config(page_title="Talo Generator", layout="wide")
st.title("⚡ ТАЛО: Генератор документів")

if "selected_items" not in st.session_state: st.session_state.selected_items = {}
if "generated_files" not in st.session_state: st.session_state.generated_files = None

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_choice = col1.selectbox("Виконавець:", list(VENDORS.keys()))
    v = VENDORS[vendor_choice]
    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса об'єкта")
    kp_num = col2.text_input("Номер КП", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_val = col2.date_input("Дата", datetime.date.today())

# Вибір товарів
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        sel = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"m_{cat}")
        for name in sel:
            key = f"{cat}_{name}"
            if key not in st.session_state.selected_items:
                st.session_state.selected_items[key] = {"name": name, "qty": 1, "p": int(EQUIPMENT_BASE[cat][name]), "cat": cat}
            
            c_n, c_q, c_p, c_s = st.columns([3, 1, 1.2, 1])
            c_n.markdown(f"<div style='padding-top:25px'>{name}</div>", unsafe_allow_html=True)
            st.session_state.selected_items[key]['qty'] = c_q.number_input("Кількість", 1, 500, st.session_state.selected_items[key]['qty'], key=f"q_{key}")
            st.session_state.selected_items[key]['p'] = c_p.number_input("Ціна за од.", 0, 1000000, st.session_state.selected_items[key]['p'], key=f"p_{key}")
            st.session_state.selected_items[key]['sum'] = st.session_state.selected_items[key]['qty'] * st.session_state.selected_items[key]['p']
            c_s.markdown(f"<div style='padding-top:30px'><b>{format_num(st.session_state.selected_items[key]['sum'])}</b> грн</div>", unsafe_allow_html=True)

# Синхронізація вибору
active_keys = [f"{cat}_{n}" for cat in EQUIPMENT_BASE for n in (st.session_state.get(f"m_{cat}") or [])]
st.session_state.selected_items = {k: v for k, v in st.session_state.selected_items.items() if k in active_keys}

# ================== ГЕНЕРАЦІЯ ==================
items = list(st.session_state.selected_items.values())
if items:
    st.divider()
    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        reps = {
            "vendor_name": v["full"], "customer": customer, "address": address, 
            "kp_num": kp_num, "date": date_val.strftime("%d.%m.%Y"), "manager": manager,
            "phone": "+380 (67) 477-17-18", "vendor_email": "o.kramarenko@talo.com.ua"
        }
        
        results = {}
        # Конфігурація файлів
        config = {
            "kp": {"tpl": "template.docx", "name": f"КП_{kp_num}.docx", "filter": lambda x: True},
            "postavka": {"tpl": "template_postavka.docx", "name": f"Spec_Postavka_{kp_num}.docx", "filter": lambda x: "роботи" not in x['cat'].lower()},
            "roboti": {"tpl": "template_roboti.docx", "name": f"Spec_Roboti_{kp_num}.docx", "filter": lambda x: "роботи" in x['cat'].lower()}
        }
        
        for key, cfg in config.items():
            if os.path.exists(cfg["tpl"]):
                doc = Document(cfg["tpl"])
                replace_headers_styled(doc, reps)
                
                filtered_items = [i for i in items if cfg["filter"](i)]
                if filtered_items:
                    # Шукаємо таблицю (зазвичай перша)
                    table = doc.tables[0]
                    total_sum = fill_spec_table(table, filtered_items, v['tax_label'], v['tax_rate'])
                    
                    # Фінальна заміна підсумків словами
                    words_reps = {
                        "total_sum_digits": format_num(total_sum),
                        "total_sum_words": amount_to_text_uk(total_sum)
                    }
                    for p in doc.paragraphs:
                        for r_k, r_v in words_reps.items():
                            if f"{{{{{r_k}}}}}" in p.text:
                                p.text = p.text.replace(f"{{{{{r_k}}}}}", r_v)
                    
                    buf = BytesIO(); doc.save(buf); buf.seek(0)
                    results[key] = {"name": cfg["name"], "data": buf}

        st.session_state.generated_files = results
        st.rerun()

if st.session_state.generated_files:
    st.success("✅ Документи сформовано!")
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(label=f"📥 {info['name']}", data=info['data'], file_name=info['name'])
