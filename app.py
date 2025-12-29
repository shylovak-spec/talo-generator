import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
def format_num(n):
    return f"{math.ceil(n):,}".replace(",", " ")

def set_cell_style(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold

def amount_to_text_uk(amount):
    val = math.ceil(amount)
    if num2words is None: return f"{format_num(val)} грн."
    try:
        words = num2words(val, lang='uk').capitalize()
        return f"{words} гривень 00 копійок"
    except: return f"{format_num(val)} грн."

def replace_headers_styled(doc, reps):
    """Робить назву поля ЖИРНОЮ, а значення - звичайним"""
    fields = ["Комерційна пропозиція:", "Дата:", "Замовник:", "Адреса:", "Виконавець:", "Відповідальний:", "Контактний телефон:", "E-mail:"]
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
        for f in fields:
            if f in p.text:
                val = mapping.get(f, "")
                p.text = "" # Очищуємо
                r1 = p.add_run(f + " ")
                r1.bold = True
                r2 = p.add_run(str(val))
                r2.bold = False

def fill_spec_table(tbl, items, tax_label, tax_rate):
    """Заповнення з групуванням та захистом від помилок індексу"""
    groups = {
        "ОБЛАДНАННЯ": ["Інвертори Deye", "Акумулятори (АКБ)"],
        "МАТЕРІАЛИ ТА КОМПЛЕКТУЮЧІ": ["Комплектуючі та щити"],
        "ПОСЛУГИ ТА РОБОТИ": ["Послуги та Роботи"]
    }

    grand_pure = 0
    col_count = len(tbl.columns)

    for g_name, g_cats in groups.items():
        g_items = [it for it in items if it['cat'] in g_cats]
        if not g_items:
            continue

        # Заголовок групи
        row = tbl.add_row().cells
        if col_count >= 2:
            row[0].merge(row[col_count - 1])
        set_cell_style(row[0], g_name, WD_ALIGN_PARAGRAPH.CENTER, True)

        for it in g_items:
            r = tbl.add_row().cells
            set_cell_style(r[0], f"- {it['name']}")

            if col_count >= 4:
                set_cell_style(r[1], it['qty'], WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_style(r[2], format_num(it['p']), WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_style(r[3], format_num(it['sum']), WD_ALIGN_PARAGRAPH.RIGHT)

            grand_pure += it['sum']

    tax_val = math.ceil(grand_pure * tax_rate)
    total_total = grand_pure + tax_val

    # Підсумкові рядки
    for label, val in [
        ("РАЗОМ, грн:", grand_pure),
        (f"{tax_label}:", tax_val),
        ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", total_total)
    ]:
        row = tbl.add_row().cells

        if col_count >= 4:
            row[0].merge(row[col_count - 2])
            set_cell_style(row[0], label, bold=True)
            set_cell_style(row[col_count - 1], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, True)

        elif col_count == 3:
            row[0].merge(row[1])
            set_cell_style(row[0], label, bold=True)
            set_cell_style(row[2], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, True)

        else:  # 1–2 колонки
            set_cell_style(row[0], f"{label} {format_num(val)}", bold=True)

    return total_total

# ================== STREAMLIT UI ==================
st.set_page_config(page_title="Talo Generator", layout="wide")
st.title("⚡ ТАЛО: Генератор")

if "selected_items" not in st.session_state: st.session_state.selected_items = {}

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_choice = col1.selectbox("Виконавець:", list(VENDORS.keys()))
    v = VENDORS[vendor_choice]
    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса")
    kp_num = col2.text_input("Номер КП", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_val = col2.date_input("Дата")

# Інтерфейс вибору як на скрінах
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        sel = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"m_{cat}")
        for name in sel:
            key = f"{cat}_{name}"
            if key not in st.session_state.selected_items:
                st.session_state.selected_items[key] = {"name": name, "qty": 1, "p": int(EQUIPMENT_BASE[cat][name]), "cat": cat}
            
            # Рядок товару
            c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
            c1.info(name)
            st.session_state.selected_items[key]['qty'] = c2.number_input("К-сть", 1, 100, st.session_state.selected_items[key]['qty'], key=f"q_{key}")
            st.session_state.selected_items[key]['p'] = c3.number_input("Ціна", 0, 1000000, st.session_state.selected_items[key]['p'], key=f"p_{key}")
            st.session_state.selected_items[key]['sum'] = st.session_state.selected_items[key]['qty'] * st.session_state.selected_items[key]['p']
            c4.metric("Сума", format_num(st.session_state.selected_items[key]['sum']))

# Очистка видалених
active_keys = [f"{cat}_{n}" for cat in EQUIPMENT_BASE for n in (st.session_state.get(f"m_{cat}") or [])]
st.session_state.selected_items = {k: v for k, v in st.session_state.selected_items.items() if k in active_keys}

items = list(st.session_state.selected_items.values())
if items and st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary"):
    reps = {"vendor_name": v["full"], "customer": customer, "address": address, "kp_num": kp_num, "date": date_val.strftime("%d.%m.%Y"), "manager": manager, "phone": "+380 (67) 477-17-18", "vendor_email": "o.kramarenko@talo.com.ua"}
    
    results = {}
    configs = {
        "kp": {"tpl": "template.docx", "name": f"КП_{kp_num}.docx", "filter": lambda x: True},
        "p": {"tpl": "template_postavka.docx", "name": f"Spec_Postavka_{kp_num}.docx", "filter": lambda x: "роботи" not in x['cat'].lower()},
        "w": {"tpl": "template_roboti.docx", "name": f"Spec_Roboti_{kp_num}.docx", "filter": lambda x: "роботи" in x['cat'].lower()}
    }

    for k, cfg in configs.items():
        if os.path.exists(cfg["tpl"]):
            doc = Document(cfg["tpl"])
            replace_headers_styled(doc, reps)
            f_items = [i for i in items if cfg["filter"](i)]
            if f_items:
                tbl = doc.tables[0]
                total = fill_spec_table(tbl, f_items, v['tax_label'], v['tax_rate'])
                # Заміна тегів внизу
                for p in doc.paragraphs:
                    if "{{total_sum_digits}}" in p.text: p.text = p.text.replace("{{total_sum_digits}}", format_num(total))
                    if "{{total_sum_words}}" in p.text: p.text = p.text.replace("{{total_sum_words}}", amount_to_text_uk(total))
                
                buf = BytesIO(); doc.save(buf); buf.seek(0)
                results[k] = {"name": cfg["name"], "data": buf}
    
    if results:
        st.success("Готово!")
        for res in results.values():
            st.download_button(res['name'], res['data'], res['name'])
