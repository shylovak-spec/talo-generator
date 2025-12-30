import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os
import math

try:
    from num2words import num2words
except ImportError:
    num2words = None

# ================== НАЛАШТУВАННЯ ТА ДАНІ ==================
VENDORS = {
    "ТОВ «ТАЛО»": {
        "full": "ТОВ «ТАЛО»",  # Виправлено згідно п.3
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
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold

def replace_headers_styled(doc, reps):
    """
    Виконує заміну тегів у тексті та гарантовано робить ключові поля ЖИРНИМИ.
    """
    bold_labels = [
        "Комерційна пропозиція:", "Дата:", "Замовник:", 
        "Адреса:", "Виконавець:", "Контактний телефон:", 
        "Відповідальний:", "E-mail:"
    ]
    
    for p in doc.paragraphs:
        # КРОК 1: Спочатку замінюємо всі теги на значення
        # Ми робимо це ДО форматування, щоб отримати готовий текст
        for key, val in reps.items():
            if f"{{{{{key}}}}}" in p.text:
                # Використовуємо .replace для рядка, це збереже текст, але скине стиль runs
                # Це нормально, бо ми зараз його переробимо нижче
                p.text = p.text.replace(f"{{{{{key}}}}}", str(val))
        
        # КРОК 2: Шукаємо ключові слова і накладаємо жирний шрифт
        for label in bold_labels:
            if label in p.text:
                # Зберігаємо повний текст, який вже містить дані (наприклад, "Замовник: ОСББ")
                full_text = p.text
                
                # Очищуємо параграф від старого форматування
                p.clear()
                
                # Розбиваємо текст на дві частини: До і Після лейблу
                # label - це, наприклад, "Замовник:"
                # parts[1] - це все, що йде після двокрапки
                parts = full_text.split(label, 1)
                
                # Додаємо сам лейбл (ЖИРНИМ)
                run_label = p.add_run(label)
                run_label.bold = True
                
                # Додаємо текст після лейблу (ЗВИЧАЙНИМ)
                if len(parts) > 1:
                    # parts[1] містить пробіл і значення, наприклад " ОСББ"
                    run_value = p.add_run(parts[1])
                    run_value.bold = False
                
                # Перериваємо цикл по лейблах для цього параграфа, 
                # щоб не обробляти один рядок двічі
                break

    # Заміна в таблицях (без форматування жирним)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in reps.items():
                        if f"{{{{{key}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{key}}}}}", str(val))

def fill_document_table(tbl, items, tax_label, tax_rate):
    """
    Універсальна функція для заповнення таблиць (КП та Специфікації)
    з групуванням на 4 категорії (п.1, п.2).
    """
    # Логіка розподілу категорій
    # Ключі словника - це назви розділів у таблиці.
    # Значення - функції або умови для фільтрації items.
    
    def get_category_name(item_cat):
        c = item_cat.lower()
        if "роботи" in c or "послуги" in c: return "РОБОТИ"
        if "комплект" in c or "щит" in c or "кріплення" in c: return "КОМПЛЕКТУЮЧІ"
        if "матеріал" in c or "кабель" in c or "провід" in c: return "МАТЕРІАЛИ"
        return "ОБЛАДНАННЯ" # Все інше - обладнання

    # Групуємо товари
    grouped_items = {"ОБЛАДНАННЯ": [], "МАТЕРІАЛИ": [], "КОМПЛЕКТУЮЧІ": [], "РОБОТИ": []}
    
    grand_pure = 0
    
    for it in items:
        cat_key = get_category_name(it['cat'])
        grouped_items[cat_key].append(it)
        grand_pure += it['sum']

    # Порядок виводу секцій
    sections_order = ["ОБЛАДНАННЯ", "МАТЕРІАЛИ", "КОМПЛЕКТУЮЧІ", "РОБОТИ"]
    col_count = len(tbl.columns)

    for section in sections_order:
        sec_items = grouped_items[section]
        if not sec_items: continue
        
        # Рядок заголовку категорії (Жирний, по центру)
        row_h = tbl.add_row().cells
        if col_count >= 4:
            row_h[0].merge(row_h[col_count-1])
        
        set_cell_style(row_h[0], section, WD_ALIGN_PARAGRAPH.CENTER, False)
        
        # Товари
        for it in sec_items:
            r = tbl.add_row().cells
            set_cell_style(r[0], it['name'], WD_ALIGN_PARAGRAPH.LEFT)
            if col_count >= 4:
                set_cell_style(r[1], str(it['qty']), WD_ALIGN_PARAGRAPH.CENTER) # Кількість по центру
                set_cell_style(r[2], format_num(it['p']), WD_ALIGN_PARAGRAPH.RIGHT) # Ціна вправо
                set_cell_style(r[3], format_num(it['sum']), WD_ALIGN_PARAGRAPH.RIGHT) # Сума вправо

    # Розрахунки підсумків
    tax_val = math.ceil(grand_pure * tax_rate)
    total_val = grand_pure + tax_val

    # Рядки підсумків (Разом, Податок, Загальна)
    footer_rows = [
        ("РАЗОМ, грн:", grand_pure, False),
        (f"{tax_label}:", tax_val, False),
        ("ЗАГАЛЬНА ВАРТІСТЬ, грн:", total_val, True) # True означає жирний рядок
    ]

    for label, val, is_bold in footer_rows:
        row = tbl.add_row().cells
        # Об'єднуємо комірки для назви (0, 1, 2)
        if col_count >= 4:
            row[0].merge(row[2])
            set_cell_style(row[0], label, WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            set_cell_style(row[3], format_num(val), WD_ALIGN_PARAGRAPH.RIGHT, is_bold)
        else:
            set_cell_style(row[0], f"{label} {format_num(val)}", WD_ALIGN_PARAGRAPH.LEFT, is_bold)
            
    return total_val

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
    
    # Використовуємо правильну назву для відображення (п.3)
    display_vendor_name = v["full"]

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
        
        # Синхронізація
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
                col_n.markdown(f"<div style='padding-top: 5px;'>{name}</div>", unsafe_allow_html=True)
                
                edit_qty = col_q.number_input("К-сть", 1, 100, 1, key=f"q_in_{key}", label_visibility="collapsed")
                edit_price = col_p.number_input("Ціна", 0, 1000000, base_price, key=f"p_in_{key}", label_visibility="collapsed")
                
                current_sum = edit_qty * edit_price
                col_s.markdown(f"<div style='padding-top: 5px;'><b>{format_num(current_sum)}</b> грн</div>", unsafe_allow_html=True)
                
                st.session_state.selected_items[key] = {
                    "name": name, "qty": edit_qty, "p": edit_price, "sum": current_sum, "cat": cat
                }

# ================== ГЕНЕРАЦІЯ ==================
all_items = list(st.session_state.selected_items.values())

if all_items:
    st.divider()
    total_pure = sum(it["sum"] for it in all_items)
    tax_amount = math.ceil(total_pure * v['tax_rate'])
    total_with_tax = total_pure + tax_amount
    
    st.info(f"💵 Сума: {format_num(total_pure)} грн | 📑 {v['tax_label']}: {format_num(tax_amount)} грн | 🚀 **РАЗОМ: {format_num(total_with_tax)} грн**")

    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        safe_addr = re.sub(r'[\\/*?:"<>|]', "", address).replace(" ", "_")
        
        base_reps = {
            "vendor_name": display_vendor_name, 
            "vendor_address": v["adr"], "vendor_inn": v["inn"],
            "vendor_iban": v["iban"], "vendor_email": email, "vendor_short_name": v["short"],
            "customer": customer, "address": address, "kp_num": kp_num, "date": date_str,
            "manager": manager, "phone": phone, "email": email, "txt_intro": txt_intro,
            "line1": l1, "line2": l2, "line3": l3,
            "total_sum_digits": format_num(total_with_tax),
            "total_sum_words": amount_to_text_uk(total_with_tax),
            "tax_label": v['tax_label'],
            "tax_amount_val": format_num(tax_amount)
        }
        
        files_results = {}

        # 1. КП (template.docx)
        if os.path.exists("template.docx"):
            doc_kp = Document("template.docx")
            replace_headers_styled(doc_kp, base_reps) # Заміна з жирними заголовками
            
            # Шукаємо таблицю
            tbl = next((t for t in doc_kp.tables if len(t.rows)>0 and "Найменування" in t.rows[0].cells[0].text), doc_kp.tables[0])
            
            fill_document_table(tbl, all_items, v['tax_label'], v['tax_rate'])
            
            buf_kp = BytesIO(); doc_kp.save(buf_kp); buf_kp.seek(0)
            files_results["kp"] = {"name": f"КП_{kp_num}_{safe_addr}.docx", "data": buf_kp}

        # 2. Специфікація Поставки
        # Фільтруємо: НЕ роботи
        hw = [i for i in all_items if "роботи" not in i["cat"].lower()]
        if hw and os.path.exists("template_postavka.docx"):
            doc_p = Document("template_postavka.docx")
            
            # Локальна сума для цього документа
            local_sum = sum(i['sum'] for i in hw)
            local_total = local_sum + math.ceil(local_sum * v['tax_rate'])
            
            reps_p = base_reps.copy()
            reps_p.update({
                "spec_id_postavka": f"№1 від {date_str}", 
                "total_sum_digits": format_num(local_total), 
                "total_sum_words": amount_to_text_uk(local_total)
            })
            
            replace_headers_styled(doc_p, reps_p)
            tbl_p = doc_p.tables[0]
            fill_document_table(tbl_p, hw, v['tax_label'], v['tax_rate'])
            
            buf_p = BytesIO(); doc_p.save(buf_p); buf_p.seek(0)
            files_results["p"] = {"name": f"Spec_Postavka_{kp_num}.docx", "data": buf_p}

        # 3. Специфікація Робіт
        # Фільтруємо: ТІЛЬКИ роботи
        wrk = [i for i in all_items if "роботи" in i["cat"].lower()]
        if wrk and os.path.exists("template_roboti.docx"):
            doc_w = Document("template_roboti.docx")
            
            local_sum = sum(i['sum'] for i in wrk)
            local_total = local_sum + math.ceil(local_sum * v['tax_rate'])
            
            reps_w = base_reps.copy()
            reps_w.update({
                "spec_id_roboti": f"№1 від {date_str}", 
                "total_sum_words": amount_to_text_uk(local_total)
            })
            
            replace_headers_styled(doc_w, reps_w)
            tbl_w = doc_w.tables[0]
            fill_document_table(tbl_w, wrk, v['tax_label'], v['tax_rate'])
            
            buf_w = BytesIO(); doc_w.save(buf_w); buf_w.seek(0)
            files_results["w"] = {"name": f"Spec_Roboti_{kp_num}.docx", "data": buf_w}

        st.session_state.generated_files = files_results
        st.rerun()

if st.session_state.generated_files:
    st.write("### 📂 Завантажити документи:")
    cols = st.columns(len(st.session_state.generated_files))
    for i, (k, info) in enumerate(st.session_state.generated_files.items()):
        cols[i].download_button(label=f"💾 {info['name']}", data=info['data'], file_name=info['name'], key=f"dl_{k}")
