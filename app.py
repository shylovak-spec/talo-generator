import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os

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
    if num2words is None: return f"{amount:,.2f} грн."
    units, cents = divmod(int(round(amount * 100)), 100)
    try:
        words = num2words(units, lang='uk').capitalize()
        return f"{words} гривень {cents:02d} копійок"
    except: return f"{amount:,.2f} грн."

def replace_placeholders_stable(doc, replacements):
    # Поля, які мають бути жирними (пункт 4)
    bold_keys = ["Комерційна пропозиція:", "Дата:", "Замовник:", "Адреса:", "Виконавець:", "Контактний телефон:"]
    
    for p in doc.paragraphs:
        # Звичайна заміна тегів
        for key, val in replacements.items():
            if f"{{{{{key}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{key}}}}}", str(val))
        
        # Обробка жирного шрифту для заголовків (пункт 4)
        for b_key in bold_keys:
            if b_key in p.text:
                full_text = p.text
                p.clear()
                parts = full_text.split(b_key, 1)
                r1 = p.add_run(b_key)
                r1.bold = True
                if len(parts) > 1:
                    r2 = p.add_run(parts[1])
                    r2.bold = False

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in replacements.items():
                        if f"{{{{{key}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{key}}}}}", str(val))

def fill_table_with_sections(table, items, tax_label, tax_amount, total_with_tax):
    """Функція для заповнення таблиці з розділами (пункт 1 та 2)"""
    # Групування
    sections = {
        "ОБЛАДНАННЯ": ["Інвертори Deye", "Акумулятори (АКБ)", "Інвертори Hoymiles", "Інвертори Victron"],
        "МАТЕРІАЛИ ТА КОМПЛЕКТУЮЧІ": ["Комплектуючі та щити"],
        "ПОСЛУГИ ТА РОБОТИ": ["Послуги та Роботи"]
    }
    
    grand_pure = 0
    
    for section_name, categories in sections.items():
        section_items = [it for it in items if it['cat'] in categories]
        if not section_items: continue
        
        # Додаємо заголовок розділу
        row_h = table.add_row().cells
        row_h[0].merge(row_h[3])
        p = row_h[0].paragraphs[0]
        p.text = section_name
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        
        # Додаємо позиції
        for it in section_items:
            row = table.add_row().cells
            row[0].text = it['name']
            row[1].text = str(it['qty'])
            row[2].text = f"{it['p']:,}".replace(",", " ")
            row[3].text = f"{it['sum']:,}".replace(",", " ")
            grand_pure += it['sum']

    # Рядки підсумку (Разом, Податок, Загальна)
    r_pure = table.add_row().cells
    r_pure[0].text = "РАЗОМ, грн:"
    r_pure[0].merge(r_pure[2])
    r_pure[3].text = f"{grand_pure:,}".replace(",", " ")
    
    r_tax = table.add_row().cells
    r_tax[0].text = tax_label
    r_tax[0].merge(r_tax[2])
    r_tax[3].text = f"{tax_amount:,.2f}".replace(",", " ")
    
    r_total = table.add_row().cells
    r_total[0].text = "ЗАГАЛЬНА ВАРТІСТЬ З УРАХУВАННЯМ ПОДАТКІВ, грн"
    r_total[0].merge(r_total[2])
    r_total[3].text = f"{total_with_tax:,.2f}".replace(",", " ")
    r_total[3].paragraphs[0].runs[0].bold = True

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
    # Пункт 3: ТОВ "ТАЛО" замість повної назви для відображення
    display_vendor_name = "ТОВ «ТАЛО»" if vendor_choice == "ТОВ «ТАЛО»" else vendor_choice
    
    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса об'єкта")
    kp_num = col2.text_input("Номер КП/Договору", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_val = col2.date_input("Дата", datetime.date.today())
    date_str = date_val.strftime("%d.%m.%Y")
    phone = col2.text_input("Телефон", "+380 (67) 477-17-18")
    email = col2.text_input("E-mail", "o.kramarenko@talo.com.ua")

# (Тут логіка вибору товарів у tabs, яку ви надали раніше...)
tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected_names = st.multiselect(f"Додати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"ms_{cat}")
        if selected_names:
            for name in selected_names:
                key = f"{cat}_{name}"
                base_price = int(EQUIPMENT_BASE[cat][name])
                c1, c2, c3, c4 = st.columns([3, 1, 1.2, 1])
                c1.write(name)
                qty = c2.number_input("К-сть", 1, 100, 1, key=f"q_{key}")
                price = c3.number_input("Ціна", 0, 1000000, base_price, key=f"p_{key}")
                cur_sum = qty * price
                c4.write(f"**{cur_sum:,}** грн")
                st.session_state.selected_items[key] = {"name": name, "qty": qty, "p": price, "sum": cur_sum, "cat": cat}

# ================== ГЕНЕРАЦІЯ ==================
all_items = list(st.session_state.selected_items.values())

if all_items and st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
    total_pure = sum(it["sum"] for it in all_items)
    tax_amount = round(total_pure * v['tax_rate'], 2)
    total_with_tax = round(total_pure + tax_amount, 2)
    full_date_ukr = f"{date_val.day} { {1:'січня',2:'лютого',3:'березня',4:'квітня',5:'травня',6:'червня',7:'липня',8:'серпня',9:'вересня',10:'жовтня',11:'листопада',12:'грудня'}[date_val.month]} {date_val.year} року"

    base_reps = {
        "vendor_name": display_vendor_name, 
        "customer": customer, "address": address, "kp_num": kp_num, "date": date_str,
        "manager": manager, "phone": phone, "email": email,
        "total_sum_digits": f"{total_with_tax:,.2f}".replace(",", " "),
        "total_sum_words": amount_to_text_uk(total_with_tax)
    }

    files_results = {}

    # Обробка шаблонів (КП та Специфікації)
    templates = {
        "kp": ("template.docx", f"КП_{kp_num}.docx"),
        "p": ("template_postavka.docx", f"Spec_Postavka_{kp_num}.docx"),
        "w": ("template_roboti.docx", f"Spec_Roboti_{kp_num}.docx")
    }

    for key, (t_file, out_name) in templates.items():
        if os.path.exists(t_file):
            doc = Document(t_file)
            replace_placeholders_stable(doc, base_reps)
            
            # Знаходимо таблицю для заповнення
            target_table = None
            for t in doc.tables:
                if "Найменування" in t.rows[0].cells[0].text:
                    target_table = t
                    break
            
            if target_table:
                fill_table_with_sections(target_table, all_items, v['tax_label'], tax_amount, total_with_tax)
            
            buf = BytesIO(); doc.save(buf); buf.seek(0)
            files_results[key] = {"name": out_name, "data": buf}

    st.session_state.generated_files = files_results
    st.rerun()

if st.session_state.generated_files:
    for k, info in st.session_state.generated_files.items():
        st.download_button(label=f"💾 {info['name']}", data=info['data'], file_name=info['name'], key=f"dl_{k}")
