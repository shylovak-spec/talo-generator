FORM_VERSION = "v_final_ultra_stable"
import streamlit as st
from database import EQUIPMENT_BASE
import datetime
from docx import Document
from io import BytesIO
import re
import os

try:
    from num2words import num2words
except ImportError:
    num2words = None

st.set_page_config(page_title="Talo КП Generator", layout="wide", page_icon="⚡")

# ================== ДОПОМІЖНІ ФУНКЦІЇ ==================
def amount_to_text_uk(amount):
    if num2words is None: return f"{amount} грн."
    units, cents = divmod(int(round(amount * 100)), 100)
    try:
        words = num2words(units, lang='uk').capitalize()
        return f"{words} гривень {cents:02d} копійок"
    except: return f"{amount} грн."

def replace_placeholders_stable(doc, replacements):
    for p in doc.paragraphs:
        for key, value in replacements.items():
            if f"{{{{{key}}}}}" in p.text:
                for run in p.runs:
                    if f"{{{{{key}}}}}" in run.text:
                        run.text = run.text.replace(f"{{{{{key}}}}}", str(value))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in replacements.items():
                        if f"{{{{{key}}}}}" in p.text:
                            for run in p.runs:
                                if f"{{{{{key}}}}}" in run.text:
                                    run.text = run.text.replace(f"{{{{{key}}}}}", str(value))

# ================== БАЗА РЕКВІЗИТІВ ==================
VENDORS_DATA = {
    "ТОВ «ТАЛО»": {
        "full_name": "ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ «ТАЛО»",
        "short_name": "О. КРАМАРЕНКО", "inn": "45274534", 
        "adr": "03115, м. Київ, вул. Крамського Івана, 9", 
        "iban": "UA443052990000026004046815601", "tax_label": "ПДВ (20%)", "tax_rate": 0.20
    },
    "ФОП Крамаренко Олексій Сергійович": {
        "full_name": "ФОП Крамаренко Олексій Сергійович",
        "short_name": "Олексій КРАМАРЕНКО", "inn": "3048920896", 
        "adr": "02156 м. Київ, вул. Кіото 9, кв. 40", 
        "iban": "UA423348510000000026009261015", "tax_label": "Податкове навантаження (6%)", "tax_rate": 0.06
    }
}

# ================== ІНТЕРФЕЙС ==================
st.title("⚡ Генератор документів Talo")

with st.expander("📌 Основна інформація", expanded=True):
    col1, col2 = st.columns(2)
    vendor_key = col1.selectbox("Виконавець:", list(VENDORS_DATA.keys()))
    v = VENDORS_DATA[vendor_key]

    customer = col1.text_input("Замовник", "ОСББ")
    address = col1.text_input("Адреса об'єкта")
    kp_num = col2.text_input("Номер КП/Договору", "1223.25")
    manager = col2.text_input("Відповідальний", "Олексій Крамаренко")
    date_val = col2.date_input("Дата", datetime.date.today())
    date_str = date_val.strftime("%d.%m.%Y")
    phone = col2.text_input("Телефон", "+380 (67) 477-17-18")
    email = col2.text_input("E-mail", "o.kramarenko@talo.com.ua")

st.subheader("📦 Товари та послуги")
if "selected_items" not in st.session_state: st.session_state.selected_items = {}

tabs = st.tabs(list(EQUIPMENT_BASE.keys()))
for i, cat in enumerate(EQUIPMENT_BASE.keys()):
    with tabs[i]:
        selected = st.multiselect(f"Обрати з {cat}:", list(EQUIPMENT_BASE[cat].keys()), key=f"s_{cat}")
        for item in selected:
            key = f"{cat}_{item}"
            cA, cB, cC, cD = st.columns([3, 0.8, 1.2, 1])
            qty = cB.number_input("К-сть", 1, 100, 1, key=f"q_{key}")
            price = cC.number_input("Ціна", 0, 1000000, int(EQUIPMENT_BASE[cat][item]), key=f"p_{key}")
            sub = qty * price
            cD.write(f"**{sub:,}** грн")
            st.session_state.selected_items[key] = {"name": item, "qty": qty, "p": price, "sum": sub, "cat": cat}

if st.session_state.selected_items:
    st.divider()
    if st.button("🚀 ЗГЕНЕРУВАТИ ВСІ ДОКУМЕНТИ", type="primary", use_container_width=True):
        full_date = f"{date_val.day} { {1:'січня',2:'лютого',3:'березня',4:'квітня',5:'травня',6:'червня',7:'липня',8:'серпня',9:'вересня',10:'жовтня',11:'листопада',12:'грудня'}[date_val.month]} {date_val.year} року"
        safe_addr = re.sub(r'[\\/*?:"<>|]', "", address).replace(" ", "_")
        all_items = list(st.session_state.selected_items.values())
        results = {}

        # РЕКВІЗИТИ (Мапінг на ваші теги в шаблонах)
        base_replacements = {
            "vendor_name": v["full_name"], "vendor_address": v["adr"], "vendor_inn": v["inn"],
            "vendor_iban": v["iban"], "vendor_email": email, "vendor_short_name": v["short_name"],
            "customer": customer, "address": address, "kp_num": kp_num, "date": date_str,
            "manager": manager, "phone": phone, "email": email
        }

        # 1. ГЕНЕРАЦІЯ КП
        if os.path.exists("template.docx"):
            doc_kp = Document("template.docx")
            replace_placeholders_stable(doc_kp, base_replacements)
            tbl = next((t for t in doc_kp.tables if "Найменування" in t.rows[0].cells[0].text), doc_kp.tables[0])
            total_raw = 0
            for it in all_items:
                row = tbl.add_row().cells
                row[0].text, row[1].text = it['name'], str(it['qty'])
                row[2].text, row[3].text = f"{it['p']:,}".replace(",", " "), f"{it['sum']:,}".replace(",", " ")
                total_raw += it['sum']
            
            tax_val = int(total_raw * v['tax_rate'])
            # Рядок податку для КП (4 колонки)
            r_tax = tbl.add_row().cells
            r_tax[0].text = v['tax_label']
            r_tax[0].merge(r_tax[2])
            r_tax[3].text = f"{tax_val:,}".replace(",", " ")
            # Рядок Разом для КП
            r_total = tbl.add_row().cells
            r_total[0].text = "ЗАГАЛЬНА ВАРТІСТЬ З УРАХУВАННЯМ ПОДАТКІВ, грн"
            r_total[0].merge(r_total[2])
            r_total[3].text = f"{total_raw + tax_val:,}".replace(",", " ")
            
            buf_kp = BytesIO(); doc_kp.save(buf_kp); buf_kp.seek(0)
            results["kp"] = (f"КП_{kp_num}_{safe_addr}.docx", buf_kp)

        # 2. ГЕНЕРАЦІЯ СПЕЦИФІКАЦІЙ (5 КОЛОНОК)
        def process_spec(template_name, items_list, spec_type_id):
            if not items_list or not os.path.exists(template_name): return None
            doc = Document(template_name)
            total_raw = sum(i['sum'] for i in items_list)
            tax_val = int(total_raw * v['tax_rate'])
            final = total_raw + tax_val
            
            reps = base_replacements.copy()
            reps.update({
                f"spec_id_{spec_type_id}": f"№1 від {full_date}",
                "total_sum_digits": f"{final:,}".replace(",", " "),
                "total_sum_words": amount_to_text_uk(final)
            })
            replace_placeholders_stable(doc, reps)
            
            # Фікс тегів у тексті (адреса в роботах)
            for p in doc.paragraphs:
                p.text = p.text.replace("{{ address }}", address).replace("{{  address }}", address)

            tbl = doc.tables[0]
            for it in items_list:
                row = tbl.add_row().cells
                row[0].text, row[1].text = it['name'], str(it['qty'])
                row[2].text, row[3].text = f"{it['p']:,}".replace(",", " "), f"{it['sum']:,}".replace(",", " ")
                row[4].text = "без ПДВ" if v['tax_rate'] < 0.1 else "з ПДВ"

            # Підсумки (для таблиці 5 колонок)
            r_tax = tbl.add_row().cells
            r_tax[0].text = v['tax_label']
            r_tax[0].merge(r_tax[2])
            r_tax[3].text = f"{tax_val:,}".replace(",", " ")
            
            r_total = tbl.add_row().cells
            r_total[0].text = "РАЗОМ"
            r_total[0].merge(r_total[2])
            r_total[3].text = f"{final:,}".replace(",", " ")
            
            buf = BytesIO(); doc.save(buf); buf.seek(0)
            return buf

        hw = [i for i in all_items if "роботи" not in i["cat"].lower()]
        res_p = process_spec("template_postavka.docx", hw, "postavka")
        if res_p: results["p"] = (f"Spec_Postavka_{kp_num}.docx", res_p)

        wrk = [i for i in all_items if "роботи" in i["cat"].lower()]
        res_w = process_spec("template_roboti.docx", wrk, "roboti")
        if res_w: results["w"] = (f"Spec_Roboti_{kp_num}.docx", res_w)

        st.session_state.ready_files = results

    if "ready_files" in st.session_state:
        st.write("### 📂 Завантажити документи:")
        cols = st.columns(len(st.session_state.ready_files))
        for i, (k, v) in enumerate(st.session_state.ready_files.items()):
            cols[i].download_button(label=f"💾 {v[0]}", data=v[1], file_name=v[0], key=f"f_{k}")
