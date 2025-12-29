# ------------------- КП -------------------
if os.path.exists("template.docx"):
    doc_kp = Document("template.docx")
    replace_placeholders_stable(doc_kp, base_reps)
    
    # Жирне форматування ключових слів
    for p in doc_kp.paragraphs:
        for bold_label in ["Комерційна пропозиція:", "Дата:", "Замовник:", "Адреса:", "Виконавець:", "Контактний телефон:"]:
            if bold_label in p.text:
                parts = p.text.split(bold_label)
                p.clear()
                run_bold = p.add_run(bold_label)
                run_bold.bold = True
                run_text = p.add_run(parts[1] if len(parts)>1 else "")
                run_text.bold = False

    tbl = doc_kp.tables[0]

    categories_order = ["Обладнання", "Матеріали", "Комплектуючі", "Роботи"]

    # Додаємо всі товари по категоріях, без "Разом <категорія>"
    for cat in categories_order:
        items_in_cat = [it for it in all_items if it['cat'].lower() == cat.lower()]
        if not items_in_cat:
            continue
        # Заголовок категорії
        row_header = tbl.add_row().cells
        row_header[0].text = cat.upper()
        row_header[0].merge(row_header[3])
        # Рядки з товарами
        for it in items_in_cat:
            r = tbl.add_row().cells
            r[0].text = it['name']
            r[1].text = str(it['qty'])
            r[2].text = f"{it['p']:,}".replace(",", " ")
            r[3].text = f"{it['sum']:,}".replace(",", " ")

    # Підсумки в кінці
    total_sum = sum(it['sum'] for it in all_items)
    tax_total = round(total_sum * v['tax_rate'], 2)
    total_with_tax = total_sum + tax_total

    # Разом
    r_total = tbl.add_row().cells
    r_total[0].text = "Разом, грн"
    r_total[0].merge(r_total[2])
    r_total[3].text = f"{total_sum:,.2f}".replace(",", " ")

    # ПДВ / Податкове навантаження
    r_tax = tbl.add_row().cells
    r_tax[0].text = v['tax_label']
    r_tax[0].merge(r_tax[2])
    r_tax[3].text = f"{tax_total:,.2f}".replace(",", " ")

    # Загальна вартість
    r_final = tbl.add_row().cells
    r_final[0].text = "ЗАГАЛЬНА ВАРТІСТЬ, грн"
    r_final[0].merge(r_final[2])
    r_final[3].text = f"{total_with_tax:,.2f}".replace(",", " ")

    buf_kp = BytesIO(); doc_kp.save(buf_kp); buf_kp.seek(0)
    files_results["kp"] = {"name": f"КП_{kp_num}_{safe_addr}.docx", "data": buf_kp}

# ------------------- Специфікації -------------------
def generate_spec(template_path, items, spec_name):
    if not items or not os.path.exists(template_path):
        return None
    doc = Document(template_path)
    reps = base_reps.copy()
    total_sum = sum(it['sum'] for it in items)
    tax_total = round(total_sum * v['tax_rate'], 2)
    total_with_tax = total_sum + tax_total
    reps.update({"total_sum_digits": f"{total_with_tax:,.2f}", "total_sum_words": amount_to_text_uk(total_with_tax)})
    replace_placeholders_stable(doc, reps)
    
    tbl = doc.tables[0]

    # Додаємо всі товари по категоріях без "Разом <категорія>"
    for cat in categories_order:
        cat_items = [it for it in items if it['cat'].lower() == cat.lower()]
        if not cat_items:
            continue
        row_header = tbl.add_row().cells
        row_header[0].text = cat.upper()
        row_header[0].merge(row_header[3])
        for it in cat_items:
            r = tbl.add_row().cells
            r[0].text, r[1].text, r[2].text, r[3].text = it['name'], str(it['qty']), f"{it['p']:,}", f"{it['sum']:,}"

    # Разом
    r_total = tbl.add_row().cells
    r_total[0].text = "Разом, грн"
    r_total[0].merge(r_total[2])
    r_total[3].text = f"{total_sum:,.2f}".replace(",", " ")

    # ПДВ / Податкове навантаження
    r_tax = tbl.add_row().cells
    r_tax[0].text = v['tax_label']
    r_tax[0].merge(r_tax[2])
    r_tax[3].text = f"{tax_total:,.2f}".replace(",", " ")

    # Загальна вартість
    r_final = tbl.add_row().cells
    r_final[0].text = "ЗАГАЛЬНА ВАРТІСТЬ, грн"
    r_final[0].merge(r_final[2])
    r_final[3].text = f"{total_with_tax:,.2f}".replace(",", " ")

    buf = BytesIO(); doc.save(buf); buf.seek(0)
    return {"name": spec_name, "data": buf}

# Специфікація Поставки
hw = [i for i in all_items if "роботи" not in i["cat"].lower()]
spec_post = generate_spec("template_postavka.docx", hw, f"Spec_Postavka_{kp_num}.docx")
if spec_post: files_results["p"] = spec_post

# Специфікація Робіт
wrk = [i for i in all_items if "роботи" in i["cat"].lower()]
spec_rob = generate_spec("template_roboti.docx", wrk, f"Spec_Roboti_{kp_num}.docx")
if spec_rob: files_results["w"] = spec_rob
