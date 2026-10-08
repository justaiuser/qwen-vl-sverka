# -*- coding: utf-8 -*-
"""
Generates realistic sample documents for testing:
1. act_1c_sample.xlsx (1C:Enterprise Excel format)
2. act_counterparty_sample.pdf (Scanned PDF with table, header, totals)
3. act_counterparty_scan.png (Raster image version of the scan)
"""
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from PIL import Image, ImageDraw, ImageFont
import pymupdf


def generate_1c_excel(filepath: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Акт сверки"
    ws.views.sheetView[0].showGridLines = True

    # Стили
    f_title = Font(name="Arial", size=14, bold=True)
    f_bold = Font(name="Arial", size=9, bold=True)
    f_norm = Font(name="Arial", size=9)
    f_header = Font(name="Arial", size=9, bold=True)
    
    fill_h = PatternFill(start_color="F2F2F2", fill_type="solid")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Заголовок 1С
    ws["B2"] = "Акт сверки"
    ws["B2"].font = f_title
    
    ws["B3"] = "взаимных расчетов за период: 01.01.2024 - 31.03.2024"
    ws["B3"].font = f_bold

    ws["B4"] = 'между ООО "Северный Альянс" и АО "Торговый Дом Восток"'
    ws["B4"].font = f_bold

    ws["B5"] = "по договору: Договор поставки № 14-П от 10.01.2023"
    ws["B5"].font = f_norm

    ws["B7"] = 'Мы, нижеподписавшиеся, ООО "Северный Альянс", с одной стороны, и АО "Торговый Дом Восток", с другой стороны, составили настоящий акт сверки о том, что состояние взаимных расчетов по данным учета следующее:'
    ws["B7"].font = f_norm

    # Шапка таблицы (двойная колонка 1С)
    ws.merge_cells("B9:F9")
    ws["B9"] = 'По данным ООО "Северный Альянс"'
    ws["B9"].font = f_bold
    ws["B9"].alignment = Alignment(horizontal="center")

    headers_1c = ["№ п/п", "Дата", "Документ", "Дебет", "Кредит"]
    for c_i, h in enumerate(headers_1c, start=2):
        cell = ws.cell(row=10, column=c_i, value=h)
        cell.font = f_header
        cell.fill = fill_h
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    # Сальдо начальное
    ws.merge_cells("B11:D11")
    ws["B11"] = "Сальдо начальное на 01.01.2024"
    ws["B11"].font = f_bold
    ws["B11"].border = border
    ws["C11"].border = border
    ws["D11"].border = border
    
    ws["E11"] = 25000.00
    ws["E11"].font = f_bold
    ws["E11"].number_format = '#,##0.00'
    ws["E11"].border = border

    ws["F11"] = ""
    ws["F11"].border = border

    # Операции в 1С
    rows = [
        ("1", "12.01.2024", "Реализация (акт, накладная, УПД) № 15 от 12.01.2024", 112000.00, None),
        ("2", "19.01.2024", "Поступление на расчетный счет № 102 от 19.01.2024", None, 137000.00),
        ("3", "05.02.2024", "Реализация (акт, накладная, УПД) № 48 от 05.02.2024", 85400.00, None),
        ("4", "16.02.2024", "Поступление на расчетный счет № 180 от 16.02.2024", None, 85400.00),
        ("5", "02.03.2024", "Реализация (акт, накладная, УПД) № 72 от 02.03.2024", 48500.00, None), # Расхождение: у контрагента 48000
        ("6", "18.03.2024", "Поступление на расчетный счет № 215 от 18.03.2024", None, 40000.00),
        ("7", "29.03.2024", "Реализация (акт, накладная, УПД) № 99 от 29.03.2024", 64000.00, None), # Только в 1С
    ]

    r_curr = 12
    for r_data in rows:
        ws.cell(row=r_curr, column=2, value=r_data[0]).border = border
        ws.cell(row=r_curr, column=3, value=r_data[1]).border = border
        ws.cell(row=r_curr, column=4, value=r_data[2]).border = border
        
        c_dt = ws.cell(row=r_curr, column=5, value=r_data[3])
        c_dt.border = border
        if r_data[3]: c_dt.number_format = '#,##0.00'

        c_kt = ws.cell(row=r_curr, column=6, value=r_data[4])
        c_kt.border = border
        if r_data[4]: c_kt.number_format = '#,##0.00'

        for c in range(2, 7):
            ws.cell(row=r_curr, column=c).font = f_norm

        r_curr += 1

    # Обороты за период
    ws.merge_cells(f"B{r_curr}:D{r_curr}")
    ws[f"B{r_curr}"] = "Обороты за период"
    ws[f"B{r_curr}"].font = f_bold
    ws[f"B{r_curr}"].alignment = Alignment(horizontal="right")
    
    ws[f"E{r_curr}"] = 309900.00
    ws[f"E{r_curr}"].font = f_bold
    ws[f"E{r_curr}"].number_format = '#,##0.00'

    ws[f"F{r_curr}"] = 262400.00
    ws[f"F{r_curr}"].font = f_bold
    ws[f"F{r_curr}"].number_format = '#,##0.00'

    for c in range(2, 7):
        ws.cell(row=r_curr, column=c).border = border

    r_curr += 1

    # Сальдо конечное
    ws.merge_cells(f"B{r_curr}:D{r_curr}")
    ws[f"B{r_curr}"] = 'Сальдо конечное на 31.03.2024 в пользу ООО "Северный Альянс"'
    ws[f"B{r_curr}"].font = f_bold
    ws[f"B{r_curr}"].alignment = Alignment(horizontal="right")

    ws[f"E{r_curr}"] = 72500.00
    ws[f"E{r_curr}"].font = f_bold
    ws[f"E{r_curr}"].number_format = '#,##0.00'

    ws[f"F{r_curr}"] = ""
    for c in range(2, 7):
        ws.cell(row=r_curr, column=c).border = border

    # Ширина колонок
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 52
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 16

    wb.save(filepath)
    print(f"Created: {filepath}")


def generate_counterparty_scan(image_path: str, pdf_path: str):
    # Создаем реалистичный скан акта контрагента (А4, 1240x1754 px)
    width, height = 1240, 1754
    img = Image.new("RGB", (width, height), color=(253, 253, 251))  # Слегка тепловатая бумага
    d = ImageDraw.Draw(img)

    try:
        font_title = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 24)
        font_bold = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 16)
        font_norm = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 15)
        font_sm = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 13)
    except Exception:
        font_title = font_bold = font_norm = font_sm = ImageFont.load_default()

    # Заголовок
    d.text((width // 2, 70), "АКТ СВЕРКИ ВЗАИМНЫХ РАСЧЕТОВ", fill=(20, 20, 20), font=font_title, anchor="mm")
    d.text((width // 2, 105), "за период: 01.01.2024 - 31.03.2024", fill=(30, 30, 30), font=font_bold, anchor="mm")
    d.text((width // 2, 135), 'между АО "Торговый Дом Восток" и ООО "Северный Альянс"', fill=(40, 40, 40), font=font_norm, anchor="mm")

    # Преамбула
    intro = 'Мы, нижеподписавшиеся, от лица АО "Торговый Дом Восток" составили настоящий акт сверки расчетов по договору поставки № 14-П. По данным бухгалтерского учета АО "Торговый Дом Восток":'
    d.text((70, 180), intro[:80], fill=(50, 50, 50), font=font_sm)
    d.text((70, 202), intro[80:], fill=(50, 50, 50), font=font_sm)

    # Начальное сальдо (для контрагента начальный долг зеркален - Кредит 25000)
    d.text((70, 240), "Сальдо начальное на 01.01.2024: задолженность в пользу ООО \"Северный Альянс\" 25 000.00 руб.", fill=(10, 10, 10), font=font_bold)

    # Координаты таблицы
    x_start = 70
    x_end = 1170
    y_table = 280

    cols = [
        (x_start, 130, "№"),
        (130, 260, "Дата"),
        (260, 770, "Документ"),
        (770, 970, "Дебет (Оплата)"),
        (970, x_end, "Кредит (Поступление)")
    ]

    # Шапка таблицы
    d.rectangle([x_start, y_table, x_end, y_table + 40], fill=(235, 238, 242), outline=(80, 80, 80), width=1)
    for c_left, c_right, c_title in cols:
        d.line([(c_right, y_table), (c_right, y_table + 40)], fill=(80, 80, 80), width=1)
        d.text(((c_left + c_right) // 2, y_table + 20), c_title, fill=(0, 0, 0), font=font_bold, anchor="mm")

    # Строки контрагента:
    # 1. 12.01.2024 Товарная накладная № 15 -> Кт 112 000.00
    # 2. 22.01.2024 Платежное поручение № 102 (дата сдвинута с 19.01 на 22.01) -> Дт 137 000.00
    # 3. 05.02.2024 УПД № 48 -> Кт 85 400.00
    # 4. 16.02.2024 Платежное поручение № 180 -> Дт 85 400.00
    # 5. 02.03.2024 Товарная накладная № 72 -> Кт 48 000.00 (РАСХОЖДЕНИЕ В СУММЕ: 48000 вместо 48500)
    # 6. 18.03.2024 Платежное поручение № 215 -> Дт 40 000.00
    # 7. 25.03.2024 Акт оказания услуг (хранение) № 310 -> Дт 7 200.00 (ТОЛЬКО У КОНТРАГЕНТА)
    # (Документ № 99 от 29.03 на 64 000 у контрагента ОТСУТСТВУЕТ!)

    cp_rows = [
        ("1", "12.01.2024", "Товарная накладная № 15 от 12.01.2024", "", "112 000.00"),
        ("2", "22.01.2024", "Платежное поручение № 102 от 22.01.2024", "137 000.00", ""),
        ("3", "05.02.2024", "УПД № 48 от 05.02.2024", "", "85 400.00"),
        ("4", "16.02.2024", "Платежное поручение № 180 от 16.02.2024", "85 400.00", ""),
        ("5", "02.03.2024", "Товарная накладная № 72 от 02.03.2024", "", "48 000.00"),
        ("6", "18.03.2024", "Платежное поручение № 215 от 18.03.2024", "40 000.00", ""),
        ("7", "25.03.2024", "Акт об оказании услуг хранения № 310", "7 200.00", ""),
    ]

    y_curr = y_table + 40
    row_h = 38
    for r_idx, row in enumerate(cp_rows):
        # Чередование заливки
        if r_idx % 2 == 1:
            d.rectangle([x_start, y_curr, x_end, y_curr + row_h], fill=(248, 249, 251))
        
        d.rectangle([x_start, y_curr, x_end, y_curr + row_h], outline=(150, 150, 150), width=1)
        
        # Разделители колонок
        for _, c_right, _ in cols:
            d.line([(c_right, y_curr), (c_right, y_curr + row_h)], fill=(150, 150, 150), width=1)

        d.text(((cols[0][0] + cols[0][1]) // 2, y_curr + row_h // 2), row[0], fill=(0, 0, 0), font=font_norm, anchor="mm")
        d.text(((cols[1][0] + cols[1][1]) // 2, y_curr + row_h // 2), row[1], fill=(0, 0, 0), font=font_norm, anchor="mm")
        d.text((cols[2][0] + 15, y_curr + row_h // 2), row[2], fill=(0, 0, 0), font=font_norm, anchor="lm")
        
        if row[3]:
            d.text((cols[3][1] - 15, y_curr + row_h // 2), row[3], fill=(0, 0, 0), font=font_norm, anchor="rm")
        if row[4]:
            d.text((cols[4][1] - 15, y_curr + row_h // 2), row[4], fill=(0, 0, 0), font=font_norm, anchor="rm")

        y_curr += row_h

    # Итоги и обороты
    d.rectangle([x_start, y_curr, x_end, y_curr + 42], fill=(235, 238, 242), outline=(80, 80, 80), width=1)
    d.text((cols[2][1] - 20, y_curr + 21), "Обороты за период:", fill=(0, 0, 0), font=font_bold, anchor="rm")
    d.text((cols[3][1] - 15, y_curr + 21), "269 600.00", fill=(0, 0, 0), font=font_bold, anchor="rm")
    d.text((cols[4][1] - 15, y_curr + 21), "245 400.00", fill=(0, 0, 0), font=font_bold, anchor="rm")

    for _, c_right, _ in cols:
        d.line([(c_right, y_curr), (c_right, y_curr + 42)], fill=(80, 80, 80), width=1)

    y_curr += 42

    # Конечное сальдо
    y_curr += 25
    d.text((70, y_curr), "По данным АО \"Торговый Дом Восток\" на 31.03.2024:", fill=(0, 0, 0), font=font_bold)
    y_curr += 25
    d.text((70, y_curr), "Задолженность в пользу АО \"Торговый Дом Восток\" составляет 800.00 руб.", fill=(0, 0, 0), font=font_norm)

    # Подписи и круглая печать
    y_curr += 100
    d.text((100, y_curr), "От АО \"Торговый Дом Восток\":", fill=(0, 0, 0), font=font_bold)
    d.text((700, y_curr), "От ООО \"Северный Альянс\":", fill=(0, 0, 0), font=font_bold)

    y_curr += 35
    d.text((100, y_curr), "Генеральный директор ____________ / Васильев А. И. /", fill=(0, 0, 0), font=font_norm)
    d.text((700, y_curr), "Генеральный директор ____________ / Смирнов Д. П. /", fill=(0, 0, 0), font=font_norm)

    # Рисуем имитацию печати (синий круг)
    stamp_center = (380, y_curr - 10)
    d.ellipse([stamp_center[0] - 60, stamp_center[1] - 60, stamp_center[0] + 60, stamp_center[1] + 60], outline=(40, 80, 180), width=3)
    d.ellipse([stamp_center[0] - 50, stamp_center[1] - 50, stamp_center[0] + 50, stamp_center[1] + 50], outline=(40, 80, 180), width=1)
    d.text((stamp_center[0], stamp_center[1] - 10), "АО ТОРГОВЫЙ ДОМ", fill=(40, 80, 180), font=font_sm, anchor="mm")
    d.text((stamp_center[0], stamp_center[1] + 10), "* ВОСТОК * ОГРН", fill=(40, 80, 180), font=font_sm, anchor="mm")

    # Сохраняем PNG
    img.save(image_path, "PNG")
    print(f"Created: {image_path}")

    # Конвертируем в PDF через PyMuPDF
    doc = pymupdf.open()
    img_doc = pymupdf.open(image_path)
    rect = img_doc[0].rect
    pdfbytes = img_doc.convert_to_pdf()
    img_pdf = pymupdf.open("pdf", pdfbytes)
    page = doc.new_page(width=rect.width, height=rect.height)
    page.show_pdf_page(rect, img_pdf, 0)
    doc.save(pdf_path)
    doc.close()
    print(f"Created: {pdf_path}")


if __name__ == "__main__":
    os.makedirs("samples", exist_ok=True)
    generate_1c_excel(os.path.join("samples", "act_1c_sample.xlsx"))
    generate_counterparty_scan(
        os.path.join("samples", "act_counterparty_scan.png"),
        os.path.join("samples", "act_counterparty_sample.pdf")
    )
