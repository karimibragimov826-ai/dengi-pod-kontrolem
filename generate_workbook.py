from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

OUT = "dengi_pod_kontrolem.xlsx"
wb = Workbook()
dash = wb.active
dash.title = "Главная"
ops = wb.create_sheet("Операции")
budget = wb.create_sheet("Бюджет")
goals = wb.create_sheet("Цели")
guide = wb.create_sheet("Инструкция")

navy = "18251F"
green = "B8F34A"
pale = "F3F7EE"
muted = "66736A"
white = "FFFFFF"
red = "FCE8E6"
line = "DCE5DA"
thin = Side(style="thin", color=line)
money_fmt = '#,##0.00 [$₽-ru-RU];[Red]-#,##0.00 [$₽-ru-RU]'
date_fmt = "DD.MM.YYYY"

def title(ws, text, end_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    c = ws.cell(1,1,text)
    c.font = Font(name="Aptos Display", size=18, bold=True, color=white)
    c.fill = PatternFill("solid", fgColor=navy)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 34

def header_row(ws, row, start, end):
    for col in range(start, end+1):
        c = ws.cell(row,col)
        c.font = Font(bold=True, color=navy)
        c.fill = PatternFill("solid", fgColor=green)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = Border(bottom=thin)
    ws.row_dimensions[row].height = 28

def base(ws):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

# Dashboard
base(dash)
title(dash, "ДЕНЬГИ ПОД КОНТРОЛЕМ — ГЛАВНАЯ", 4)
dash["A3"] = "Выбранный месяц (гггг-мм)"
dash["A3"].font = Font(bold=True, color=navy)
dash["B3"] = "2026-11"
dash["B3"].fill = PatternFill("solid", fgColor=pale)
dash["B3"].font = Font(bold=True, color=navy)
metrics = [
    ("Доходы", '=SUMIFS(Операции!$E$5:$E$204,Операции!$B$5:$B$204,"Доход",Операции!$F$5:$F$204,$B$3)'),
    ("Расходы", '=SUMIFS(Операции!$E$5:$E$204,Операции!$B$5:$B$204,"Расход",Операции!$F$5:$F$204,$B$3)'),
    ("Остаток", "=B5-B6"),
    ("Доля накоплений", "=IFERROR(B7/B5,0)")
]
for r,(label,formula) in enumerate(metrics,5):
    dash.cell(r,1,label).font = Font(bold=True,color=navy)
    dash.cell(r,2,formula).font = Font(bold=True,size=13,color=navy)
    dash.cell(r,1).fill = PatternFill("solid", fgColor=pale)
    dash.cell(r,2).fill = PatternFill("solid", fgColor=pale)
    dash.cell(r,2).number_format = "0.0%" if r == 8 else money_fmt
dash["A10"] = "Быстрый старт"
dash["A10"].font = Font(bold=True,size=13,color=navy)
for r,txt in enumerate([
    "1. Записывайте доходы и расходы на вкладке «Операции».",
    "2. Установите лимиты на вкладке «Бюджет».",
    "3. Добавьте цели и обновляйте накопленные суммы.",
    "4. Меняйте месяц выше, чтобы посмотреть другой период."
],11):
    dash.cell(r,1,txt)
dash.column_dimensions["A"].width = 52
dash.column_dimensions["B"].width = 22
dash.column_dimensions["C"].width = 18
dash.column_dimensions["D"].width = 18

# Transactions
base(ops)
title(ops, "ОПЕРАЦИИ", 6)
ops.merge_cells("A2:F2")
ops["A2"] = "Добавляйте каждое поступление и расход. Суммы указывайте положительными числами."
ops["A2"].font = Font(italic=True,color=muted)
headers = ["Дата","Тип","Категория","Описание","Сумма, ₽","Месяц"]
for c,h in enumerate(headers,1): ops.cell(4,c,h)
header_row(ops,4,1,6)
for r in range(5,205):
    ops.cell(r,6, f'=IF(A{r}="","",TEXT(A{r},"yyyy-mm"))')
    ops.cell(r,1).number_format = date_fmt
    ops.cell(r,5).number_format = money_fmt
    for c in range(1,7):
        ops.cell(r,c).border = Border(bottom=thin)
type_dv = DataValidation(type="list", formula1='"Доход,Расход"', allow_blank=True)
cat_dv = DataValidation(type="list", formula1='"Еда,Жильё,Транспорт,Покупки,Подписки,Здоровье,Развлечения,Образование,Другое,Зарплата,Подарки,Прочее"', allow_blank=True)
ops.add_data_validation(type_dv); type_dv.add("B5:B204")
ops.add_data_validation(cat_dv); cat_dv.add("C5:C204")
for col,width in {"A":16,"B":16,"C":20,"D":34,"E":18,"F":16}.items(): ops.column_dimensions[col].width=width
ops.auto_filter.ref = "A4:F204"

# Monthly budget
base(budget)
title(budget, "БЮДЖЕТ НА МЕСЯЦ", 5)
budget.merge_cells("A2:E2")
budget["A2"] = "Укажите плановые лимиты. Фактические расходы считаются за месяц, выбранный на вкладке «Главная»."
budget["A2"].font = Font(italic=True,color=muted)
for c,h in enumerate(["Категория","План, ₽","Факт, ₽","Осталось, ₽","Статус"],1): budget.cell(4,c,h)
header_row(budget,4,1,5)
cats = ["Еда","Жильё","Транспорт","Покупки","Подписки","Здоровье","Развлечения","Образование","Другое"]
for r,cat in enumerate(cats,5):
    budget.cell(r,1,cat)
    budget.cell(r,2,0)
    budget.cell(r,3,f'=SUMIFS(Операции!$E$5:$E$204,Операции!$C$5:$C$204,A{r},Операции!$B$5:$B$204,"Расход",Операции!$F$5:$F$204,Главная!$B$3)')
    budget.cell(r,4,f'=B{r}-C{r}')
    budget.cell(r,5,f'=IF(D{r}<0,"Превышен лимит","В пределах плана")')
    for c in range(2,5): budget.cell(r,c).number_format = money_fmt
    for c in range(1,6): budget.cell(r,c).border = Border(bottom=thin)
budget.conditional_formatting.add("D5:D13", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid",fgColor=red)))
for col,width in {"A":22,"B":18,"C":18,"D":18,"E":26}.items(): budget.column_dimensions[col].width=width

# Savings goals
base(goals)
title(goals, "ЦЕЛИ НАКОПЛЕНИЙ", 6)
goals.merge_cells("A2:F2")
goals["A2"] = "Записывайте, на что копите. Обновляйте сумму накоплений — прогресс пересчитается автоматически."
goals["A2"].font = Font(italic=True,color=muted)
for c,h in enumerate(["Цель","Нужно, ₽","Накоплено, ₽","Осталось, ₽","Прогресс","Срок"],1): goals.cell(4,c,h)
header_row(goals,4,1,6)
for r in range(5,15):
    goals.cell(r,4,f'=IF(B{r}="","",MAX(B{r}-C{r},0))')
    goals.cell(r,5,f'=IFERROR(C{r}/B{r},0)')
    goals.cell(r,5).number_format = "0%"
    goals.cell(r,2).number_format = money_fmt
    goals.cell(r,3).number_format = money_fmt
    goals.cell(r,4).number_format = money_fmt
    goals.cell(r,6).number_format = date_fmt
    for c in range(1,7): goals.cell(r,c).border = Border(bottom=thin)
goals.conditional_formatting.add("E5:E14", DataBarRule(start_type="num",start_value=0,end_type="num",end_value=1,color=green,showValue=True))
for col,width in {"A":28,"B":18,"C":18,"D":18,"E":16,"F":16}.items(): goals.column_dimensions[col].width=width

# Guide
base(guide)
title(guide, "ИНСТРУКЦИЯ", 2)
guide_rows = [
    ("Что это","Простая таблица для учёта доходов, расходов, бюджета и целей накоплений."),
    ("Шаг 1 — Операции","Введите дату, выберите тип и категорию, добавьте описание и сумму."),
    ("Шаг 2 — Бюджет","Укажите плановые лимиты. Фактические расходы считаются за месяц, выбранный на вкладке «Главная»."),
    ("Шаг 3 — Цели","Добавьте название цели, нужную сумму, текущие накопления и при желании срок."),
    ("Шаг 4 — Главная","Выберите месяц в формате гггг-мм, например 2026-11."),
    ("Важно","Это инструмент учёта бюджета, а не финансовая, инвестиционная, налоговая или юридическая рекомендация."),
    ("Совместимость","Файл предназначен для Excel. При импорте в Google Таблицы проверьте формулы и даты.")
]
for r,(a,b) in enumerate(guide_rows,3):
    guide.cell(r,1,a).font = Font(bold=True,color=navy)
    guide.cell(r,1).fill = PatternFill("solid",fgColor=pale)
    guide.cell(r,2,b).alignment = Alignment(wrap_text=True,vertical="top")
    guide.row_dimensions[r].height = 36
guide.column_dimensions["A"].width = 24
guide.column_dimensions["B"].width = 100
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.alignment = Alignment(vertical="center", wrap_text=cell.alignment.wrap_text)
wb.calculation = CalcProperties(calcMode="auto")
wb.save(OUT)
print(f"Created {OUT}")
