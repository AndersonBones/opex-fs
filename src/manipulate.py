from openpyxl import load_workbook
from openpyxl.styles import numbers
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment, NamedStyle
import os


file_path = "result.xlsx"

wb = load_workbook(file_path)

# grab the active worksheet
ws = wb['Opex']




def set_font():
    for row in ws["A1:W100"]:
        for cell in row:
            cell.font = Font(name='Montserrat',
                size=10,
                bold=False,
                italic=False,
                vertAlign=None,
                underline='none',
                strike=False,
                color='000000')

def number_format():
    for row in ws["B4:Q47"]:
        for cell in row:
            cell.number_format  = numbers.FORMAT_NUMBER_COMMA_SEPARATED2



def despesas_operacionais():
    for cell in ws["A4:Q4"][0]:
        cell.fill = PatternFill("solid", fgColor="1A7753")
        cell.font = Font(name='Montserrat',
                size=10,
                bold=True,
                italic=False,
                vertAlign=None,
                underline='none',
                strike=False,
                color='ffffff')


def adjust_column():
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.9
        ws.column_dimensions[column_letter].width = adjusted_width


def insert_row():
    for index, cell in enumerate(ws["A4:A80"]):
        if cell[0].value is not None:
            if cell[0].value[0].isalpha() == True:
                ws.insert_rows(index+5, 1)


adjust_column()
set_font()
number_format()
despesas_operacionais()

insert_row()

def set_datetime_header():
    ws["C3"].style = NamedStyle(name='datetime', number_format='mmm-yy')

wb.save(file_path)


