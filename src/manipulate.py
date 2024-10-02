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



def set_subtotais():
    for cell in ws["A4:Q"][0]:

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



def get_subtotal_rows():
    rows = []
    for index, cell in enumerate(ws["A5:A80"]):
        index_account = index+5
        idx = index_account+1
        if cell[0].value is not None:
            if cell[0].value[0].isalpha() == True:
                rows.append(index_account)

    return rows


def insert_row():
    rows = get_subtotal_rows()
    for index, row in enumerate(rows):
        ws.insert_rows(rows[index]+len(rows)-len(rows)+index, amount=1)


insert_row()

adjust_column()
set_font()
number_format()
set_subtotais()


def set_datetime_header():
    ws["C3"].style = NamedStyle(name='datetime', number_format='mmm-yy')

wb.save(file_path)




