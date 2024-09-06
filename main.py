import openpyxl
from openpyxl import Workbook, load_workbook

book = load_workbook('Microbial Composition.xlsx')
sheet = book.active

for row in range(3,818):
    tax_name = sheet['A' + str(row)]
    print(tax_name)
book.save('Microbial Composition Fixed.xlsx')