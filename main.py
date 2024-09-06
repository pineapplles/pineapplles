import openpyxl
from openpyxl import Workbook, load_workbook

book = load_workbook('Microbial Composition.xlsx')
sheet = book.active

for row in range(3,818):
    tax_name = sheet['A' + str(row)].value
    tax_array = tax_name.split(";")

    tax_letter = ['d','p','c','o','f','g','s']

    for i in range(0,7):
        if (tax_array[i][0] == tax_letter[i]):
            sheet[chr(82+i) + str(row)].value = tax_array[i][3:]
            print(tax_array[i][3:])
        else:
            sheet[chr(82+i) + str(row)].value = 'remainder'
            print('remainder')

book.save('Microbial Composition Fixed.xlsx')