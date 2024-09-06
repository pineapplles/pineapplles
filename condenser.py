import openpyxl
from openpyxl import Workbook, load_workbook

book = load_workbook('Microbial Composition by Family.xlsx')
sheet = book.active
stop = 818
genus_names = []
updated_data = []

for row in range(3,stop):
    data = []
    for i in range(0,16):
        data.append(sheet[chr(65+i) + str(row)].value)
    if data[0] in genus_names:
        index = genus_names.index(data[0])
        for j in range(1, 16):
            updated_data[index][j] = updated_data[index][j] + data[j]
    else:
        genus_names.append(data[0])
        updated_data.append(data)

print(len(genus_names))
print(len(updated_data))

for delete in range(3, stop):
    sheet.delete_rows(3)

for new in range(len(updated_data)):
    for k in range(16):
        sheet[chr(65+k) + str(new + 3)].value = updated_data[new][k]

book.save('Microbial Composition Condensed Family.xlsx')