from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
import csv

wb = Workbook()
ws = wb.active

# path_name = input("What is the path of the run to compile the data from? (ex: ../runs/val/[NAME]) \n")
file_name = input("What do you want to name the finished spreadsheet? \n")
# sheet_mode = input("What class is this for? (0 = berries, 1 = bushes) \n")

reader = []

with open('Validation_Metrics.txt', 'a') as file:
    reader = csv.reader(file, delimiter=',')
    # lines = file.readlines()
    # for index, line in enumerate(lines):
    #     print(line)
    #     data = line.strip().split(',')
    #     print(data)
    #     reader[index] = csv.reader(data)

for r in reader:
    print(r)

fake_input = [
    [8268, 0.668, 0.392, 0.439, 0.2],
    [686, 0.657, 0.331, 0.392, 0.18],
    [7582, 0.679, 0.453, 0.485, 0.22]
]

data = [
    ['0', 'All', 12195, '=SUM(D3:D4)', '=SUM(E3:E4)', '=SUM(F3:F4)',
     fake_input[0][0], fake_input[0][1], fake_input[0][2], fake_input[0][3], fake_input[0][4]],
    ['0', 'Blue', 979,  '=ROUND(G3*H3, 0)', '=G3-D3', '=ROUND((D3-(D3*I3))/I3, 0)',
     fake_input[1][0], fake_input[1][1], fake_input[1][2], fake_input[1][3], fake_input[1][4]],
    ['0', 'Green', 11216, '=ROUND(G4*H4, 0)', '=G4-D4', '=ROUND((D4-(D4*I4))/I4, 0)',
     fake_input[2][0], fake_input[2][1], fake_input[2][2], fake_input[2][3], fake_input[2][4]]
]

# conf = [0, 0.1, 0.2]

# Add column headings
ws.append(['', '', 'GT', 'TP', 'FP', 'FN', '0', '0', '0', '0', '0'])


# Add data to the worksheet
for row in data:
    ws.append(row)

ws.merge_cells(start_row=1, start_column=7, end_row=1, end_column=11)
ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=1)

ws['G1'].alignment = Alignment(horizontal='center')
ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

# ws.insert_cols()

# Create a table
tab = Table(displayName='Table1', ref='A1:K4')

# # Add a default style with striped rows and banded columns
# style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
#                        showLastColumn=False, showRowStripes=True, showColumnStripes=True)
# tab.tableStyleInfo = style

# Add the table to the worksheet
ws.add_table(tab)

wb.save(file_name + '.xlsx')

wb.close()
