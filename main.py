from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment

# Take input
# path_name = input("What is the path of the run to compile the data from? (ex: ../runs/val/[NAME]) \n")
file_name = input("What do you want to name the finished spreadsheet? \n")
# sheet_mode = input("What class is this for? (0 = berries, 1 = bushes) \n")


# Create a workbook and worksheet
wb = Workbook()
ws = wb.active

# Parse the Validation_Metrics file
file_data = []
with open('Validation_Metrics.txt', 'r') as file:
    for line in file:
        line = line.strip().split(',')
        file_data.append(line)
# print(data)

# Format parsed data
data = [
    ['0', file_data[0][0], 12195, '=SUM(D3:D4)', '=SUM(E3:E4)', '=SUM(F3:F4)',
     file_data[0][2], round(float(file_data[0][3]), 4), round(float(file_data[0][4]), 4), round(float(file_data[0][5]), 4), round(float(file_data[0][6]), 4)],
    ['0', file_data[1][0], 979,  '=ROUND(G3*H3, 0)', '=G3-D3', '=ROUND((D3-(D3*I3))/I3, 0)',
     file_data[1][2], round(float(file_data[1][3]), 4), round(float(file_data[1][4]), 4), round(float(file_data[1][5]), 4), round(float(file_data[1][6]), 4)],
    ['0', file_data[2][0], 11216, '=ROUND(G4*H4, 0)', '=G4-D4', '=ROUND((D4-(D4*I4))/I4, 0)',
     file_data[2][2], round(float(file_data[2][3]), 4), round(float(file_data[2][4]), 4), round(float(file_data[2][5]), 4), round(float(file_data[2][6]), 4)]
]

# conf = [0, 0.1, 0.2]

# TOP LEVEL COLUMN HEADING
# ws.append([])

### START OF BODY ###

# Merge headers
ws.merge_cells(start_row=1, start_column=7, end_row=1, end_column=11)
ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=1)

# Add column headings
ws.append(['', '', 'GT', 'TP', 'FP', 'FN', '0', '0', '0', '0', '0'])

# Add data to the worksheet
for row in data:
    ws.append(row)

# Center headers
ws['G1'].alignment = Alignment(horizontal='center')
ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

# Create a table
table = Table(displayName='Table1', ref='A1:K4')

# # Add a default style with striped rows and banded columns
# style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
#                        showLastColumn=False, showRowStripes=True, showColumnStripes=True)
# tab.tableStyleInfo = style

### END OF BODY ###

# Add the table to the worksheet
ws.add_table(table)

# Save workbook
wb.save(file_name + '.xlsx')

# Gets rid of an error
table.autoFilter = None

# Save workbook again to update autoFilter setting
wb.save(file_name + '.xlsx')

# Close workbook
wb.close()
