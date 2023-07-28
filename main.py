from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = Workbook()
ws = wb.active

data = [
    ['0', 'All', 12195, '=SUM(D5:D6)', '', '', 8268, 0.668, 0.392, 0.439, 0.2],
    ['0', 'Blue', 979,  '=ROUND(G5*H5, 0)', '', '', 686, 0.657, 0.331, 0.392, 0.18],
    ['0', 'Green', 11216, '=ROUND(G6*H6, 0)', '', '', 7582, 0.679, 0.453, 0.485, 0.22],
]

# conf = [0, 0.1, 0.2]

# Add column headings
ws.append(['a', 'a', 'GT', 'TP', 'FP', 'FN', 'a0', 'a0', 'a0', 'a0', 'a0'])


# Add data to the worksheet
for row in data:
    ws.append(row)

ws.merge_cells(start_row=1, start_column=7, end_row=1, end_column=11)

# ws.insert_cols()

# Create a table
tab = Table(displayName='Table1', ref='A1:K4')

# # Add a default style with striped rows and banded columns
# style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
#                        showLastColumn=False, showRowStripes=True, showColumnStripes=True)
# tab.tableStyleInfo = style

# Add the table to the worksheet
ws.add_table(tab)

wb.save('file6.xlsx')