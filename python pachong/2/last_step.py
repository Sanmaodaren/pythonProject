import xlrd

wb = xlrd.open_workbook('vision2.xls')
for i in range(1, 21):
    ws = wb.sheet_by_name(f'sheet{i}')
    ws.merged_cells