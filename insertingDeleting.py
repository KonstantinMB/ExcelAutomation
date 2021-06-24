from openpyxl import Workbook, load_workbook

wb = load_workbook('me.xlsx')
ws = wb.active

# to insert/delete rows/cols:
ws.insert_rows(5)
ws.insert_cols(3)
ws.delete_rows(4)
ws.delete_cols(2)

# in order to move cols or rows:
ws.move_range("C1:D11", rows=2,cols=2)  # rows=2 means that the data will be moved two rows up
# ...same for cols but to the right


wb.save('me.xlsx')

