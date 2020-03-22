import xlrd, xlwt
import io


data = io.open('data/data.xls', 'r', encoding="utf-16")
data = data.readlines()[12:]

destino = xlwt.Workbook()
sheet1 = destino.add_sheet("Home")

columns = [0,5,6,7,8,10,11,12,13,14,16,19]


for row, row_index in zip(data, range(len(data))):
    xls_row = sheet1.row(row_index)
    row_array = row.strip().split('\t')
    column_counter = 0
    for column, column_index in zip(row_array, range(len(row_array))):
        if column_index in columns:
            xls_row.write(column_counter, column)
            column_counter += 1


destino.save('data/output.xls')

data = xlrd.open_workbook('data/output.xls')

