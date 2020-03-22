import xlrd, xlwt
import io
from itertools import groupby

class product:
    def __init__(self, n_orden_compra, fecha_emision, fecha_entrega, ce_co, rut_proveedor, cod_sap,
                 descripcion, glosa, unidad, cantidad, precio_unitario,
                 subtotal):

        self.n_orden_compra = n_orden_compra
        self.fecha_emision = fecha_emision
        self.fecha_entrega = fecha_entrega
        self.ce_co = ce_co
        self.rut_proveedor = rut_proveedor
        self.cod_sap = cod_sap
        self. descripcion = descripcion
        self. glosa = glosa
        self.unidad = unidad
        self.cantidad = cantidad.replace('.','').replace(',','.')
        self.precio_unitario = precio_unitario.replace('.','').replace(',','.')
        self.subtotal = float(subtotal.replace('.','').replace(',','.'))

data = io.open('data/data.xls', 'r', encoding="utf-16")
data = data.readlines()[12:]

destino = xlwt.Workbook()
sheet1 = destino.add_sheet("Home")

columns = [0,5,6,7,8,10,11,12,13,14,16,19]

product_list = []

for row, row_index in zip(data, range(len(data))):
    xls_row = sheet1.row(row_index)
    row_array = row.strip().split('\t')
    r = row_array

    if(len(r)>= 23 and r[0] != 'N OC'):
        # print(len(r))
        product_list.append(product(r[0],r[5],r[6],r[7],r[8],r[10],r[11],r[12],r[13],r[14], r[16], r[19]))

        column_counter = 0
        for column, column_index in zip(row_array, range(len(row_array))):
            if column_index in columns:
                xls_row.write(column_counter, column)

                column_counter += 1

for item in product_list:
    print("detalle: {},\t{} {}".format(item.descripcion, str((item.subtotal)), item.unidad))
