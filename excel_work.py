import xlwt
import io
from datetime import datetime

class order_product:
    def __init__(self, n_orden_compra, fecha_emision, fecha_entrega, ce_co, rut_proveedor, cod_sap,
                 descripcion, glosa, unidad, cantidad, precio_unitario,
                 subtotal):

        self.n_orden_compra = (n_orden_compra)
        self.fecha_emision = fecha_emision
        self.fecha_entrega = datetime.strptime(fecha_entrega, '%d-%m-%Y')
        self.ce_co = ce_co
        self.rut_proveedor = rut_proveedor
        self.cod_sap = cod_sap
        self.descripcion = descripcion
        self.glosa = glosa
        self.unidad = unidad
        self.cantidad = cantidad.replace('.','').replace(',','.')
        self.precio_unitario = precio_unitario.replace('.','').replace(',','.')
        self.subtotal = float(subtotal.replace('.','').replace(',','.'))

class order:
    def __init__(self, n_order_id, product_list):
        self.n_order_id = n_order_id
        self.products = product_list

def retrieve_data():

    # retrieve data from xls file
    data = io.open('data/data.xls', 'r', encoding="utf-16")
    data = data.readlines()[12:]

    order_product_list = []

    for row  in data:
        row_array = row.strip().split('\t')

        if(len(row_array)>= 23 and row_array[0] != 'N OC'):
            order_product_list.append(order_product(row_array[0], row_array[5], row_array[6], row_array[7], row_array[8], row_array[10],
                                              row_array[11], row_array[12], row_array[13], row_array[14], row_array[16], row_array[19]))

    return order_product_list

def sort_by_date(order_product_list):

    current_date = datetime.now().date()
    current_year = current_date.year
    current_month = current_date.month

    sorted_by_date = (list(sorted(order_product_list, key=lambda x: x.fecha_entrega)))
    filtered_next_week = list(filter(lambda x: datetime.strftime(x.fecha_entrega, "%W") == str(int(datetime.strftime(datetime.now(), "%W")) + 1), sorted_by_date))

    # Separated by day of the week
    splited = []
    for day_of_week in range(7):
        day = list(filter(lambda x: x.fecha_entrega.isoweekday() == day_of_week + 1, filtered_next_week))
        splited.append(day)

    return splited

def format_excel_sheet(sorted_data):

    book = xlwt.Workbook('data/output.xsl')

    for daily_data, day_of_week in zip(sorted_data, range(len(sorted_data))):

            # group by order number
            order_id_list = list(map(lambda x: x.n_orden_compra, daily_data))
            order_id_list = list(dict.fromkeys(order_id_list))
            print(order_id_list)

            order_list = []
            for order_id in order_id_list:
                product_list = list(filter(lambda x: x.n_orden_compra == order_id, daily_data))
                order_list.append(order(order_id, product_list))


            current_sheet = book.add_sheet('{}'.format(day_of_week))

            row_counter = 0
            # separate by white cell each order
            for n in order_list:
                for i in n.products:
                    row = current_sheet.row(row_counter)
                    headers = [i.n_orden_compra, i.fecha_emision, i.fecha_entrega, i.ce_co,
                                i.rut_proveedor, i.cod_sap, i.descripcion, i.glosa, i.unidad,
                                i.cantidad, i.precio_unitario, i.subtotal]

                    for header, index in zip(headers, range(len(headers))):
                        row.write(index, header)
                    row_counter += 1
                row_counter += 1

    book.save('data/output.xls')



data = retrieve_data()
sorted_data = sort_by_date(data)
format_excel_sheet(sorted_data)

