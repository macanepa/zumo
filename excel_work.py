import io, sys, xlwt, xlrd, os, json
from datetime import datetime, timedelta


class order_product:
    def __init__(self, n_orden_compra, fecha_emision, fecha_entrega, ce_co, rut_proveedor, cod_sap,
                 descripcion, glosa, unidad, cantidad, precio_unitario,
                 subtotal, precio_compra):

        self.n_orden_compra = (n_orden_compra)
        self.fecha_emision = datetime.strptime(fecha_emision, '%d-%m-%Y')
        self.fecha_entrega = datetime.strptime(fecha_entrega, '%d-%m-%Y')
        self.ce_co = ce_co
        self.rut_proveedor = rut_proveedor
        self.cod_sap = cod_sap
        self.descripcion = descripcion
        self.glosa = glosa
        self.unidad = unidad
        self.cantidad = float(cantidad.replace('.','').replace(',','.'))
        self.precio_unitario = float(precio_unitario.replace('.','').replace(',','.'))
        self.subtotal = float(subtotal.replace('.','').replace(',','.'))
        self.precio_compra = float(precio_compra)

class order:
    def __init__(self, n_order_id, product_list):
        self.n_order_id = n_order_id
        self.products = product_list


def get_dates_of_week():
    today = datetime.now().date() + timedelta(days=7)
    cant_days = today.weekday()
    start_date = today - timedelta(days=cant_days)
    list_days = []

    for day in range(7):
        week_day = start_date + timedelta(days=day)
        list_days.append(datetime.strftime(week_day, "%d-%b-%Y"))

    return list_days

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def get_prices():

    with open("zumo.json", "r") as config_file:
        data = json.load(config_file)
        path = data["save_directory"]

    workbook = xlrd.open_workbook(path)
    current_sheet = workbook.sheet_by_index(0)

    current_row = 6  # data begins here
    row = current_sheet.row(current_row)

    dict = {}
    while row[0].value != "":
        id = str(int(row[0].value))
        price = int(row[4].value)
        row = current_sheet.row(current_row)
        dict[id] = price

        current_row += 1

    return dict


def retrieve_data():

    precios_venta = get_prices()


    # retrieve data from xls file
    data_path = 'data/data.xls'
    data_file = io.open((data_path), 'r', encoding="utf-16")
    data = data_file.readlines()[12:]
    data_file.close()
    os.remove(data_path)

    order_product_list = []
    s_precios_venta = 0

    for row  in data:
        row_array = row.strip().split('\t')


        if(len(row_array)>= 23 and row_array[0] != 'N OC'):
            if precios_venta.get(str(row_array[10])) != None:
                s_precio_compra = precios_venta["{}".format(row_array[10])]

            order_product_list.append(order_product(row_array[0], row_array[5], row_array[6], row_array[7], row_array[8], row_array[10],
                                              row_array[11], row_array[12], row_array[13], row_array[14], row_array[16], row_array[19], s_precio_compra))

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

def rename_file():

    os.chdir('data')
    for item in os.listdir():
        os.rename(item, 'data.xls')
    os.chdir('..')

def get_width(num_characters):
    return int((1+num_characters) * 275)

def format_excel_sheet(sorted_data):

    book = xlwt.Workbook(('data/output.xsl'))
    next_week_dates = get_dates_of_week()
    max_char = {}

    for daily_data, day_of_week in zip(sorted_data, range(len(sorted_data))):


            # group by order number
            order_id_list = list(map(lambda x: x.n_orden_compra, daily_data))
            order_id_list = list(dict.fromkeys(order_id_list))
            # print(order_id_list)


            order_list = []
            for order_id in order_id_list:
                product_list = list(filter(lambda x: x.n_orden_compra == order_id, daily_data))
                order_list.append(order(order_id, product_list))


            current_sheet = book.add_sheet('{}'.format(next_week_dates[day_of_week]))
            elements_glosa = []

            row_counter = 1
            # separate by white cell each order
            for n in order_list:
                for i in n.products:
                    row = current_sheet.row(row_counter)
                    headers = [i.n_orden_compra, datetime.strftime(i.fecha_emision, '%d-%m-%Y'),
                               datetime.strftime(i.fecha_entrega, '%d-%m-%Y'), i.ce_co, i.rut_proveedor,
                               i.cod_sap, i.descripcion, i.glosa, i.unidad, i.cantidad, i.precio_unitario, i.subtotal, i.precio_compra]


                    if(i.glosa != ""):
                        elements_glosa.append(i)

                    for header, index in zip(headers, range(len(headers))):
                        row.write(index, header)


                        if max_char.get(index) == None:
                            max_char[index] = 0

                        if max_char[index] < len(str(header)):
                            max_char[index] = len(str(header))

                    row_counter += 1

                row_counter += 1

            headers = ['NOC', 'Fecha Emision', 'Fecha Entrega', 'CeCo', 'Rut Proveedor', 'Cod Sap',
                           'Descripcion', 'Glosa', 'Unidad', 'Cantidad', 'Prec. Unit.', 'Sub Total', 'Precio Compra']

            row_counter += 1
            for element in elements_glosa:
                row = current_sheet.row(row_counter)
                row.write(0,element.n_orden_compra)
                row.write(1,element.glosa)
                row_counter += 1

            row = current_sheet.row(0)
            for header, index in zip(headers, range(len(headers))):
                row.write(index, header)


            for col in max_char:
                s_col = current_sheet.col(col)
                s_col.width = get_width(max_char[col])

            max_char = {}

            for i in [7,8,9]:
                col = current_sheet.col(i)
                col.width = 1000



    book.save(('data/output.xls'))


def begin():
    rename_file()
    data = retrieve_data()
    sorted_data = sort_by_date(data)
    format_excel_sheet(sorted_data)
