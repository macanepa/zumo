from datetime import datetime, timedelta
import io, xlwt, xlrd, os
import utilities
import mcutils as mc


class product:
    def __init__(self, n_orden_compra, fecha_emision, fecha_entrega, ce_co, rut_cliente, cod_sap,
                 descripcion, glosa, unidad, cantidad, precio_unitario,
                 subtotal, precio_compra, estado_oc, contacto=""):

        self.n_orden_compra = (n_orden_compra)
        self.fecha_emision = datetime.strptime(fecha_emision, '%d-%m-%Y')
        self.fecha_entrega = datetime.strptime(fecha_entrega, '%d-%m-%Y')
        self.ce_co = ce_co
        self.rut_cliente = rut_cliente
        self.cod_sap = cod_sap
        self.descripcion = descripcion
        self.glosa = glosa
        self.unidad = unidad
        self.cantidad = float(cantidad.replace('.','').replace(',','.'))
        self.precio_unitario = float(precio_unitario.replace('.','').replace(',','.'))
        self.subtotal = float(subtotal.replace('.','').replace(',','.'))
        self.precio_compra = float(precio_compra)
        self.contacto = contacto
        self.estado_oc = estado_oc
        self.total_precio_compra = self.precio_compra * self.cantidad

    def generate_dictionary(self):
        dictionary = {
            "n_order_compra": self.n_orden_compra,
            "fecha_entrega": datetime.strftime(self.fecha_entrega, "%d-%m-%Y"),
            "fecha_emision": datetime.strftime(self.fecha_emision, "%d-%m-%Y"),
            "ce_co": self.ce_co,
            "rut_cliente": self.rut_cliente,
            "cod_sap": self.cod_sap,
            "descripcion": self.descripcion,
            "glosa": self.glosa,
            "unidad": self.unidad,
            "cantidad": self.cantidad,
            "precio_unitario": self.precio_unitario,
            "subtotal": self.subtotal,
            "precio_compra": self.precio_compra,
            "contacto": self.contacto,
            "total_precio_compra": self.total_precio_compra,
            "estado_oc": self.estado_oc
        }
        return dictionary

class order:
    def __init__(self, n_order_id, product_list):
        self.n_order_id = n_order_id
        self.products = product_list

class Week:
    WEEK_SHIFT = 1

def get_dates_of_week():
    today = datetime.now().date() + timedelta(days=(7*Week.WEEK_SHIFT))
    start_date = today - timedelta(days=today.weekday())

    week_days = []
    for day in range(7):
        week_day = start_date + timedelta(days=day)
        week_days.append(datetime.strftime(week_day, "%d-%b-%Y"))

    return week_days

def generate_chunks(products_in_order):
    max = 10
    length = len(products_in_order.products)
    base_length = length
    rest = 0
    if length > max and length % max != 0:
        whole_part = (length // max)
        length = length - (whole_part - 1) * max
        base_length = (whole_part - 1) * max
        rest = length % max  # 2
        while rest < 5:
            max -= 1
            rest = length % max
    else:
        max = base_length
        base_length = 0

    return (base_length // max, max, rest)

def get_sale_prices():
    data = utilities.get_json()
    path = data["precios_compra_venta_file"]
    if not os.path.isfile(path):
        mc.register_error(error_string="No se ha encontrado el archivo {}".format(path))
        mc.exit_application(enter_quit=True)

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

def get_latest_file(directory):
    dir_list = []
    for file in os.listdir(directory):
        dir_list.append(os.path.join(directory, file))

    if len(dir_list) == 0:
        mc.exit_application(text="No data files found at {}".format(directory))

    latest_file_path = max(dir_list, key=os.path.getmtime)
    return latest_file_path

def retrieve_data():
    sale_prices = get_sale_prices()

    # retrieve data from xls file
    downloads_directory = utilities.get_json()["downloads_directory"]
    data_path = get_latest_file(downloads_directory)
    data_file = io.open((data_path), 'r', encoding="utf-16")
    data = data_file.readlines()[12:]
    data_file.close()

    product_list = []
    for row  in data:
        row_array = row.strip().split('\t')
        if(len(row_array)>= 23 and row_array[0] != 'N OC'):
            temp_precio_compra = 0
            if sale_prices.get(str(row_array[10])) != None:
                temp_precio_compra = sale_prices["{}".format(row_array[10])]

            product_list.append(product(n_orden_compra=row_array[0],
                                        fecha_emision=row_array[5],
                                        fecha_entrega=row_array[6],
                                        ce_co=row_array[7],
                                        rut_cliente=row_array[8],
                                        cod_sap=row_array[10],
                                        descripcion=row_array[11],
                                        glosa=row_array[12],
                                        unidad=row_array[13],
                                        cantidad=row_array[14],
                                        precio_unitario=row_array[16],
                                        subtotal=row_array[19],
                                        precio_compra=temp_precio_compra,
                                        estado_oc=row_array[3]))

    product_list = filter_products(product_list)

    return product_list

def sort_by_date(order_product_list):
    sorted_by_date = (list(sorted(order_product_list, key=lambda x: x.fecha_entrega)))

    mc_select_week = mc.Menu(title="Elegir semana", options=["Proxima Semana", "Esta Semana", "Introducir Fecha Manual"], back=False)
    mc_select_week.show()
    week_index = int(mc_select_week.returned_value)
    if week_index == 2:
        Week.WEEK_SHIFT = 0
        selected_date = datetime.now()
    elif week_index == 1:
        Week.WEEK_SHIFT = 1
        selected_date = datetime.now() + timedelta(days=7)
    else:
        custom_date = mc.date_generator()
        current_date = datetime.now()

        monday1 = (current_date.date() - timedelta(days=current_date.weekday()))
        monday2 = (custom_date.date() - timedelta(days=custom_date.weekday()))

        Week.WEEK_SHIFT = ((monday2 - monday1).days / 7)
        mc.mcprint("WEEK SHIFT: {}".format(Week.WEEK_SHIFT), color=mc.Color.RED)
        selected_date = custom_date


    filtered_next_week = list(filter(lambda x: (datetime.strftime(x.fecha_entrega, "%W") ==  datetime.strftime(selected_date, "%W")) and
                                     (datetime.strftime(x.fecha_entrega, "%Y") == datetime.strftime(selected_date, "%Y")), sorted_by_date))

    # Separated by day of the week
    splited = []
    for day_of_week in range(7):
        day = list(filter(lambda x: x.fecha_entrega.isoweekday() == day_of_week + 1, filtered_next_week))
        splited.append(day)

    return splited

def get_width(num_characters):
    return int((1+num_characters) * 275)

def format_excel_sheet(sorted_data):

    book = xlwt.Workbook('output/output.xsl')
    next_week_dates = get_dates_of_week()
    max_char = {}

    # contiene todos los dias
    week_data_dictionary = {}
    for daily_data, day_of_week in zip(sorted_data, range(len(sorted_data))):

            # group by order number
            order_id_list = list(map(lambda x: x.n_orden_compra, daily_data))
            order_id_list = list(dict.fromkeys(order_id_list))
            order_list = []

            for order_id in order_id_list:
                product_list = list(filter(lambda x: x.n_orden_compra == order_id, daily_data))
                order_list.append(order(order_id, product_list))

            current_sheet = book.add_sheet('{}'.format(next_week_dates[day_of_week]))
            elements_glosa = []
            row_counter = 1 # separate by white cell each order

            # contiene todas las ordenes de un mismo dia
            week_data_dictionary["day_{}".format(day_of_week)] = {}

            for n in order_list:

                # contiene todos los chunks
                week_data_dictionary["day_{}".format(day_of_week)]["order_id_{}".format(n.n_order_id)] = {}

                dist_10, dist_1, dist_last = generate_chunks(n)
                if dist_10==0:
                    dist_10=dist_1
                for i, index in zip(n.products, range(len(n.products))):

                    if index < (dist_10 * 10):
                        current_page = index // 10
                    elif index < dist_10 * 10 + dist_1:
                        current_page = dist_10
                    else:
                        current_page = dist_10 + 1

                    dict_keys =list(week_data_dictionary["day_{}".format(day_of_week)]["order_id_{}".format(n.n_order_id)].keys())
                    if len(dict_keys) == 0:
                        week_data_dictionary["day_{}".format(day_of_week)]["order_id_{}".format(n.n_order_id)][
                            "chunk_{}".format(current_page)] = {}
                    dict_keys = list(week_data_dictionary["day_{}".format(day_of_week)]["order_id_{}".format(n.n_order_id)].keys())
                    if not "chunk_{}".format(current_page) in dict_keys:
                        week_data_dictionary["day_{}".format(day_of_week)]["order_id_{}".format(n.n_order_id)][
                            "chunk_{}".format(current_page)] = {}

                    contact_id = i.n_orden_compra.split("-")[0]
                    contacts_path = utilities.get_json()["resources_directory"]
                    contacts_path = os.path.join(contacts_path, "contacts.json")
                    contact_dictionary = utilities.get_json(contacts_path)

                    if not "{}".format(contact_id) in contact_dictionary:
                        new_contact_dictionary = utilities.add_contact(contact_id)
                        contact_dictionary["{}".format(contact_id)] = new_contact_dictionary

                    contact_data = contact_dictionary["{}".format(contact_id)]["contacto"]
                    rut_cliente = contact_dictionary["{}".format(contact_id)]["rut_cliente"]

                    i.contacto = contact_data
                    i.rut_cliente = rut_cliente


                    week_data_dictionary["day_{}".format(day_of_week)]["order_id_{}".format(n.n_order_id)]["chunk_{}".format(current_page)]["product_{}".format(index)] = i.generate_dictionary()

                    row = current_sheet.row(row_counter)
                    headers = [i.n_orden_compra, datetime.strftime(i.fecha_emision, '%d-%m-%Y'),
                               datetime.strftime(i.fecha_entrega, '%d-%m-%Y'), i.ce_co, i.rut_cliente,
                               i.cod_sap, i.descripcion, i.glosa, i.unidad, i.cantidad, i.precio_unitario, i.subtotal, i.precio_compra, i.total_precio_compra, i.contacto]


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

            headers = ['NOC', 'Fecha Emision', 'Fecha Entrega', 'CeCo', 'Rut Cliente', 'Cod Sap',
                           'Descripcion', 'Glosa', 'Unidad', 'Cantidad', 'Prec. Unit.', 'Sub Total', 'Precio Compra', 'Total Precio Compra', 'Contacto']

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

    convert_data_to_json(week_data_dictionary)

    current_time = datetime.now()
    file_name = datetime.strftime(current_time, "output %d%m%Y %H%M%S.xls")
    save_path = os.path.join(utilities.get_json()["output_directory"], file_name)
    book.save(save_path)
    mc.mcprint(text="output file saved at: {}".format(save_path), color=mc.Color.GREEN)

    dir_manager = mc.Directory_Manager()
    mf_open_file = mc.Menu_Function("Abrir archivo {}".format(save_path), dir_manager.open_file, *[save_path])
    mc_open_file = mc.Menu(title="Abrir Archivo?", options=[mf_open_file])
    mc_open_file.show()

def convert_data_to_json(data):
    path = utilities.get_json()["output_directory"]
    path = os.path.join(path, "temp.json")
    mc.generate_json(path, data)

def filter_products(orders):
    accepted_orders = list(filter(lambda x: ((("Aceptada" in x.estado_oc) and (x.cantidad != 0) and not ("No Aceptada" in x.estado_oc)) or
                                             ("Nueva Orden" in x.estado_oc) or
                                             ("En Proceso" in x.estado_oc)), orders))

    # for order in accepted_orders:
    #     print("Wiwi", order.estado_oc)

    # return orders
    return accepted_orders

def begin():
    data = retrieve_data()
    sorted_data = sort_by_date(data)
    format_excel_sheet(sorted_data)
