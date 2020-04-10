import mcutils as mc
import xl_generation
from datetime import datetime
import xlwt
import os
import utilities

class output:
    def __init__(self, cod_sap, descripcion, unidad, cantidad, precio_unitario, subtotal, precio_compra, total_precio_compra):
        self.cod_sap = cod_sap
        self.descripcion = descripcion
        self.unidad = unidad
        self.cantidad = cantidad
        self.precio_unitario = precio_unitario
        self.subtotal = subtotal
        self.precio_compra = precio_compra
        self.total_precio_compra = total_precio_compra

def format_excel_sheet(data):

    book = xlwt.Workbook('output/output.xsl')
    ids = list(map(lambda x: x.cod_sap, data))
    ids = list(set(ids))
    current_sheet = book.add_sheet("{}".format(datetime.strftime(data[0].fecha_entrega, "%b %Y")))

    orders = []
    for i in ids:
        group = list(filter(lambda x: x.cod_sap == i, data))

        cod_sap = group[0].cod_sap
        descripcion = group[0].descripcion
        unidad = group[0].unidad
        precio_unitario = group[0].precio_unitario
        precio_compra = group[0].precio_compra
        if precio_compra == 0: mc.log("No se ha encontrado un precio compra de producto ({}:{})".format(cod_sap, descripcion))

        cantidad = sum(map(lambda x: x.cantidad, group))
        subtotal = sum(map(lambda x: x.subtotal, group))
        total_precio_compra = sum(map(lambda x: x.total_precio_compra, group))

        o = output(cod_sap=cod_sap, descripcion=descripcion, unidad=unidad, cantidad=cantidad,
                   precio_unitario=precio_unitario, subtotal=subtotal, precio_compra=precio_compra,
                   total_precio_compra=total_precio_compra)

        orders.append(o)
    orders = sorted(orders, key=lambda x: x.descripcion)
    max_char = {}

    for o, row_counter in zip(orders, range(1, len(orders) + 1)):

        row = current_sheet.row(row_counter)
        headers = [o.cod_sap, o.descripcion, o.unidad, o.cantidad, o.precio_unitario, o.subtotal, o.precio_compra,
                   o.total_precio_compra]

        for header, index in zip(headers, range(len(headers))):
            row.write(index, header)

            if max_char.get(index) == None:
                max_char[index] = 0

            if max_char[index] < len(str(header)):
                max_char[index] = len(str(header))


    row_counter += 1
    row = current_sheet.row(row_counter)
    headers = ['Cod Sap', 'Descripcion', 'Unidad', 'Cantidad', 'Prec. Unit.', 'Sub Total', 'Precio Compra', 'Total Precio Compra']
    subtotal_index = headers.index('Sub Total')
    total_precio_compra_index = headers.index('Total Precio Compra')
    row.write(subtotal_index, sum(map(lambda x: x.subtotal, orders)))
    row.write(total_precio_compra_index, sum(map(lambda x: x.total_precio_compra, orders)))


    row = current_sheet.row(0)
    for header, index in zip(headers, range(len(headers))):
        row.write(index, header)

        if max_char.get(index) == None:
            max_char[index] = 0

        if max_char[index] < len(str(header)):
            max_char[index] = len(str(header))

    for col in max_char:
        s_col = current_sheet.col(col)
        s_col.width = xl_generation.get_width(max_char[col])


    current_time = datetime.now()
    file_name = datetime.strftime(current_time, "month output %d%m%Y %H%M%S.xls")
    save_path = os.path.join(utilities.get_json()["output_directory"], file_name)
    book.save(save_path)
    mc.mcprint(text="output file saved at: {}".format(save_path), color=mc.Color.GREEN)

    dir_manager = mc.Directory_Manager()
    mf_open_file = mc.Menu_Function("Abrir archivo {}".format(save_path), dir_manager.open_file, *[save_path])
    mc_open_file = mc.Menu(title="Abrir Archivo?", options=[mf_open_file])
    mc_open_file.show()

def begin(year_report=False):
    mc.mcprint("Se utilizaran los precios del archivo '{}'".format(utilities.get_json()["precios_compra_venta_file"]), color=mc.Color.YELLOW)

    data = xl_generation.retrieve_data()
    if not year_report:
        date = mc.date_generator(day=1)
        accepted_orders = list(filter(lambda x: ((("Aceptada" in x.estado_oc) and (x.cantidad != 0) and not ("No Aceptada" in x.estado_oc)) or
                                                 ("Nueva Orden" in x.estado_oc) or
                                                 ("En Proceso" in x.estado_oc))
                                                and x.cantidad != 0
                                                and x.fecha_entrega.year == date.year
                                                and x.fecha_entrega.month == date.month, data))
    else:
        date = mc.date_generator(day=1, month=1)
        accepted_orders = list(filter(lambda x: ((("Aceptada" in x.estado_oc) and (x.cantidad != 0) and not ("No Aceptada" in x.estado_oc)) or
                                                 ("Nueva Orden" in x.estado_oc) or
                                                 ("En Proceso" in x.estado_oc))
                                                and x.cantidad != 0
                                                and x.fecha_entrega.year == date.year, data))


    format_excel_sheet(accepted_orders)
