import mcutils as mc
import xl_generation
from datetime import datetime
import xlwt
import os
import utilities


def format_excel_sheet(data):

    book = xlwt.Workbook('output/output.xsl')
    current_sheet = book.add_sheet("Hoja 1")
    for i, row_counter in zip(data, range(1, len(data) + 1)):

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

        row = current_sheet.row(row_counter)
        headers = [i.n_orden_compra, datetime.strftime(i.fecha_entrega, '%d-%m-%Y'),
                   i.unidad, i.cantidad, i.precio_unitario, i.subtotal, i.precio_compra, i.total_precio_compra]

        for header, index in zip(headers, range(len(headers))):
            row.write(index, header)


    headers = ['NOC', 'Fecha Entrega', 'Unidad', 'Cantidad', 'Prec. Unit.', 'Sub Total', 'Precio Compra', 'Total Precio Compra']
    row = current_sheet.row(0)
    for header, index in zip(headers, range(len(headers))):
        row.write(index, header)

    current_time = datetime.now()
    file_name = datetime.strftime(current_time, "month output %d%m%Y %H%M%S.xls")
    save_path = os.path.join(utilities.get_json()["output_directory"], file_name)
    book.save(save_path)
    mc.mcprint(text="output file saved at: {}".format(save_path), color=mc.Color.GREEN)

    dir_manager = mc.Directory_Manager()
    mf_open_file = mc.Menu_Function("Abrir archivo {}".format(save_path), dir_manager.open_file, *[save_path])
    mc_open_file = mc.Menu(title="Abrir Archivo?", options=[mf_open_file])
    mc_open_file.show()

def begin():

    date = mc.date_generator(day=1)
    data = xl_generation.retrieve_data()
    accepted_orders = list(filter(lambda x: ("Aceptada" in x.estado_oc)
                                            and x.cantidad != 0
                                            and x.fecha_entrega.year == date.year
                                            and x.fecha_entrega.month == date.month, data))

    format_excel_sheet(accepted_orders)
