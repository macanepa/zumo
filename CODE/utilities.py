#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os, json, webbrowser, sys
import mcutils as mc
import logging

def initialize_config():
    output_directory = os.path.join(os.getcwd(), "output")
    config_directory = os.path.join(os.getcwd())
    resources_directory = os.path.join(os.getcwd(), "resources")
    downloads_directory = os.path.join(os.getcwd(), "downloads")
    directories = [output_directory, config_directory, resources_directory, downloads_directory]

    config_file_directory = os.path.join(config_directory, "zumo.json")

    if not os.path.isfile(config_file_directory):
        logging.info("Creating file: {}".format(config_file_directory))
        config = {
            "precios_compra_venta_file": "{}".format(os.path.join(resources_directory, "PRECIOS COMPRA VENTA.xlsx")),
            "output_directory": "{}".format(output_directory),
            "resources_directory": "{}".format(resources_directory),
            "downloads_directory": "{}".format(downloads_directory),
            "configuration_path": "{}".format(config_file_directory),
            "shortener": {"RESOLUCION": "RES", "CALIBRE": "CAL", "RESOLUCIÓN": "RES"}
        }

        with open("zumo.json", "w") as config_file:
            json.dump(config, config_file)

    # create required folders
    for dir in directories:
        if not os.path.isdir(dir):
            text = "Creating folder: {}".format(dir)
            logging.info(text)
            os.mkdir(dir)


def download_update():
    mf_download_zumo_update = mc.MenuFunction(title='Descargar Actualización de ZUMO',
                                              function=webbrowser.open,
                                              url="https://github.com/macanepa/zumo/releases")
    mf_download_chromedriver_update = mc.MenuFunction(title='Descargar Actualización de Chrome Driver',
                                                       function=webbrowser.open,
                                                       url='https://chromedriver.chromium.org/downloads')

    mc_updates = mc.Menu(title='Descargar Actualizaciones', options=[mf_download_zumo_update,
                                                                     mf_download_chromedriver_update])

    mc_updates.show()



def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def get_json(config_directory=os.path.join(os.getcwd(), "zumo.json")):
    try:
        return mc.get_dict_from_json(config_directory)
    except FileNotFoundError as e:
        text = "No se ha encontrado el archivo {}".format(e.filename)
        logging.error(text)
        mc.exit_application(enter_quit=True)


def import_contacts(path=None):
    if not path:
        path = get_json()["resources_directory"]
        path = os.path.join(path, "LISTA CECO CASINOS.csv")

    contacts_dictionary = {}
    with open(path, "r", encoding="ISO-8859-1") as file:
        for line in file:
            line = line.strip().split(";")
            contacts_dictionary[line[0]] = {"contacto": line[1], "rut_cliente": line[2]}
            mc.mcprint(line, color=mc.Color.GREEN)

    # mc.mcprint(contacts_dictionary, color=mc.Color.YELLOW)
    path = get_json()["resources_directory"]
    path = os.path.join(path, "contacts.json")
    mc.generate_json(path, contacts_dictionary)

    mc.mcprint("\nContactos importados exitosamente", color=mc.Color.GREEN)


def manage_contacts():
    mc_add_contact = mc.MenuFunction("Agregar Nuevo Contacto", add_contact)
    mc_show_contacts = mc.MenuFunction("Mostrar Contactos", show_contacts)
    mc_import_contacts = mc.MenuFunction("Importar Contactos", import_contacts)
    mc_edit_contacts = mc.MenuFunction("Editar Contactos", edit_contact)
    mc_remove_contact = mc.MenuFunction("Eliminar Contacto", remove_contact)
    mc_export_contacts = mc.MenuFunction(
        "{}Exportar Contactos (usar con precaucion){}".format(mc.Color.RED, mc.Color.RESET), export_contact)

    mc_contacts = mc.Menu(title="Gestionar Contactos",
                          options=[mc_add_contact, mc_edit_contacts, mc_show_contacts, mc_import_contacts,
                                   mc_export_contacts, mc_remove_contact])
    mc_contacts.show()


def edit_contact():
    contacts_path = get_json()["resources_directory"]
    contacts_path = os.path.join(contacts_path, "contacts.json")
    contacts_dictionary = get_json(contacts_path)
    contacts_ids = list(contacts_dictionary.keys())

    show_contacts()
    print()

    contact_id = mc.get_input(text="Introduce el id", valid_options=contacts_ids)
    contact = contacts_dictionary[contact_id]
    mc.mcprint(contact, color=mc.Color.CYAN)

    contact_keys = (list(contact.keys()))
    contact_keys.insert(0, "id")

    select_menu = mc.Menu(title="Seleccione item a modificar", options=contact_keys)
    select_menu.show()
    selected_option = select_menu.returned_value

    if selected_option == "0":
        return
    new_value = mc.get_input(text="Indique nuevo valor de {}".format(contact_keys[int(selected_option) - 1]))
    if selected_option == "1":
        contacts_dictionary[new_value] = contacts_dictionary.pop(contact_id)
    else:
        contacts_dictionary[contact_id][contact_keys[int(selected_option) - 1]] = new_value

    mc.generate_json(contacts_path, contacts_dictionary)

    mc.mcprint(text="Contacto modificado exitosamente", color=mc.Color.GREEN)


def add_contact(contact_id=None):
    mc.mcprint("Adding new contact", color=mc.Color.YELLOW)
    default_rut = "94623000-6"

    new_contact_dictionary = {}

    contact_string = ""
    if contact_id == None:
        contact_id = mc.get_input(text="id contacto")
    else:
        contact_string = " id({})".format(contact_id)
    contact_data = mc.get_input(text="Introduce informacion de nuevo contacto{}".format(contact_string))
    contact_rut = mc.get_input(text="Introduce RUT Cliente (dejar en blanco para rut defecto '{}')".format(default_rut))
    if contact_rut == "":
        contact_rut = default_rut

    new_contact_dictionary["contacto"] = contact_data
    new_contact_dictionary["rut_cliente"] = contact_rut

    path = get_json()["resources_directory"]
    path = os.path.join(path, "contacts.json")
    contact_dictionary = get_json(path)
    contact_dictionary[contact_id] = new_contact_dictionary
    mc.generate_json(path, contact_dictionary)

    mc.mcprint("\nid: '{}'\t{}".format(contact_id, new_contact_dictionary), color=mc.Color.YELLOW)
    mc.mcprint("Nuevo contacto ingresado a base de datos exitosamente", color=mc.Color.GREEN)
    return new_contact_dictionary


def show_contacts():
    path = get_json()["resources_directory"]
    path = os.path.join(path, "contacts.json")
    contact_dictionary = get_json(path)

    for key in contact_dictionary:
        mc.mcprint("{}\t{}".format(key, contact_dictionary[key]), color=mc.Color.YELLOW)


def add_contact_info(contact_id):
    mc.mcprint("Agregando contacto {} a la base de datos".format(contact_id), color=mc.Color.GREEN)
    path = get_json()["resources_directory"]
    path = os.path.join(path, "contacts.json")
    contact_dictionary = get_json(path)
    contact_data = mc.get_input(text="Introduce informacion de nuevo contacto", color=mc.Color.YELLOW)
    contact_dictionary[contact_id] = {}
    contact_dictionary[contact_id]["contacto"] = contact_data
    contact_dictionary[contact_id]["rut_cliente"] = "94623000-6"
    mc.generate_json(path, contact_dictionary)
    mc.mcprint(text="contacto {} agregado correctamente a la base de datos\n".format(contact_id), color=mc.Color.GREEN)
    return contact_data


def remove_contact(contact_id=None):
    contacts_path = get_json()["resources_directory"]
    contacts_path = os.path.join(contacts_path, "contacts.json")
    contacts_dictionary = get_json(contacts_path)

    contacts_ids = list(contacts_dictionary.keys())
    show_contacts()
    print()
    if contact_id == None:
        contact_id = mc.get_input(text="Indique el id del contacto a eliminar", valid_options=contacts_ids)
    try:
        del contacts_dictionary[contact_id]
    except:
        logging.error("No se ha econtrador el contacto ({})".format(contact_id))
        return
    mc.generate_json(contacts_path, contacts_dictionary)
    mc.mcprint("Contacto ({}) ha sido eliminado exitosamente".format(contact_id), color=mc.Color.RED)


def export_contact(path=None):
    resource_path = get_json()["resources_directory"]
    if path == None:
        csv_path = os.path.join(resource_path, "LISTA CECO CASINOS.csv")

    json_path = os.path.join(resource_path, "contacts.json")
    contacts_dictionary = get_json(json_path)

    upload_list = []
    with open(csv_path, "r", encoding="ISO-8859-1") as csv_file:
        for line in csv_file:
            id = line.split(";")[0]
            upload_list.append(id)

    keys = list(contacts_dictionary.keys())
    upload_list_id = list(filter(lambda x: not x in upload_list, keys))

    with open(csv_path, "a+", encoding="ISO-8859-1") as csv_file:
        for id in upload_list_id:
            mc.mcprint(text="Inserting to .CSV ['{}': {}]".format(id, contacts_dictionary[id]), color=mc.Color.GREEN)
            csv_file.write(
                "{};{};{}\n".format(id, contacts_dictionary[id]["contacto"], contacts_dictionary[id]["rut_cliente"]))

    mc.mcprint("\nArchivo actualizado exitosamente", color=mc.Color.GREEN)


def show_shorteners():
    dictionary = get_json()
    print()
    for shortener_key in dictionary["shortener"]:
        mc.mcprint(text="{}: {}".format(shortener_key, dictionary["shortener"][shortener_key]), color=mc.Color.YELLOW)


def add_shortener():
    show_shorteners()
    dictionary = get_json()
    path = get_json()["configuration_path"]
    print()
    shortener_key = mc.get_input(text="String a reemplazar")
    shortener = mc.get_input(text="Nuevo string")
    dictionary["shortener"][shortener_key] = shortener
    mc.mcprint(text="\nAcortador '{}': '{}' ha sido creado exitosamente".format(shortener_key, shortener),
               color=mc.Color.GREEN)

    mc.generate_json(path=path, dictionary=dictionary)


def remove_shortener():
    show_shorteners()
    dictionary = get_json()
    path = get_json()["configuration_path"]

    shortener_keys = list(dictionary["shortener"].keys())
    selection_menu = mc.Menu(title="Seleccione acortador para eliminar", options=shortener_keys)
    selection_menu.show()
    if selection_menu.returned_value != "0":
        selected_key = shortener_keys[int(selection_menu.returned_value) - 1]
        del dictionary["shortener"][selected_key]

        mc.mcprint(dictionary, color=mc.Color.RED)
        mc.generate_json(path=path, dictionary=dictionary)
        logging.info("El acortador '{}' se ha eliminado".format(selected_key))


def manage_shorteners():
    mf_add_shortener = mc.MenuFunction("Agregar/Editar Acortador", add_shortener)
    mf_remove_shortener = mc.MenuFunction("Remover Acortador", remove_shortener)
    mf_show_shorteners = mc.MenuFunction("Mostrar Acortadores", show_shorteners)

    manage_shorteners_menu = mc.Menu(title="Gestionar Acortadores",
                                     options=[mf_add_shortener, mf_remove_shortener, mf_show_shorteners])
    manage_shorteners_menu.show()


def regenerate_configuration_file():
    configuration_path = get_json()["configuration_path"]
    os.remove(configuration_path)
    logging.info("file removed '{}'".format(configuration_path))
    initialize_config()


# Shortener Manager Menu
mf_add_shortener = mc.MenuFunction("Agregar/Editar Acortador", add_shortener)
mf_remove_shortener = mc.MenuFunction("Remover Acortador", remove_shortener)
mf_show_shorteners = mc.MenuFunction("Mostrar Acortadores", show_shorteners)

mc_manage_shorteners = mc.Menu(title="Gestionar Acortadores",
                               options=[mf_add_shortener, mf_remove_shortener, mf_show_shorteners])

# Contact Manager Menu
mc_add_contact = mc.MenuFunction("Agregar Nuevo Contacto", add_contact)
mc_show_contacts = mc.MenuFunction("Mostrar Contactos", show_contacts)
mc_import_contacts = mc.MenuFunction("Importar Contactos", import_contacts)
mc_edit_contacts = mc.MenuFunction("Editar Contactos", edit_contact)
mc_remove_contact = mc.MenuFunction("Eliminar Contacto", remove_contact)
mc_export_contacts = mc.MenuFunction(
    "{}Exportar Contactos (usar con precaucion){}".format(mc.Color.RED, mc.Color.RESET), export_contact)

mc_manage_contacts = mc.Menu(title="Gestionar Contactos",
                             options=[mc_add_contact, mc_edit_contacts, mc_show_contacts, mc_import_contacts,
                                      mc_export_contacts, mc_remove_contact])

# Regenerate configuration Menu
mf_regenerate_configuration_file = mc.MenuFunction("{}Restablecer{}".format(mc.Color.RED, mc.Color.RESET),
                                                    regenerate_configuration_file)
mc_regenerate_config = mc.Menu(title="{}Restablecer Configuracion{}".format(mc.Color.YELLOW, mc.Color.RESET),
                               subtitle="¿Seguro que desea restablecer a la configuracion inicial? "
                                        "{}(Se perderan todos los acortadores){}".format(mc.Color.RED, mc.Color.RESET),
                               options=[mf_regenerate_configuration_file])


