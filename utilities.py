#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os, json, webbrowser, sys
import mcutils as mc

def initialize_config():

    ul = mc.Log_Manager(developer_mode=True)

    output_directory = os.path.join(os.getcwd(),"output")
    config_directory = os.path.join(os.getcwd())
    resources_directory = os.path.join(os.getcwd(),"resources")
    downloads_directory = os.path.join(os.getcwd(),"downloads")
    directories = [output_directory, config_directory, resources_directory, downloads_directory]

    config_file_directory = os.path.join(config_directory, "zumo.json")


    if not os.path.isfile(config_file_directory):
        ul.log(text="Creating File: {}".format(config_file_directory))
        config = {
            "precios_compra_venta_file": "{}".format(os.path.join(resources_directory,"PRECIOS COMPRA VENTA.xlsx")),
            "output_directory": "{}".format(output_directory),
            "resources_directory": "{}".format(resources_directory),
            "downloads_directory": "{}".format(downloads_directory)
        }

        with open("zumo.json", "w") as config_file:
            json.dump(config, config_file)

    # create required folders
    for dir in directories:
        if not os.path.isdir(dir):
            text = "Creating folder: {}".format(dir)
            ul.log(text=text)
            os.mkdir(dir)

def download_update(url):
    webbrowser.open(url)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def get_json(config_directory = os.path.join(os.getcwd(),"zumo.json")):
    try:
        return mc.get_dict_from_json(config_directory)
    except FileNotFoundError as e:
        text = "No se ha encontrado el archivo {}".format(e.filename)
        mc.register_error(error_string=text)
        mc.exit_application(enter_quit=True)

def import_contacts(path=None):

    if path == None:
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

def edit_contacts():

    mc_add_contact = mc.Menu_Function("Agregar Nuevo Contacto", add_contact)
    mc_show_contacts = mc.Menu_Function("Mostrar Contactos", show_contacts)
    mc_import_contacts = mc.Menu_Function("Importar Contactos", import_contacts)
    mc_contacts = mc.Menu(title="Gestionar Contactos", options=[mc_add_contact, mc_show_contacts, mc_import_contacts])

    mc_contacts.show()

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
        mc.mcprint("{}\t{}".format(key,contact_dictionary[key]), color=mc.Color.YELLOW)

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