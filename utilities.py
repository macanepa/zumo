import os, json, webbrowser, sys

def initialize_config():

    output_directory = os.path.join(os.getcwd(),"output")
    config_directory = os.path.join(os.getcwd())
    resources_directory = os.path.join(os.getcwd(),"precios compra venta")
    downloads_directory = os.path.join(os.getcwd(),"downloads")
    directories = [output_directory, config_directory, resources_directory, downloads_directory]

    config_file_directory = os.path.join(config_directory, "zumo.json")


    if not os.path.isfile(config_file_directory):
        print("Creating File: {}".format(config_directory))
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
            print("Creating folder: {}".format(dir))
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
    with open(config_directory, "r") as json_file:
        dictionary = json.load(json_file)
    return dictionary