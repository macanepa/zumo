import mcutils as mc
import web_driver, excel_work
import webbrowser, os, json
dev = 0

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def load_config():


    directory = os.getcwd() + "\\zumo.json"
    if not os.path.isfile(directory):
        print("Creating File: {}".format(directory))
        config = {
            "save_directory": "{}\\PRECIOS COMPRA VENTA.xlsx".format(os.getcwd())
        }

        with open("zumo.json", "w") as config_file:
            json.dump(config, config_file)
    with open("zumo.json", "r") as config_file:
        data = json.load(config_file)
        print("loaded: {}".format(config_file.name))



def get_update(url):
    webbrowser.open(url)


mf_start = mc.Menu_Function("Obtener Excel", web_driver.get_excel)
mf_excel_work = mc.Menu_Function("Formatear Excel", excel_work.begin)
mf_update_zumo = mc.Menu_Function("Descargar Actualizacion", get_update, "https://github.com/macanepa/zumo/releases")
mf_exit = mc.Menu_Function("Salir", mc.exit_application)
m_main_menu = mc.Menu(title="Menu Principal", options=[mf_start, mf_excel_work, mf_update_zumo, mf_exit], back=False)

path = os.getcwd()+"\\data"
if not os.path.isdir(path):
    os.mkdir("data")

load_config()

if(dev==0):
    while True:
        m_main_menu.show()
else:
    web_driver.get_excel()
    input("Presione cualquier tecla para continuar...")