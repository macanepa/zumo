import mcutils as mc
import web_driver, excel_work
import webbrowser, os
dev = 0


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


if(dev==0):
    while True:
        m_main_menu.show()
else:
    web_driver.get_excel()
    input("Presione cualquier tecla para continuar...")