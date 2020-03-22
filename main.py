import mcutils as mc
import web_driver
dev = 0

mf_start = mc.Menu_Function("Obtener Excel", web_driver.get_excel)
mf_exit = mc.Menu_Function("Salir", mc.exit_application)
m_main_menu = mc.Menu(title="Menu Principal", options=[mf_start, mf_exit], back=False)


if(dev==0):
    while True:
        m_main_menu.show()
else:
    web_driver.get_excel()
    input("Presione cualquier tecla para continuar...")