import mcutils as mc
import web_driver, excel_work
dev = 0

mf_start = mc.Menu_Function("Obtener Excel", web_driver.get_excel)
mf_excel_work = mc.Menu_Function("Formatear Excel", excel_work.begin)
mf_exit = mc.Menu_Function("Salir", mc.exit_application)
m_main_menu = mc.Menu(title="Menu Principal", options=[mf_start, mf_excel_work, mf_exit], back=False)


if(dev==0):
    while True:
        m_main_menu.show()
else:
    web_driver.get_excel()
    input("Presione cualquier tecla para continuar...")