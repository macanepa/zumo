import mcutils as mc
import xl_generation, web_driver, utilities

# mcutils logger
mc.Log_Settings.display_logs = True
mc.Color_Settings.is_dev = True

# mcutils Menu declarations
mf_start = mc.Menu_Function("Obtener Excel", web_driver.get_excel)
mf_generate_xl = mc.Menu_Function("Formatear Excel", xl_generation.begin)
mf_download_update = mc.Menu_Function("Descargar Actualizacion", utilities.download_update, "https://github.com/macanepa/zumo/releases")
mf_exit = mc.Menu_Function("Salir", mc.exit_application, *["Exiting Application"])
m_main_menu = mc.Menu(title="MENU PRINCIPAL", subtitle=" Seleccione una de las siguientes opciones\n" , options=[mf_start, mf_generate_xl, mf_download_update, mf_exit], back=False)

# Initialize software configurations
utilities.initialize_config()

while True:
    m_main_menu.show()
