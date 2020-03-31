import mcutils as mc
import utilities
import xl_generation
import web_driver
import upload_sii

# mcutils logger
mc.Log_Settings.display_logs = True
mc.Color_Settings.is_dev = False

# mcutils Menu declarations
mf_start = mc.Menu_Function("Obtener Excel", web_driver.get_excel)
mf_generate_xl = mc.Menu_Function("Formatear Excel", xl_generation.begin)
mf_upload_sii = mc.Menu_Function("Generar Factura SII", upload_sii.begin)
mf_manage_contacts = mc.Menu_Function("Gestionar Contactos", utilities.edit_contacts)
mf_download_update = mc.Menu_Function("Descargar Actualizacion", utilities.download_update, "https://github.com/macanepa/zumo/releases")
mf_exit = mc.Menu_Function("Salir", mc.exit_application, *["Exiting Application"])
m_main_menu = mc.Menu(title="MENU PRINCIPAL", subtitle=" Seleccione una de las siguientes opciones\n", options=[mf_start, mf_generate_xl, mf_upload_sii, mf_manage_contacts, mf_download_update, mf_exit], back=False)

# Initialize software configurations
utilities.initialize_config()

while True:
    m_main_menu.show()
