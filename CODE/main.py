import mcutils as mc
import utilities
import xl_generation
import web_driver
import upload_sii
import remove_drafts
import generate_monthly_report
import os

mc.activate_mc_logger(console_log_level='info')

folder_name = os.path.basename(os.getcwd())

# mcutils logger
mc.ColorSettings.is_dev = False

# mcutils Menu declarations
mf_generate_monthly_report = mc.MenuFunction("Generar Reporte Mensual", generate_monthly_report.begin)
m_reports = mc.Menu(title="Generar Reporte", options=[mf_generate_monthly_report])

# Utilities Menu
mf_remove_drafts = mc.MenuFunction("Remover Borradores", remove_drafts.begin)
mf_download_update = mc.MenuFunction("Descargar Actualizaciones", utilities.download_update)
mf_regenerate_configuration_file = mc.MenuFunction("{}Restablecer configuracion{}".format(mc.Color.YELLOW,
                                                                                           mc.Color.RESET),
                                                    utilities.regenerate_configuration_file)
m_utilities = mc.Menu(title="Herramientas", options=[utilities.mc_manage_contacts, utilities.mc_manage_shorteners,
                                                     mf_remove_drafts, mf_download_update,
                                                     utilities.mc_regenerate_config])

mf_start = mc.MenuFunction("Obtener Excel", web_driver.get_excel)
mf_generate_xl = mc.MenuFunction("Formatear Excel", xl_generation.begin)
mf_upload_sii = mc.MenuFunction("Generar Factura SII", upload_sii.begin)
mf_exit = mc.MenuFunction(title="Salir", function=mc.exit_application, text="Exiting Application")
m_main_menu = mc.Menu(title="MENU PRINCIPAL", subtitle=" Seleccione una de las siguientes opciones\n",
                      options=[mf_start, mf_generate_xl, mf_upload_sii, m_reports, m_utilities, mf_exit], back=False)

# Initialize software configurations
utilities.initialize_config()

while True:
    m_main_menu.show()
