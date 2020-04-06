import mcutils as mc
import utilities
import xl_generation
import web_driver
import upload_sii
import remove_drafts
import generate_monthly_report

# mcutils logger
mc.Log_Settings.display_logs = True
mc.Color_Settings.is_dev = True

# mcutils Menu declarations
mf_generate_monthly_report = mc.Menu_Function("Generar Reporte Mensual", generate_monthly_report.begin)
m_reports = mc.Menu(title="Generar Reporte", options=[mf_generate_monthly_report])

mf_remove_drafts = mc.Menu_Function("Remover Borradores", remove_drafts.begin)
mf_manage_contacts = mc.Menu_Function("Gestionar Contactos", utilities.edit_contacts)
m_utilities = mc.Menu(title="Herramientas", options=[mf_manage_contacts, mf_remove_drafts])

mf_start = mc.Menu_Function("Obtener Excel", web_driver.get_excel)
mf_generate_xl = mc.Menu_Function("Formatear Excel", xl_generation.begin)
mf_upload_sii = mc.Menu_Function("Generar Factura SII", upload_sii.begin)
mf_download_update = mc.Menu_Function("Descargar Actualizacion", utilities.download_update, "https://github.com/macanepa/zumo/releases")
mf_exit = mc.Menu_Function("Salir", mc.exit_application, *["Exiting Application"])
m_main_menu = mc.Menu(title="MENU PRINCIPAL", subtitle=" Seleccione una de las siguientes opciones\n", options=[mf_start, mf_generate_xl, mf_upload_sii, m_reports, m_utilities, mf_download_update, mf_exit], back=False)

# Initialize software configurations
utilities.initialize_config()

while True:
    m_main_menu.show()
