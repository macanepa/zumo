from datetime import datetime, timedelta
from selenium import webdriver
import sys, os, time, utilities
import mcutils as mc


def generate_web_driver():
    download_directory = utilities.get_json()["downloads_directory"]
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': '{}'.format(download_directory)}
    chrome_options.add_experimental_option('prefs', prefs)

    if getattr(sys, 'frozen', False):
        chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
        driver = webdriver.Chrome(options=chrome_options, executable_path=chromedriver_path)
    else:
        driver = webdriver.Chrome(options=chrome_options)

    return driver

def get_credentials():
    username = mc.get_input(text="usuario")
    organization = mc.get_input(text="organizacion")
    password = mc.get_input(text="clave")
    return username, organization, password

def input_field(driver, id, value):
    field = driver.find_element_by_id(id)
    field.click()
    field.send_keys(value)

def calculate_date():
    date = (datetime.now().date().replace(day=1) - timedelta(days=1))
    date = date.replace(day=1)
    return date.strftime("%d-%m-%Y")

def get_excel():

    # Get login credentials from user
    username, organization, password = get_credentials()

    # Head to website
    url = "https://sodexo.iconstruye.com/"
    driver = generate_web_driver()
    driver.get(url)

    # Login
    input_field(driver=driver, id="txtUsuario", value=username)
    input_field(driver=driver, id="txtEmpresa", value=organization)
    input_field(driver=driver, id="txtClave", value=password)

    login_button = driver.find_element_by_id("Image2")
    login_button.click()

    try:
        # Retrieve xls file
        url = "https://sodexo.iconstruye.com/Reportes/compra/producto_detallado_proveedor.aspx"
        driver.get(url)
        time.sleep(1) # wait a second to load start_date element
        start_date = driver.find_element_by_id("ctrRangoFechaDespachoFECHADESDE")
        driver.execute_script("arguments[0].setAttribute(arguments[1], arguments[2]);",
                              start_date,
                              "value",
                              calculate_date())
    except:
        mc.register_error(error_string="Credentials are invalid", print_error=True)
        driver.quit()
        return

    search_button = driver.find_element_by_id("btnBuscar")
    search_button.click()
    driver.switch_to.alert.accept()

    excel_download_button = driver.find_element_by_id("lnkExcel")
    excel_download_button.click()

    # use quit() to close all windows
    driver.close()

