from datetime import datetime, timedelta
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import sys, os, time
import utilities
import mcutils as mc
import logging


def generate_web_driver(headless: bool = False):
    download_directory = utilities.get_json()["downloads_directory"]
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': '{}'.format(download_directory)}
    chrome_options.add_experimental_option('prefs', prefs)

    #chrome_options.add_argument('--remote-debugging-port=0')
    chrome_options.add_argument('--remote-debugging-port=9222')
    chrome_options.add_argument('--remote-debugging-host=0.0.0.0')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--no-sandbox')

    # chrome_options.add_argument("--window-size=240,320")
    if headless:
        chrome_options.add_argument('--headless')
    # chrome_options.add_argument('--disable-gpu')

    driver = webdriver.Chrome(options=chrome_options, executable_path=ChromeDriverManager().install())

    # driver.set_window_size(480, 320)
    return driver

def get_credentials():
    mc_credentials = mc.Menu(title='Introduzca sus credenciales de SODEXO',
                             subtitle='Introuzca 0 en cualquier campo para volver',
                             options={'usuario': [str, '>0'],
                                      'organizacion': [str, '>0'],
                                      'clave': [str, '>0'],
                                      }, input_each=True)
    mc_credentials.show()
    username = mc_credentials.returned_value['usuario']
    organization = mc_credentials.returned_value['organizacion']
    password = mc_credentials.returned_value['clave']
    return username, organization, password


def input_field(driver, id, value):
    field = driver.find_element_by_id(id)
    field.click()
    field.send_keys(value)


def calculate_date():
    date = (datetime.now().date().replace(day=1) - timedelta(days=1))
    date = date.replace(day=1)
    return date.strftime("%d-%m-%Y")


def wait_for_download(original_count: int):
    download_directory = utilities.get_json()["downloads_directory"]
    timeout = 5
    while timeout > 0:
        time.sleep(0.5)
        if original_count != len(os.listdir(download_directory)):
            logging.info('Download completed')
            return
        timeout -= 1
        time.sleep(4.5)
    logging.error('Timeout error. excel file didn\t download successfully')


def get_excel(fecha_desde=None):

    # Get login credentials from user
    username, organization, password = get_credentials()
    if '0' in [username, organization, password]:
        return

    # Head to website
    url = "https://sodexo.iconstruye.com/"
    driver = generate_web_driver(headless=True)
    driver.get(url)
    # Login
    input_field(driver=driver, id="txtUsuario", value=username)
    input_field(driver=driver, id="txtEmpresa", value=organization)
    input_field(driver=driver, id="txtClave", value=password)

    login_button = driver.find_element_by_id("Image2")
    login_button.click()

    date = calculate_date()
    if fecha_desde != None:
        date = fecha_desde


    try:
        # Retrieve xls file
        url = "https://sodexo.iconstruye.com/Reportes/compra/producto_detallado_proveedor.aspx"
        driver.get(url)
        time.sleep(1) # wait a second to load start_date element
        start_date = driver.find_element_by_id("ctrRangoFechaDespachoFECHADESDE")
        driver.execute_script("arguments[0].setAttribute(arguments[1], arguments[2]);",
                              start_date,
                              "value",
                              date)
    except:
        logging.error("Credentials are invalid")
        driver.quit()
        return

    search_button = driver.find_element_by_id("btnBuscar")
    search_button.click()

    try:
        driver.switch_to.alert.accept()
    except Exception:
        logging.warning('Alert element was not found')


    excel_download_button = driver.find_element_by_id("lnkExcel")
    
    download_directory = utilities.get_json()["downloads_directory"]
    original_count = len(os.listdir(download_directory))
    excel_download_button.click()
    wait_for_download(original_count)
    driver.quit()

