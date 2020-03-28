from selenium import webdriver
import requests, sys, os, time
import mcutils as mc
from datetime import datetime, timedelta


def connect_website(url):

    cwd = os.getcwd()+"\\data"
    print("CWD:", cwd)
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': '{}'.format(cwd)}
    chrome_options.add_experimental_option('prefs', prefs)
    # chrome_options.add_argument("headless")
    options = chrome_options

    if getattr(sys, 'frozen', False):
        chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
        driver = webdriver.Chrome(options=options, executable_path=chromedriver_path)
    else:
        driver = webdriver.Chrome(options=options)

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

def date_calculate():

    date = datetime.now().date().replace(day=1) - timedelta(days=1)
    date = date.replace(day=1)
    return date.strftime("%d-%m-%Y")

def clear_folder():
    os.chdir('data')
    #Clear folder
    for item in os.listdir():
        os.remove(item)
    os.chdir('..')

def get_excel():

    #Login to sodexo
    username, organization, password = get_credentials()

    print("Login Screen")

    url = "https://sodexo.iconstruye.com/"
    driver = connect_website(url)
    driver.get(url)

    input_field(driver=driver, id="txtUsuario", value=username)
    input_field(driver=driver, id="txtEmpresa", value=organization)
    input_field(driver=driver, id="txtClave", value=password)

    login_button = driver.find_element_by_id("Image2")
    login_button.click()



    #Get the excel file
    url = "https://sodexo.iconstruye.com/Reportes/compra/producto_detallado_proveedor.aspx"
    driver.get(url)

    time.sleep(1)

    start_date = driver.find_element_by_id("ctrRangoFechaDespachoFECHADESDE")
    driver.execute_script("arguments[0].setAttribute(arguments[1], arguments[2]);",
                          start_date,
                          "value",
                          date_calculate())

    search_button = driver.find_element_by_id("btnBuscar")
    search_button.click()

    print("Login Successful")

    driver.switch_to.alert.accept()

    clear_folder()

    excel_download_button = driver.find_element_by_id("lnkExcel")
    excel_download_button.click()


    driver.close()

