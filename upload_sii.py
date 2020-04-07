import mcutils as mc
import utilities
import web_driver
import os, time
from selenium.webdriver.common.keys import Keys
from datetime import datetime

def get_credentials():
    username = mc.get_input(text="rut")
    password = mc.get_input(text="clave")
    return username, password

def login(driver, rut, password):

        url = "https://www1.sii.cl/cgi-bin/Portal001/mipeLaunchPage.cgi?OPCION=33&TIPO=4"
        driver.get(url)

        web_driver.input_field(driver=driver, id="rutcntr", value=rut)
        web_driver.input_field(driver=driver, id="clave", value=password)

        login_button = driver.find_element_by_id("bt_ingresar")
        login_button.click()

        url = "https://www1.sii.cl/cgi-bin/Portal001/mipeSelEmpresa.cgi?DESDE_DONDE_URL=OPCION%3D33%26TIPO%3D4"
        driver.get(url)

        # Select second option
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/form/div/div[1]/div/div/select/optgroup/option[2]").click()
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/form/div/div[2]/button").click()


def load_chunks():
    path = utilities.get_json()["output_directory"]
    path = os.path.join(path, "temp.json")
    dictionary = utilities.get_json(path)
    day_list = list(dictionary.keys())
    mc_select = mc.Menu(title="Seleccionar dia", options=day_list)
    mc_select.show()
    returned_value = mc_select.returned_value

    day_dictionary = {}
    if returned_value != 0:
        day_dictionary = dictionary["day_{}".format(int(returned_value) - 1)]
    return day_dictionary

def load_sii(chunk_dictionary, driver):
    set_page(driver)
    # mc.mcprint("\n{}".format(chunk_dictionary), color=mc.Color.CYAN)

    for key in chunk_dictionary:
        rut_cliente = chunk_dictionary[key]["rut_cliente"].split("-")
        folio = "OC{}".format(chunk_dictionary[key]["n_order_compra"])
        contacto = chunk_dictionary[key]["contacto"]
        s_fecha = chunk_dictionary[key]["fecha_emision"]
        d_fecha = datetime.strptime(s_fecha, "%d-%m-%Y")

    for key, i in zip(chunk_dictionary, range(1, len(chunk_dictionary)+1)):
        product = chunk_dictionary[key]
        # mc.mcprint(product, color=mc.Color.YELLOW)

        list = []
        list.append(driver.find_element_by_name("EFXP_NMB_{:0>2d}".format(i))) # name
        list.append(driver.find_element_by_name("EFXP_QTY_{:0>2d}".format(i))) # cant
        list.append(driver.find_element_by_name("EFXP_UNMD_{:0>2d}".format(i))) # unit
        list.append(driver.find_element_by_name("EFXP_PRC_{:0>2d}".format(i))) # price

        product_data = [product["descripcion"], product["cantidad"], product["unidad"], product["precio_unitario"]]
        for element, element_index in zip(list, range(len(list))):
            element.click()
            element.send_keys(str(product_data[element_index]))



    rut_field = driver.find_element_by_id("EFXP_RUT_RECEP")
    rut_field.click()
    rut_field.send_keys(rut_cliente[0])
    rut_field.send_keys(Keys.TAB)
    rut_field = driver.find_element_by_id("EFXP_DV_RECEP")
    rut_field.click()
    rut_field.send_keys(rut_cliente[1])
    rut_field.send_keys(Keys.TAB)



    xpath = "/html/body/div[1]/div[2]/div[1]/div/form/div/div[4]/div[2]/div/div[4]/div/div/select/optgroup/option[8]"
    giro = driver.find_element_by_xpath(xpath)
    giro.click()

    # introducir contacto
    xpath = "/html/body/div[1]/div[2]/div[1]/div/form/div/div[4]/div[2]/div/div[5]/div[1]/div/input"
    contacto_field = driver.find_element_by_xpath(xpath)
    contacto_field.click()
    contacto_field.send_keys(contacto)


    # click en referencias
    xpath = "/html/body/div[1]/div[2]/div[1]/div/form/div/div[7]/table[2]/tbody/tr[1]/th[1]/input"
    referencias_button = driver.find_element_by_xpath(xpath)
    referencias_button.click()

    # click en orden de compra
    xpath = "/html/body/div[1]/div[2]/div[1]/div/form/div/div[7]/table[2]/tbody/tr[3]/td[1]/select/optgroup/option[22]"
    oc_button = driver.find_element_by_xpath(xpath)
    oc_button.click()

    # folio

    xpath = "/html/body/div[1]/div[2]/div[1]/div/form/div/div[7]/table[2]/tbody/tr[3]/td[3]/input"
    folio_field = driver.find_element_by_xpath(xpath)
    folio_field.click()
    folio_field.send_keys(folio)

    # set date
    dia_field = driver.find_element_by_name("cbo_dia_boleta_ref_01")
    mes_field = driver.find_element_by_name("cbo_mes_boleta_ref_01")
    anio_field = driver.find_element_by_name("cbo_anio_boleta_ref_01")

    dia_field = dia_field.find_elements_by_xpath("*")[0]
    dia_field = dia_field.find_elements_by_xpath("*")[d_fecha.day - 1]
    dia_field.click()

    mes_field = mes_field.find_elements_by_xpath("*")[0]
    mes_field = mes_field.find_elements_by_xpath("*")[d_fecha.month - 1]
    mes_field.click()

    anio_field = anio_field.find_elements_by_xpath("*")[0]
    anio_field = anio_field.find_elements_by_xpath("*")[2]
    anio_field.click()

    xpath = "/html/body/div[1]/div[2]/div[1]/div/form/p/button[4]"
    borrador_button = driver.find_element_by_xpath(xpath)
    borrador_button.click()

    time.sleep(.5)


def set_page(driver):
    url = "https://www1.sii.cl/cgi-bin/Portal001/mipeLaunchPage.cgi?OPCION=33&TIPO=4"
    driver.get(url)
    time.sleep(2)

    for i in range(9):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(.05)
        selection = driver.find_element_by_id("rowDet_Botones")
        selection = selection.find_elements_by_xpath("*")
        add_button = selection[0]
        add_button.click()

def input_credentials():
    while True:
        try:
            rut, password = get_credentials()
            driver = web_driver.generate_web_driver()
            login(driver=driver, rut=rut, password=password)
            break
        except:
            mc.register_error(error_string="Credentials are invalid", print_error=True)
            mc.mcprint(text="Vuelva a introducir las credenciales", color=mc.Color.YELLOW)
            driver.quit()
    return driver

def begin():

    driver = input_credentials()

    day_dictionary = load_chunks()
    for order, index in zip(day_dictionary, range(len(day_dictionary))):
        for chunk in day_dictionary[order]:
            while True:
                try:
                    mc.progress_bar(current_index=index, total=range(len(day_dictionary)))
                    load_sii(chunk_dictionary=day_dictionary[order][chunk], driver=driver)
                    break
                except:
                    mc.clear(10)
                    mc.register_error(error_string="Ocurrio un error en la subida de la orden [{}]".format(order),
                                      print_error=True)
                    mc.mcprint(text="Reintentando en 5 segundos", color=mc.Color.YELLOW)
                    time.sleep(5)
