import mcutils as mc
import utilities
import web_driver, upload_sii
import os, time
import tqdm

def remove_draft(driver, url):
    sleep_time = 0.5
    driver.get(url)
    time.sleep(sleep_time)
    delete_button = driver.find_element_by_name("Button_Delete_Borrador")
    delete_button.click()
    time.sleep(sleep_time)
    driver.switch_to.alert.accept()


def get_draft_urls(driver):
    url = "https://www4.sii.cl/mipymeinternetui/#!/borradores"
    driver.get(url)

    time.sleep(5)
    xpath = "/html/body/div[1]/div[2]/div/div/div/div/div/div[2]/div[1]/div[1]/div/label/select/option[4]"
    documents_per_page_list = driver.find_element_by_xpath(xpath)
    documents_per_page_list.click()

    time.sleep(5)
    xpath = "/html/body/div[1]/div[2]/div/div/div/div/div/div[2]/div[2]/div/table/tbody"
    table = driver.find_element_by_xpath(xpath)
    rows = table.find_elements_by_xpath("*")

    draft_urls = []
    for element in rows:
        first_element = element.find_elements_by_xpath("*")[0]
        first_element = first_element.find_elements_by_xpath("*")[0]
        url = first_element.get_attribute("href")
        draft_urls.append(url)

    return draft_urls

def begin():

    driver = upload_sii.input_credentials()
    if driver:
        draft_urls = get_draft_urls(driver)
        for draft_url in tqdm.tqdm(draft_urls):
            remove_draft(driver, draft_url)
        mc.mcprint(text="Se han eliminado todos los borradores", color=mc.Color.YELLOW)