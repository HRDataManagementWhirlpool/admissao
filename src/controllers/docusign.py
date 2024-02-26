from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

import time
import os
from dotenv import load_dotenv
load_dotenv()

class DocusignController:
    def __init__(self, email, senha):
        self.SITE_LINK = {
            "home": "https://apps.docusign.com/send/documents",
            "templates": "https://apps.docusign.com/send/templates"
        }
        self.SITE_MAP = {
            "buttons": {
                "login_btn":{
                    "xpath": "//button[@type='submit']",
                },
                "username_btn": {
                    "xpath": "//button[@data-qa='submit-username']",
                    "name": "submit-username"
                },
                "password_btn": {
                    "xpath": "//button[@data-qa='submit-password']",
                },
                "template_btn": {
                    "xpath": "//button[@data-qa='header-TEMPLATES-tab-button']"
                },
                "use_template_btn": {
                    "xpath": "//button[@data-qa='templates-main-list-row-b4024bba-b898-4c4d-a1c5-3ef9b408ee3c-actions-use']"
                },
                "document_next_btn": {
                    "xpath": "//button[@data-qa='footer-add-fields-link']"
                },
                "zoom_btn": {
                    "xpath": "//button[@data-qa='zoom-button']",
                    "zoom_50": "//button[@data-qa='zoom-level-50']"
                },
                "rubrica_btn": {
                    "xpath": "//button[@data-qa='Initial']"
                },
                "assinatura_btn": {
                    "xpath": "//button[@data-qa='Signature']"
                },
                "pages_to_sign_btn": {
                    "xpath": "//button[@data-qa='tagger-documents']"
                },
                "popup_btn": {
                    "xpath": "//button[@data-qa='modal-cancel-btn']"
                }
            },
            "inputs": {
                "username": {
                    "xpath": "//input[@data-qa='username']",
                    "keys": email
                },
                "password": {
                    "xpath": "//input[@data-qa='password']",
                    "keys": senha
                },
                "template_searchbar": {
                    "xpath": "//input[@data-qa='templates-main-header-form-input']",
                    "keys": "Modelo de Contrato - CC"
                },
                "document_send": {
                    "id": "windows-drag-handler-wrapper",
                },
                "document_subject": {
                    "xpath": "//input[@data-qa='prepare-subject']",
                    "keys": "Contrato Teste"
                },
                "document_message": {
                    "xpath": "//textarea[@data-qa='prepare-message']",
                    "keys": "Mensagem Teste"
                },
                "pages_area": {
                    "xpath": "//div[@data-qa='document-accordion-region']"
                },
                "pages_signed": {
                    "xpath": "//div[@data-qa='indicator-tag']"
                },
            }
        }
        
        self.options = webdriver.ChromeOptions()
        self.options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\Test Data")
        self.options.add_argument(r'--profile-directory=Default')
        
        self.driver = webdriver.Chrome(options=self.options)
        self.driver.maximize_window()
        
        self.wait = WebDriverWait(self.driver, 30)
    
    def open_website(self):
        self.driver.get(self.SITE_LINK['home'])
        
    def login(self):
        element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, self.SITE_MAP['buttons']['login_btn']['xpath']))
        )
        if element.get_dom_attribute("data-qa")  == self.SITE_MAP['buttons']['username_btn']['name']:
            element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, self.SITE_MAP['inputs']['username']['xpath']))
            )
            element.send_keys(Keys.CONTROL+"a")
            element.send_keys(Keys.DELETE)
            element.send_keys(self.SITE_MAP['inputs']['username']['keys'])
            element = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['username_btn']['xpath']))
            )
            element.click()
        else:
            print(self.SITE_MAP['inputs']['username']['keys'])
            element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, self.SITE_MAP['inputs']['password']['xpath']))
            )
            element.send_keys(Keys.CONTROL+"a")
            element.send_keys(Keys.DELETE)
            element.send_keys(self.SITE_MAP['inputs']['password']['keys'])
            element = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['password_btn']['xpath']))
            ).click()

    def select_template(self):
        self.driver.get(self.SITE_LINK['templates'])
        element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, self.SITE_MAP['inputs']['template_searchbar']['xpath']))
        )
        element.send_keys(self.SITE_MAP['inputs']['template_searchbar']['keys'])
        element.send_keys(Keys.ENTER)
        element = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['use_template_btn']['xpath']))
        ).click()
        
    def select_document_to_sign(self, document):
        element = self.wait.until(
            EC.presence_of_element_located((By.ID, self.SITE_MAP['inputs']['document_send']['id']))
        )
        element.find_element(By.TAG_NAME, "input").send_keys(document)
        
        element = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['inputs']['document_subject']['xpath']))
        )
        element.send_keys(Keys.CONTROL+"a")
        element.send_keys(Keys.DELETE)
        
        element.send_keys(self.SITE_MAP['inputs']['document_subject']['keys'])
        element = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['inputs']['document_message']['xpath']))
        )
        element.send_keys(Keys.CONTROL+"a")
        element.send_keys(Keys.DELETE)
        element.send_keys(self.SITE_MAP['inputs']['document_message']['keys'])
        
        element = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['document_next_btn']['xpath']))
        ).click()

    def sign_document_select_zoom(self):
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['zoom_btn']['xpath']))
        ).click()
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['zoom_btn']['zoom_50'])) ## Zoom 50 para rodar em monitores e no notebook sem quebrar
        ).click()

    def sign_pages(self):
        element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, self.SITE_MAP['inputs']['pages_area']['xpath']))
        )
        paginas = element.find_elements(By.XPATH, self.SITE_MAP['buttons']['pages_to_sign_btn']['xpath'])
        rubrica = self.driver.find_element(By.XPATH, self.SITE_MAP['buttons']['rubrica_btn']['xpath'])
        assinatura = self.driver.find_element(By.XPATH, self.SITE_MAP['buttons']['assinatura_btn']['xpath'])
        
        count = 0
        while len(paginas) != count:
            count = len(self.driver.find_elements(By.XPATH, self.SITE_MAP['inputs']['pages_signed']['xpath']))
            try:
                for i, pagina in enumerate(paginas):
                    element = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['pages_to_sign_btn']['xpath']))
                    )
                    alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
                    if "tem campos nesta página" not in alt_value:
                        pagina.click()
                        if i % 2 == 0 or i == 0:
                            ActionChains(self.driver)\
                                .click_and_hold(rubrica)\
                                .move_by_offset(332, 200)\
                                .release()\
                                .perform()
                        else:
                            ActionChains(self.driver)\
                                .click_and_hold(assinatura)\
                                .move_by_offset(265, 180)\
                                .release()\
                                .perform()
                        rubrica.click()
            except:
                time.sleep(1)
                element = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()
                time.sleep(1)
                rubrica.click()
        time.sleep(2)
        self.driver.quit()