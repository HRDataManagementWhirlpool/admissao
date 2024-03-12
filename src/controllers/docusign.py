from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
import time
import os

class DocusignController:
    def __init__(self, email, senha):
        self.SITE_LINK = {
            "home": "https://apps.docusign.com/send/documents",
            "templates": "https://apps.docusign.com/send/templates",            
            "done": "https://apps.docusign.com/send/documents?label=completed"
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

    def open_website(self, tag='home'):
        self.driver.get(self.SITE_LINK[tag])

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

    def select_template_mensalista(self):
        self.driver.get('https://apps.docusign.com/send/templates?view=favorites')
        element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, self.SITE_MAP['inputs']['template_searchbar']['xpath']))
        )
        element.send_keys('CONTRATO DE TRABALHO -')
        element.send_keys(Keys.ENTER)
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span/span/button"))
        ).click()
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='quick-send-advanced-edit']"))
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
        while True:
            try:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['zoom_btn']['xpath']))
                ).click()
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['zoom_btn']['zoom_50'])) ## Zoom 50 para rodar em monitores e no notebook sem quebrar
                ).click()
                break
            except:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()

    def select_document_to_sign_mensalista(self, document, name:str, mail:str, title:str):
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='upload-file-button']"))
        ).click()
        self.wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[@data-qa='upload-file-input']"))
        ).send_keys(document)
        
        nome = self.driver.find_elements(By.XPATH, "//input[@data-qa='recipient-name']")
        email = self.driver.find_elements(By.XPATH, "//input[@data-qa='recipient-email']")
        nome[0].send_keys(name)
        email[0].send_keys(mail)
        nome[4].send_keys(name)
        email[4].send_keys(mail)
        
        subject = self.driver.find_element(By.XPATH, "//input[@data-qa='prepare-subject']")
        subject.send_keys(Keys.CONTROL+'a')
        subject.send_keys(title)
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['document_next_btn']['xpath']))
        ).click()

    def sign_pages_mensalista(self, nome):
        documentos = self.wait.until(
            EC.presence_of_all_elements_located((By.XPATH, "//div[@data-qa='doc-thumbnail-list']/button"))
        )
        while True:
            try:
                documentos[0].click()
                documentos[1].click()
                break
            except:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()
        
        paginas = self.driver.find_elements(By.XPATH, self.SITE_MAP['buttons']['pages_to_sign_btn']['xpath'])
        rubrica = self.driver.find_element(By.XPATH, self.SITE_MAP['buttons']['rubrica_btn']['xpath'])
        assinatura = self.driver.find_element(By.XPATH, self.SITE_MAP['buttons']['assinatura_btn']['xpath'])
        
        count = 0
        while (len(paginas)*2) != (count):
            try:
                for i, pagina in enumerate(paginas):
                    self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['pages_to_sign_btn']['xpath']))
                    )
                    alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
                    if "têm campos nesta página" not in alt_value:
                        pagina.click()
                        if i == 0 :
                            if (", "+nome) not in alt_value:
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-selector']"))
                                ).click()
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-1']"))
                                ).click()
                                ActionChains(self.driver)\
                                    .click_and_hold(rubrica)\
                                    .move_by_offset(200, 170)\
                                    .release()\
                                    .perform()
                            if "Claudinei Luis dos Santos" not in alt_value:
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-selector']"))
                                ).click()
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-2']"))
                                ).click()
                                ActionChains(self.driver)\
                                    .click_and_hold(rubrica)\
                                    .move_by_offset(320, 170)\
                                    .release()\
                                    .perform()
                        elif i == 1 :
                            if (", "+nome) not in alt_value:
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-selector']"))
                                ).click()
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-1']"))
                                ).click()
                                ActionChains(self.driver)\
                                    .click_and_hold(rubrica)\
                                    .move_by_offset(210, 155)\
                                    .release()\
                                    .perform()
                            if "Claudinei Luis dos Santos" not in alt_value:
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-selector']"))
                                ).click()
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-2']"))
                                ).click()
                                ActionChains(self.driver)\
                                    .click_and_hold(rubrica)\
                                    .move_by_offset(320, 155)\
                                    .release()\
                                    .perform()
                        else:
                            if (", "+nome) not in alt_value:
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-selector']"))
                                ).click()
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-1']"))
                                ).click()
                                ActionChains(self.driver)\
                                    .click_and_hold(assinatura)\
                                    .move_by_offset(132, 142)\
                                    .release()\
                                    .perform()
                            if "Claudinei Luis dos Santos" not in alt_value:
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-selector']"))
                                ).click()
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-2']"))
                                ).click()
                                ActionChains(self.driver)\
                                    .click_and_hold(assinatura)\
                                    .move_by_offset(258, 142)\
                                    .release()\
                                    .perform()
                        rubrica.click()
                    count = len(self.driver.find_elements(By.XPATH, self.SITE_MAP['inputs']['pages_signed']['xpath']))
                    alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
            except ElementClickInterceptedException:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()
                time.sleep(1)
                rubrica.click()
            except Exception as e:
                return e
        #self.driver.find_element(By.XPATH, "//button[@data-qa='footer-send-button']").click()
        time.sleep(2)

    def sign_pages_horista(self):
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
        #self.driver.find_element(By.XPATH, "//button[@data-qa='footer-send-button']").click()
        self.driver.quit()

    def sign_pages_aprendiz(self):
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
                        if (i+1) % 3 == 0:
                            ActionChains(self.driver)\
                                .click_and_hold(assinatura)\
                                .move_by_offset(260, 225)\
                                .release()\
                                .perform()
                        else:
                            ActionChains(self.driver)\
                                .click_and_hold(rubrica)\
                                .move_by_offset(330, 195)\
                                .release()\
                                .perform()
                        rubrica.click()
            except:
                try:
                    element = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                    ).click()
                except:
                    pass
                time.sleep(1)
                rubrica.click()
        #self.driver.find_element(By.XPATH, "//button[@data-qa='footer-send-button']").click()

    def horista_send_process(self, file):
        self.open_website()
        self.login()
        self.select_template()
        self.select_document_to_sign(file)
        self.sign_document_select_zoom()
        self.sign_pages_horista()

    def mensalista_send_process(self, mensalistas:dict):
        folder = os.listdir('files')
        self.open_website()
        self.login()
        for mensalista in mensalistas:
            for file in folder:
                if mensalista['re'] in file:
                    self.select_template_mensalista()
                    self.select_document_to_sign_mensalista(os.path.abspath(os.path.join('files', file)), mensalista['nome'], mensalista['email'], file[:-4])
                    self.sign_document_select_zoom()
                    self.sign_pages_mensalista(mensalista['nome'])

    def aprendiz_send_process(self, file):
        self.open_website()
        self.login()
        self.select_template()
        self.select_document_to_sign(file)
        self.sign_document_select_zoom()
        self.sign_pages_aprendiz()

    def get_signed_document(self, doc_name):
        self.open_website('done')
        
        element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[@data-qa='manage-envelopes-header-form-input']"))
        )
        element.send_keys(doc_name)
        element.send_keys(Keys.ENTER)
        
        try:
            element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//span/span/button"))
            ).click()
            
            self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//span[@data-qa='download-all-label-label-text']"))
            ).click()
            
            self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//span[@data-qa='download-document-label-label-text']"))
            ).click()
            
            buttons = self.driver.find_elements(By.XPATH, "//button")
            for button in buttons:
                if button.text == 'Baixar':
                    button.click()
            time.sleep(4)
        except:
            return False
        else:
            return True