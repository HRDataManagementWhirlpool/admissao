from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

import time
from pikepdf import Pdf
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
        
        self.wait = WebDriverWait(self.driver, 60)

#################### FUNÇÕES ########################

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
        while True:
            try:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['zoom_btn']['xpath']))
                ).click()
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['zoom_btn']['zoom_50'])) ## Zoom 50 para rodar em monitores e no notebook sem quebrar
                ).click()
            except:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()
            else:
                break

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
                element = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()
                time.sleep(1)
                rubrica.click()
        self.finish_process()

    def finish_process(self):
        self.driver.quit()

    def select_template_mens(self):
        self.driver.get("https://apps.docusign.com/send/templates?view=favorites")
        element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, self.SITE_MAP['inputs']['template_searchbar']['xpath']))
        )
        element.send_keys("CONTRATO DE TRABALHO -")
        element.send_keys(Keys.ENTER)
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='templates-main-list-row-5875e90e-4707-4b13-ae7e-7f34dde2d7cb-actions-use']"))
        ).click()

    def array_mensalista(self):
        return [
            {"nome": "GABRIELA CRISTINA DE GODOY SCHWAB", "re": "10231589", "email": "gabriela.gschwab@gmail.com"},
            {"nome": "CARLOS EDUARDO DAMACENO", "re": "10231590", "email": "carlosedamaceno@outlook.com.br"},
            {"nome": "FERNANDO ZIMMERMANN", "re": "10231591", "email": "fernandoz.zimmermann@gmail.com"},
            {"nome": "MARCO TULIO OTAVIO DA CRUZ", "re": "10231593", "email": "mgmtoc@gmail.com"},
            {"nome": "ZELIO GERALDO DOS SANTOS", "re": "10231594", "email": "zeliogsantos@yahoo.com.br"},
            {"nome": "MARCOS ANTONIO VIEIRA", "re": "10231635", "email": "antonio.marcos.aa8@gmail.com"},
            {"nome": "GABRIELA ALONSO", "re": "10231636", "email": "gabriela.2405@gmail.com"},
            {"nome": "PAULO CESAR APARECIDO SILVA", "re": "10231852", "email": "cidaos@gmail.com"}
        ]

    def select_document_to_sign_mens(self, document, mensalist, assunto):
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='quick-send-advanced-edit']"))
        ).click()
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='upload-file-button']"))
        ).click()
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@data-qa='upload-file-input']"))
        ).send_keys(document)
        
        element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Mensalista']"))
        )
        nome = element.find_elements(By.XPATH, "//input[@data-qa='recipient-name']")
        email = element.find_elements(By.XPATH, "//input[@data-qa='recipient-email']")
        
        nome[0].send_keys(mensalist['nome'])
        nome[0].send_keys(Keys.ESCAPE)
        email[0].send_keys(mensalist['email'])
        email[0].send_keys(Keys.ESCAPE)
        
        nome[4].send_keys(mensalist['nome'])
        nome[4].send_keys(Keys.ESCAPE)
        email[4].send_keys(mensalist['email'])
        email[4].send_keys(Keys.ESCAPE)
        
        element = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['inputs']['document_subject']['xpath']))
        )
        element.send_keys(Keys.CONTROL+"a")
        element.send_keys(Keys.DELETE)
        element.send_keys(assunto)
        
        element = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['document_next_btn']['xpath']))
        ).click()

    def sign_pages_mens(self):
        while True:
            try:
                self.wait.until(
                    EC.presence_of_element_located((By.XPATH, self.SITE_MAP['inputs']['pages_area']['xpath']))
                )
                contratos = self.driver.find_elements(By.XPATH, "//div[@data-qa='doc-thumbnail-list']/button")
            except:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()
            else:
                break
    
        while True:
            try:
                if contratos[0].get_dom_attribute('aria-expanded') == 'true':
                    contratos[0].click()
                if contratos[1].get_dom_attribute('aria-expanded') == 'false':
                    contratos[1].click()
                break
            except:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()
                
        rubrica = self.driver.find_element(By.XPATH, self.SITE_MAP['buttons']['rubrica_btn']['xpath'])
        assinatura = self.driver.find_element(By.XPATH, self.SITE_MAP['buttons']['assinatura_btn']['xpath'])
        selec_assinante = self.driver.find_element(By.XPATH, "//button[@data-qa='recipient-selector']")
        paginas = self.driver.find_elements(By.XPATH, self.SITE_MAP['buttons']['pages_to_sign_btn']['xpath'])
        
        while True:
            try:
                selec_assinante.click()
                break
            except:
                element = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                ).click()
        
        count = 0
        while (len(paginas)*2) != count:
            count = len(self.driver.find_elements(By.XPATH, self.SITE_MAP['inputs']['pages_signed']['xpath']))
            try:
                for i, pagina in enumerate(paginas):
                    element = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['pages_to_sign_btn']['xpath']))
                    )
                    alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
                    if "têm campos nesta página" not in alt_value: #########FAZER VERIFICAÇÃO COM OS NOMES DOS ASSINANTES
                        pagina.click()
                        count = len(self.driver.find_elements(By.XPATH, self.SITE_MAP['inputs']['pages_signed']['xpath']))
                        if i != 2: #### Adicionar mais uma assinatura para o mensalista
                            count = len(self.driver.find_elements(By.XPATH, self.SITE_MAP['inputs']['pages_signed']['xpath']))
                            if count % 2 == 0:
                                selec_assinante.click()
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-1']"))
                                ).click()
                                ActionChains(self.driver)\
                                    .click_and_hold(rubrica)\
                                    .move_by_offset(232, 200)\
                                    .release()\
                                    .perform()
                            count = len(self.driver.find_elements(By.XPATH, self.SITE_MAP['inputs']['pages_signed']['xpath']))
                            if count % 2 == 1:
                                selec_assinante.click()
                                self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-2']"))
                                ).click()
                                ActionChains(self.driver)\
                                    .click_and_hold(rubrica)\
                                    .move_by_offset(332, 200)\
                                    .release()\
                                    .perform()
                        else:
                            selec_assinante.click()
                            self.wait.until(
                                EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-1']"))
                            ).click()
                            ActionChains(self.driver)\
                                .click_and_hold(assinatura)\
                                .move_by_offset(132, 145)\
                                .release()\
                                .perform()
                            selec_assinante.click()
                            self.wait.until(
                                EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='recipient-2']"))
                            ).click()
                            ActionChains(self.driver)\
                                .click_and_hold(assinatura)\
                                .move_by_offset(258, 145)\
                                .release()\
                                .perform()
                        rubrica.click()
            except:
                try:
                    self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
                    ).click()
                except:
                    pass
                time.sleep(1)
                rubrica.click()

#################### PROCESSOS ########################

    def apprentice_process(self, pdf_file):
        self.open_website()
        self.login()
        self.select_template()
        self.select_document_to_sign(pdf_file)
        self.sign_document_select_zoom()
        self.sign_pages()
        self.finish_process()

    def monthly_process(self, pdf_folder):
        self.open_website()
        self.login()
        folder = os.listdir(pdf_folder)
        mensalistas = self.array_mensalista()
        for mensalista in mensalistas:
            for file in folder:
                if mensalista['re'] in file:
                    file_path= os.path.join(pdf_folder, file)
                    self.select_template_mens()
                    self.select_document_to_sign_mens(file_path, mensalista, file[:-4])
                    self.sign_document_select_zoom()
                    self.sign_pages_mens()
        self.finish_process()
