from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

class DocusignController:
    def start_docusign_apprentice_process(document):
        options = webdriver.ChromeOptions()
        options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\User Data")
        options.add_argument(r'--profile-directory=Default')

        driver = webdriver.Chrome(options=options)

        driver.get("https://apps.docusign.com/send/documents")
        
        try:
            DocusignController.docusign_login(driver)
        except:
            return print("não foi possível efetuar o login")
        else:
            try:
                DocusignController.select_template(driver)
            except:
                return print("não foi possível selecionar o modelo do contrato")
            else:
                try:
                    DocusignController.select_document_to_sign(driver, document)
                except:
                    return print(f"não foi possível selecionar o arquivo: {document}")
                else:
                    try:
                        DocusignController.sign_document_select_zoom(driver)
                    except:
                        return print("não foi possível alterar o zoom")
                    else:
                        try:
                            campos_de_assinatura = DocusignController.start_sign_process(driver)
                        except:
                            return print("não foi possível finalizar o processo")
                        else:
                            return campos_de_assinatura
        
    def docusign_login(driver):
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
        )
        if element.text  == "NEXT":
            element.click()
            try:
                element = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
                )
                element.click()
            except:
                return print("erro no botão de login")
        else:
            element.click()
        
    def select_template(driver):
        time.sleep(4)
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='header-TEMPLATES-tab-button']"))
        )
        element.click()
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@data-qa='templates-main-header-form-input']"))
        )
        element.send_keys("Modelo de Contrato - CC")
        element.send_keys(Keys.ENTER)
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='templates-main-list-row-b4024bba-b898-4c4d-a1c5-3ef9b408ee3c-actions-use']"))
        )
        element.click()
            
    def select_document_to_sign(driver, document):
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "windows-drag-handler-wrapper"))
        )
        element.find_element(By.TAG_NAME, "input").send_keys(document)
        
        element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@data-qa='prepare-subject']"))
        )
        element.send_keys(Keys.CONTROL+"a")
        element.send_keys(Keys.DELETE)
        element.send_keys("Contrato teste")
        element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//textarea[@data-qa='prepare-message']"))
        )
        element.send_keys(Keys.CONTROL+"a")
        element.send_keys(Keys.DELETE)
        element.send_keys("mensagem teste")
        
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='footer-add-fields-link']"))
        )
        element.click()
            
    def sign_document_select_zoom(driver):
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//button[@data-qa='zoom-button']"))
        )
        element.click()
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//button[@data-qa='zoom-level-75']"))
        )
        element.click()
            
    def start_sign_process(driver):
        paginas_assinadas=[]
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//div[@data-qa='document-accordion-region']"))
        )
        rubrica = driver.find_element(By.XPATH, "//button[@data-qa='Initial']")
        assinatura = driver.find_element(By.XPATH, "//button[@data-qa='Signature']")
        paginas = element.find_elements(By.XPATH, "//button[@data-qa='tagger-documents']")
        try:
            DocusignController.sign_pages(driver, rubrica, assinatura)
        except:
            element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//button[@data-qa='modal-cancel-btn']"))
            )
            element.click()
        finally:
            DocusignController.sign_pages(driver, rubrica, assinatura)
            for pagina in paginas:
                alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
                paginas_assinadas.append(alt_value)
            return paginas_assinadas
            
    def sign_pages(driver, rubrica, assinatura, rubricaX=450, assinaturaX=332, rubricaY=370, assinaturaY=325):
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//div[@data-qa='document-accordion-region']"))
        )
        paginas = element.find_elements(By.XPATH, "//button[@data-qa='tagger-documents']")
        for i, pagina in enumerate(paginas):
            element = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='tagger-documents']"))
            )
            alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
            if "tem campos nesta página" not in alt_value:
                pagina.click()
                if i % 2 == 0 or i == 0:
                    ActionChains(driver)\
                        .click_and_hold(rubrica)\
                        .move_by_offset(rubricaX, rubricaY)\
                        .release()\
                        .perform()
                else:
                    ActionChains(driver)\
                        .click_and_hold(assinatura)\
                        .move_by_offset(assinaturaX, assinaturaY)\
                        .release()\
                        .perform()
                rubrica.click()
