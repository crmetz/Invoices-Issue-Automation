from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from time import sleep
from PIL import Image, ImageWin
import win32print
import win32ui
import os
from pdf2image import convert_from_path
from selenium.webdriver.common.alert import Alert
import sys
from selenium.common.exceptions import NoSuchElementException
import json
import re

class EmissaoNfe:
    def __init__(self, paginas, num_pedidos, half_screen_width, screen_height):
        script_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

        self.download_directory = os.path.join(script_dir, 'nfes')
        self.chrome_driver_path = os.path.join(script_dir, 'chromedriver.exe')
        self.output_folder = os.path.join(script_dir, 'nfes')

        self.paginas = paginas
        self.num_pedidos = num_pedidos
        self.half_screen_width = half_screen_width
        self.screen_height = screen_height 
        self.driver = None
        

    def convert_pdf_to_images(self, pdf_path, output_folder):
        images = convert_from_path(pdf_path, 300)

        os.makedirs(output_folder, exist_ok=True)

        for index, image in enumerate(images):
            image_path = os.path.join(output_folder, f'{pdf_path}page_{index + 1}.png')
            image.save(image_path, 'PNG')
            self.print_image(image_path)

    def print_image(self, image_path):
        PHYSICALWIDTH = 110
        PHYSICALHEIGHT = 111

        printer_name = win32print.GetDefaultPrinter()

        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)

        bmp = Image.open(image_path)

        hDC.StartDoc(image_path)
        hDC.StartPage()

        dib = ImageWin.Dib(bmp)
        dib.draw(hDC.GetHandleOutput(), (0, 0, printer_size[0], printer_size[1]))

        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()

    def change_page(self, nPag):
        xpath_expression = '//*[@id="LayerLista"]/div/div/ul/li[' + str(nPag) + ']'
        pageBtn = WebDriverWait(self.driver, 25).until(EC.presence_of_element_located((By.XPATH, xpath_expression)))
        pageBtn.click()

    def return_main_tab(self):
        self.driver.get('https://app.vhsys.com.br/index.php?Secao=Vendas.Pedidos&Modulo=Vendas')

    def login(self):
        with open('config.json', 'r') as config_file:
            config = json.load(config_file)

        username = config.get('username')
        password = config.get('password')

        chromeOptions = webdriver.ChromeOptions()
        prefs = {"download.default_directory": self.download_directory}
        chromeOptions.add_experimental_option("prefs", prefs)

        service = Service(self.chrome_driver_path)
        self.driver = webdriver.Chrome(service=service, options=chromeOptions)

        self.driver.get('https://app.vhsys.com.br/index.php?Secao=Vendas.Pedidos&Modulo=Vendas')

        username_field = self.driver.find_element('id', 'login')
        password_field = self.driver.find_element('id', 'senha')
        username_field.send_keys(username)
        password_field.send_keys(password)
        password_field.send_keys(Keys.RETURN)

    def run_automation(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.login()

        i = 0
        for numero_pedido in self.num_pedidos:


            if self.paginas[i]:
                self.change_page(self.paginas[i])
                print("Trocando a pagina")
            if numero_pedido:
                print("Emitir nfe")
                chave_acesso = self.emitir_nfe(numero_pedido)
                sleep(4)
                #chave de acesso e ja colocar dinamico
                pdf_path = os.path.join(script_dir, r'nfes\{}.pdf'.format(chave_acesso))
                self.convert_pdf_to_images(pdf_path, self.output_folder)
                self.return_main_tab()
            i += 1
        self.driver.quit()

        #self.driver.quit()
    def emitir_nfe(self, numero_pedido):
        WebDriverWait(self.driver, 25).until(EC.presence_of_element_located((By.ID, 'btn-inserir')))
        print("First driverwait finished")
        row_element = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located((By.XPATH, f'//tr[@numero="{numero_pedido}"]'))
        )

        last_td = row_element.find_element(By.XPATH, './/td[last()]')
        button_div = last_td.find_element(By.ID, 'opcoes_nota')
        button_div.click()
        
        dropdown_menu = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, 'actions-wrapper'))
        )

        # Clica para emitir nota
        emitir_btn = dropdown_menu.find_element(By.ID, 'emitir')
        emitir_btn.click()

        #Coloca a natureza do pedido
        natureza_field = WebDriverWait(self.driver, 15).until(
             EC.visibility_of_element_located((By.ID, 'natureza_pedido'))
        )
        natureza_field = self.driver.find_element('id', 'natureza_pedido')
        natureza_field.send_keys('5104')
        sleep(5)

        autocomplete = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'autocomplete_1'))
        )
        
        autocomplete.click()
        sleep(10)


        #Abre o menu de confirmação
        salvar_emitir = WebDriverWait(self.driver, 25).until(
             EC.visibility_of_element_located((By.ID, 'salvar-emitir'))
        )
        salvar_emitir.click()

        #Abre o alert de confirmação
        emitir_nota = WebDriverWait(self.driver, 15).until(
            EC.visibility_of_element_located((By.ID, "alterar"))
        )
        emitir_nota.click()

        #Emiti a nota
        WebDriverWait(self.driver, 10).until(EC.alert_is_present())
        self.driver.switch_to.alert.accept()

        #Baixa a nota
        download_btn = WebDriverWait(self.driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div[17]/div[1]/div[2]/div[2]/div[1]/form/table/tbody/tr/td/div[2]/div/div/div[2]/div/div/div/a"))
        )
        sleep(5)
        print("Antes de baixar a nota")
        download_btn.click()

        # Salva a chave de acesso(nome do arquivo)
        try:
            # Localize o elemento pai que contém o texto que você deseja acessar
            element_pai = WebDriverWait(self.driver, 15).until(
                EC.visibility_of_element_located((By.XPATH, "//*[@id='Layer_Transmitir']/div/div/div[2]/div/div"))
            )

            # Obtenha todo o texto dentro do elemento pai
            texto_pai = element_pai.text

            # Use regex para encontrar a chave de acesso dentro do texto pai
            chave_acesso_match = re.search(r'Chave de acesso: (\d+)', texto_pai)

            if chave_acesso_match:
                chave_acesso = chave_acesso_match.group(1)
                print("Número da Chave de Acesso:", chave_acesso)
                return chave_acesso
            else:
                print("Chave de acesso não encontrada.")
                sleep(200)
                sleep(200)
            #     print("First sleep finished")
                return None

        except NoSuchElementException as e:
            # Lidar com a exceção específica que ocorre quando o elemento não é encontrado
            print("Erro ao encontrar o elemento:", e)
            # Adicione seus comandos de sono aqui
            sleep(200)
            print("First sleep finished")
            sleep(200)
            print("Second sleep finished")
            sleep(200)
            return None
        except Exception as e:
            # Lidar com outras exceções não previstas
            print("Erro não esperado:", e)
            # Adicione seus comandos de sono aqui
            sleep(200)
            print("First sleep finished")
            sleep(200)
            print("Second sleep finished")
            sleep(200)
            return None
        

if __name__ == "__main__":
    orders_automation = EmissaoNfe()
    orders_automation.run_automation()


    