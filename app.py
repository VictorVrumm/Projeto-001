from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

class Scrappy:
    
    def iniciar(self):
        self.coleta_de_dados()
        self.planilha()
    
    def coleta_de_dados(self):

        driver = webdriver.Chrome()
        driver.get("https://www.kabum.com.br/hardware/placa-de-video-vga/placa-de-video-nvidia")
        
        self.titulos = driver.find_elements(By.XPATH,"//span [@class='sc-d79c9c3f-0 nlmfp sc-27518a44-9 iJKRqI nameCard']")
        
        self.preco_original = driver.find_elements(By.XPATH,"//span [@class='sc-57f0fd6e-2 hjJfoh priceCard']")
    
    def planilha(self):

        #criando planilha

        workbook = openpyxl.Workbook()
        pagina_padrao = workbook.active
        workbook.remove(pagina_padrao)

        #criando pagina "Produtos"

        workbook.create_sheet('Produtos')
        sheet_Produtos = workbook['Produtos']
        sheet_Produtos['A1'].value = 'Nome do Produto'
        sheet_Produtos['B1'].value = 'Preço Original'

        #inserindo dados em Produtos

        for titulos, preco_original in zip(self.titulos, self.preco_original):
            sheet_Produtos.append([titulos.text, preco_original.text])
        
        workbook.save('Produtos.xlsx')

start = Scrappy()
start.iniciar()