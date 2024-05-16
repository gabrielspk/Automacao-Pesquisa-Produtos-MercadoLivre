from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui as tempoEspera
from openpyxl import Workbook
import os

#Selecionado o caminho a ser salvo o arquivo excel
caminhoNovoArquivo = r'C:\Users\gabriel.ferreira\Music\Arquivos estudo PYTHON\estudo próprio\Arquivo_Desafio_RPA\MercadoLivre.xlsx'

#criando a planilha nova
criandoPlanilhaMercado = Workbook()

#ativando a planilha
planilhaMercado = criandoPlanilhaMercado.active

#dando um nome a sheet atual
planilhaMercado.title = "Dados"

#selecionado a sheet
sheet = criandoPlanilhaMercado['Dados']

#utilizando a variável navegador para manipular o selenium
navegador = opcoesSelenium.Chrome()
navegador.get("https://www.mercadolivre.com.br/")

#tempo de espera
tempoEspera.sleep(2)

#pesquisando os produtos "carteira" na barra de pesquisa
navegador.find_element(By.NAME, "as_word").send_keys("Carteira")

#clicando para pesquisar
navegador.find_element(By.XPATH, "/html/body/header/div/div[2]/form/button").click()

#declarando variável para contabilizar o laço while e o paginador
i = 3

#declarando a variável root para ser usado na formatação da string do páginador, para não usar duas "aspas"
root = '"root-app"'

while i < 7:
    
    #pegando os dados da UL do navegador
    dadosProdutos = navegador.find_elements(By.CLASS_NAME, "ui-search-layout__item")
    
    #percorrendo cada elementos do elemento li
    for produto in dadosProdutos:
            
            #pegando o nome do produto pela classe e convertendo para texto
            nomeProduto = produto.find_element(By.CLASS_NAME, "ui-search-item__title").text
            #pegando o preco do produto pela classe e convertendo para texto
            precoProduto = produto.find_element(By.CLASS_NAME, "andes-money-amount__fraction").text

            #adicionando uma tratativa caso não haja centavos no valor
            try:
                centavosProduto = produto.find_element(By.CLASS_NAME, "andes-money-amount__cents").text
            except:
                centavosProduto = "00"

            #pega a url do produto
            urlProduto = produto.find_element(By.TAG_NAME, "a").get_attribute("href")

            #contabilizando a quantidade de linhas na planilhas, para iterar a cada repetição do laço
            linhaPlanilha = len(sheet['A']) + 1

            #juntando a coluna com a linha
            colunaA = 'A' + str(linhaPlanilha)
            colunaB = 'B' + str(linhaPlanilha)
            
            #adicionando o nome do produto a coluna A
            sheet[colunaA] = nomeProduto
            #adicionando o preço do produto a coluna B
            sheet[colunaB] = precoProduto + "," + centavosProduto 

            print(f"{nomeProduto} // {precoProduto},{centavosProduto}")

    #passando para a próxima página de acordo com a iteração da variável i
    navegarPaginador = navegador.find_element(By.XPATH, f"//*[@id={root}]/div/div[3]/section/nav/ul/li[{i}]/button")

    #indo até o elemento selecionado estar visível, para clicar com eficácia no paginador
    navegador.execute_script("arguments[0].scrollIntoView();", navegarPaginador)
    
    #clicando no páginador selecionado
    navegarPaginador.click()

    #somando a variável i para o laço, e para ser utilizada no paginador
    i+=1

    tempoEspera.sleep(3)

#salva as alterações na planilha
criandoPlanilhaMercado.save(filename=caminhoNovoArquivo)

#após finalizado, inicia o arquivo
os.startfile(caminhoNovoArquivo)

tempoEspera.sleep(30)