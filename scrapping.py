import csv
import re
from typing import Any

from lxml import html
import requests
import xlsxwriter


# Arquivo final de XLS
workbook = xlsxwriter.Workbook('fiis.xlsx')
worksheet = workbook.add_worksheet() # Adiciona uma aba

sheetRow = 1
currency_format = workbook.add_format({'num_format': '[$R$ -416]#,##0.00'})

writers = {
    "A1": "ATIVO",
    "B1": "CNPJ",
    "C1": "RAZÃO SOCIAL",
    "D1": "COTAÇÃO ATUAL",
    "E1": "DY",
    "F1": "P/VP",
    "G1": "TOTAL DE COTISTAS",
    "H1": "TOTAL DE COTAS",
    "I1": "VALOR PATRIMÔNIAL POR COTA",
    "J1": "VALOR DO PATRIMÔNIO",
    "K1": "VALOR DE MERCADO",
    "L1": "LIQUIDEZ MEDIA 30D",
    "M1": "ULTIMO RENDIMENTO",
    "N1": "TIPO",
    "O1": "TIPO DE GESTÃO",
    "P1": "QTD. NEG. DIÁRIAS",
    "Q1": "SEGMENTO",
    "R1": "QTD DE ATIVOS",
}

# Use the worksheet object to write
# data via the write() method.
for key, val in writers.items():
    worksheet.write(key, val)

# Ler base de dados com os FIIs
data = []
with open('fonte.csv', "r") as csvfile:
    reader = csv.reader(csvfile, delimiter=";")
    for row in reader:
        data.append(row[0])

# Objetos com as informações
treeSI = Any
treeI10 = Any
treeFE = Any

# Lê as informações e salva no Excel
def processData(fii):
    sheetColumn = 0
    worksheet.write(sheetRow, sheetColumn, fii)
    sheetColumn += 1

    qtdAtivos = treeFE.xpath('(//*[@id="fund-actives-chart-info-wrapper"]/span[1])/text()') # QTD DE ATIVOS
    # STATUS INVEST
    dadosFII = [
        treeSI.xpath('//*[@id="fund-section"]/div/div/div[2]/div/div[1]/div/div/strong/text()')[0], # CNPJ
        treeSI.xpath('//*[@id="fund-section"]/div/div/div[2]/div/div[2]/div/div/div/strong/text()')[0], # NOME
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong/text()')[0], # COTAÇÃO ATUAL
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong/text()')[0], # DIVIDEND YELD
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[2]/div/div[1]/strong/text()')[0], # P/VP
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[6]/div/div[1]/strong/text()')[0], # TOTAL DE COTISTAS
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[6]/div/div[2]/span[2]/text()')[0], # TOTAL DE COTAS

        # TO-DO - TROCAR O ULTIMO PONTO DO VALOR POR VÍRGULA, PARA SEGUIR O PADRÃO.
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[1]/div/div[1]/strong/text()')[0].replace("R$","").replace(" ",""), # VALOR PATRIMONIAL POR COTA
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[1]/div/div[2]/span[2]/text()')[0].replace("R$","").replace(" ",""), # VALOR DO PATRIMONIO
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[2]/div/div[2]/span[2]/text()')[0].replace("R$","").replace(" ",""), # VALOR DE MERCADO
        treeSI.xpath('//*[@id="main-2"]/div[2]/div[6]/div/div/div[3]/div/div/div/strong/text()')[0].replace("R$","").replace(" ","").replace(",","."), # LIQUIDEZ MEDIA 30 DIAS
        treeSI.xpath('//*[@id="dy-info"]/div/div[1]/strong/text()')[0], # ULTIMO RENDIMENTO

        # INVESTIDOR 10
        treeI10.xpath('//*[@id="table-indicators"]/div[6]/div[2]/div/span/text()')[0].replace("\n",""), # TIPO DO FUNDO
        treeI10.xpath('//*[@id="table-indicators"]/div[8]/div[2]/div/span/text()')[0].replace("\n",""), # TIPO DE GESTÃO

        # FUNDS EXPLORER
        treeFE.xpath('(//span[@class="indicator-value"])[1]/text()')[0].replace("\n","").replace(" ",""), # QTD. DE NEGOCIAÇÕES DIÁRIAS
        treeFE.xpath('(//span[@class="description"])[12]/text()')[0].replace("\n","").replace(" ",""), # SEGMENTO
        re.sub("[^0-9]", "", qtdAtivos[0]),
    ] 
    #vacancia = ... PENDENTE
    
    # Gravação do XLSX
    for dado in dadosFII:
        if sheetColumn >= 8 and sheetColumn <= 11:
            worksheet.write(sheetRow, sheetColumn, dado, currency_format)
        else:
            worksheet.write(sheetRow, sheetColumn, dado)

        sheetColumn += 1

for fii in data:
    pageSI = requests.get('https://statusinvest.com.br/fundos-imobiliarios/' + fii)
    treeSI = html.fromstring(pageSI.content)

    pageI10 = requests.get('https://investidor10.com.br/fiis/' + fii)
    treeI10 = html.fromstring(pageI10.content)  

    pageFE = requests.get('https://www.fundsexplorer.com.br/funds/' + fii)
    treeFE = html.fromstring(pageFE.content)

    processData(fii)
    sheetRow +=1

# Aplicar filtro
lastRow = "R"+str(sheetRow)
range = "A1:" + lastRow
worksheet.autofilter(range)

workbook.close()
