# Imports
from typing import Any
from lxml import html
import requests
import re
import xlsxwriter
import csv

# Arquivo final de XLS
workbook = xlsxwriter.Workbook('fiis.xlsx')
worksheet = workbook.add_worksheet() # Adiciona uma aba

sheetRow = 1
currency_format = workbook.add_format({'num_format': '[$R$ -416]#,##0.00'})

# Use the worksheet object to write
# data via the write() method.
worksheet.write('A1', 'ATIVO')
worksheet.write('B1', 'CNPJ')
worksheet.write('C1', 'RAZÃO SOCIAL')
worksheet.write('D1', 'COTAÇÃO ATUAL')
worksheet.write('E1', 'DY')
worksheet.write('F1', 'P/VP')
worksheet.write('G1', 'TOTAL DE COTISTAS')
worksheet.write('H1', 'TOTAL DE COTAS')

worksheet.write('I1', 'VALOR PATRIMÔNIAL POR COTA')
worksheet.write('J1', 'VALOR DO PATRIMÔNIO')
worksheet.write('K1', 'VALOR DE MERCADO')
worksheet.write('L1', 'LIQUIDEZ MEDIA 30D')
worksheet.write('M1', 'ULTIMO RENDIMENTO')

worksheet.write('N1', 'TIPO')
worksheet.write('O1', 'TIPO DE GESTÃO')
worksheet.write('P1', 'QTD. NEG. DIÁRIAS')
worksheet.write('Q1', 'SEGMENTO')
worksheet.write('R1', 'QTD DE ATIVOS')

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
    dadosFII = []
    sheetColumn = 0
    worksheet.write(sheetRow, sheetColumn, fii)
    sheetColumn += 1

    # STATUS INVEST
    dadosFII.append(treeSI.xpath('//*[@id="fund-section"]/div/div/div[2]/div/div[1]/div/div/strong/text()')[0]) # CNPJ
    dadosFII.append(treeSI.xpath('//*[@id="fund-section"]/div/div/div[2]/div/div[2]/div/div/div/strong/text()')[0]) # NOME
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong/text()')[0]) # COTAÇÃO ATUAL
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong/text()')[0]) # DIVIDEND YELD
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[2]/div/div[1]/strong/text()')[0]) # P/VP
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[6]/div/div[1]/strong/text()')[0]) # TOTAL DE COTISTAS
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[6]/div/div[2]/span[2]/text()')[0]) # TOTAL DE COTAS

    # TO-DO - TROCAR O ULTIMO PONTO DO VALOR POR VÍRGULA, PARA SEGUIR O PADRÃO.
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[1]/div/div[1]/strong/text()')[0].replace("R$","").replace(" ","")) # VALOR PATRIMONIAL POR COTA
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[1]/div/div[2]/span[2]/text()')[0].replace("R$","").replace(" ","")) # VALOR DO PATRIMONIO
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[2]/div/div[2]/span[2]/text()')[0].replace("R$","").replace(" ","")) # VALOR DE MERCADO
    dadosFII.append(treeSI.xpath('//*[@id="main-2"]/div[2]/div[6]/div/div/div[3]/div/div/div/strong/text()')[0].replace("R$","").replace(" ","").replace(",",".")) # LIQUIDEZ MEDIA 30 DIAS
    dadosFII.append(treeSI.xpath('//*[@id="dy-info"]/div/div[1]/strong/text()')[0]) # ULTIMO RENDIMENTO

    # INVESTIDOR 10
    dadosFII.append(treeI10.xpath('//*[@id="table-indicators"]/div[6]/div[2]/div/span/text()')[0].replace("\n","")) # TIPO DO FUNDO
    dadosFII.append(treeI10.xpath('//*[@id="table-indicators"]/div[8]/div[2]/div/span/text()')[0].replace("\n","")) # TIPO DE GESTÃO

    # FUNDS EXPLORER
    dadosFII.append(treeFE.xpath('(//span[@class="indicator-value"])[1]/text()')[0].replace("\n","").replace(" ","")) # QTD. DE NEGOCIAÇÕES DIÁRIAS
    dadosFII.append(treeFE.xpath('(//span[@class="description"])[12]/text()')[0].replace("\n","").replace(" ","")) # SEGMENTO

    qtdAtivos = treeFE.xpath('(//*[@id="fund-actives-chart-info-wrapper"]/span[1])/text()') # QTD DE ATIVOS
    dadosFII.append(re.sub("[^0-9]", "", qtdAtivos[0]))
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
