import csv
import re
from typing import Any

from lxml import html
import requests
import xlsxwriter

# XLSX file
workbook = xlsxwriter.Workbook('fiis.xlsx')
worksheet = workbook.add_worksheet()

sheetRow = 1
currency_format = workbook.add_format({'num_format': '[$R$ -416]#,##0.00'})

# Columns
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

# Use the worksheet object to write data via the write() method.
for key, val in writers.items():
    worksheet.write(key, val)

# Read FIIs code from database
data = []
with open('fonteCompleta.csv', "r") as csvfile:
    reader = csv.reader(csvfile, delimiter=",")
    for row in reader:
        data.append(row[0])

# Objects with FIIs informations
treeSI = Any
treeI10 = Any
treeFE = Any

# Replace last ocurrence
def replace_right(source, target, replacement, replacements=1):
    if type(source) is str:
        return replacement.join(source.rsplit(target, replacements))
    else:
        return source

# Check position and return data
def checkData(data):
    if data:
        value = data[0].replace("\n","").rstrip().lstrip()
        return value
    else:
        return ""

def checkValues(data):
    if data:
        value = data[0].replace("\n","").replace("R$","").replace(" ","").replace(",", ".")
        if value == "N/A" or value == "-":
            return 0
        else:
            return value
    else:
        return 0

# Read and save informations
def processData(fii):
    print(fii)

    sheetColumn = 0
    worksheet.write(sheetRow, sheetColumn, fii)
    sheetColumn += 1

    # STATUS INVEST
    pvp = checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[2]/div/div[1]/strong/text()')) # P/VP
    cotacaoAtual = checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong/text()')) # COTAÇÃO ATUAL
    dividendYeld = checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong/text()')) # DIVIDEND YELD
    vlrPatrimonialPorCota = checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[1]/div/div[1]/strong/text()')) # VALOR PATRIMONIAL POR COTA
    vlrPatrimonio = checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[1]/div/div[2]/span[2]/text()')) # VALOR DO PATRIMONIO
    vlrMercado = checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[2]/div/div[2]/span[2]/text()')) # VALOR DE MERCADO
    lqdzDiaria = checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[6]/div/div/div[3]/div/div/div/strong/text()')) # LIQUIDEZ MEDIA 30 DIAS    
    ultimoRendimento = checkValues(treeSI.xpath('//*[@id="dy-info"]/div/div[1]/strong/text()')) # ULTIMO RENDIMENTO

    # FUND EXPLORER
    qtdAtivos = checkValues(treeFE.xpath('(//*[@id="fund-actives-chart-info-wrapper"]/span[1])/text()')) # QTD DE ATIVOS
    if qtdAtivos:
        qtdAtivos = re.sub("[^0-9]", "", qtdAtivos)
    else:
        qtdAtivos = 0

    dadosFII = [
        # STATUS INVEST
        checkData(treeSI.xpath('//*[@id="fund-section"]/div/div/div[2]/div/div[1]/div/div/strong/text()')), # CNPJ
        checkData(treeSI.xpath('//*[@id="fund-section"]/div/div/div[2]/div/div[2]/div/div/div/strong/text()')), # NOME
        replace_right(cotacaoAtual, ".", ","),
        replace_right(dividendYeld, ".", ","),
        replace_right(pvp, ".", ","),
        checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[6]/div/div[1]/strong/text()')), # TOTAL DE COTISTAS
        checkValues(treeSI.xpath('//*[@id="main-2"]/div[2]/div[5]/div/div[6]/div/div[2]/span[2]/text()')), # TOTAL DE COTAS

        replace_right(vlrPatrimonialPorCota, ".", ","),
        replace_right(vlrPatrimonio, ".", ","),
        replace_right(vlrMercado, ".", ","),
        replace_right(lqdzDiaria, ".", ","),
        replace_right(ultimoRendimento, ".", ","),

        # INVESTIDOR 10
        checkData(treeI10.xpath('//*[@id="table-indicators"]/div[6]/div[2]/div/span/text()')), # TIPO DO FUNDO
        checkData(treeI10.xpath('//*[@id="table-indicators"]/div[8]/div[2]/div/span/text()')), # TIPO DE GESTÃO

        # FUNDS EXPLORER
        checkValues(treeFE.xpath('(//span[@class="indicator-value"])[1]/text()')), # QTD. DE NEGOCIAÇÕES DIÁRIAS
        checkData(treeFE.xpath('(//span[@class="description"])[12]/text()')), # SEGMENTO
        qtdAtivos

        # TO-DO - OBTER OS DADOS DA VACANCIA
    ]
    
    # Write XLSX file
    for dado in dadosFII:
        # TO-DO - CONVERTER TODOS OS CAMPOS DE NÚMERO PARA TIPO DE NÚMERO, POIS ESTÁ INDO COMO TEXTO.
        if sheetColumn >= 8 and sheetColumn <= 12:
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
    treeFE = html.fromstring(pageFE.content.decode('utf-8'))

    processData(fii)
    sheetRow +=1

# Apply filter
lastRow = "R"+str(sheetRow)
range = "A1:" + lastRow
worksheet.autofilter(range)

workbook.close()
