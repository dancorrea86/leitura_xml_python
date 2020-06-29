import xml.etree.cElementTree as ET
import requests
import os
import pandas as pd
import numpy as np
from openpyxl import Workbook

arrayComArquivos = []
arrayComFonenedores = []
arrayComItens = []

def ler_aquivo():
    for arquivo in os.listdir('./XML'):
        arrayComArquivos.append('./XML/' + arquivo)

ler_aquivo()

# print(arrayComArquivos)

def geraArrayComClientes():
    for arquivo in arrayComArquivos:
        tree = ET.parse(arquivo)
        root = tree.getroot()
        arrayComFonenedores.append([root[0][0][0][1].text, root[0][0][2][0].text, root[0][0][2][1].text, root[0][0][2][2][5].text])
        arrayComItens.append([root[0][0][0][1].text, root[0][0][3][0][2].text, root[0][0][3][0][3].text, root[0][0][3][0][5].text, 
                            root[0][0][3][0][9].text])
        arrayComItens.append([root[0][0][0][1].text, root[0][0][4][0][2].text, root[0][0][4][0][3].text, root[0][0][4][0][5].text, 
                            root[0][0][4][0][9].text])
        # arrayComItens.append([root[0][0][0][1].text, root[0][0][5][0][2].text, root[0][0][5][0][3].text, root[0][0][5][0][5].text, 
        #                     root[0][0][5][0][9].text])

geraArrayComClientes()

for item in arrayComItens:
    print (len(item))

# book = Workbook()
# sheet = book.active

# n = 1

# for item in arrayComFonenedores:
#     indexNF = "A" + str(n)
#     indexCNPJ = "B" + str(n)
#     indexName = "C" + str(n)
    

#     sheet[indexNF] = item[0]
#     sheet[indexCNPJ] = item[1]
#     sheet[indexName] = item[2]

#     n = n + 1

# book.save('sample.xlsx')

book2 = Workbook()
sheet = book2.active

n = 1

for item in arrayComItens:
    indexNF = "A" + str(n)
    indexProduto = "B" + str(n)
    indexNCM = "C" + str(n)
    indexCFOP = "D" + str(n)
    indexValor = "E" + str(n)
        
    

    sheet[indexNF] = item[0]
    sheet[indexProduto] = item[1]
    sheet[indexNCM] = item[2]
    sheet[indexCFOP] = item[3]
    sheet[indexValor] = item[4]

    n = n + 1

book2.save('sample2.xlsx')