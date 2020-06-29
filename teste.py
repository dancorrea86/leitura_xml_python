import xml.etree.ElementTree as ET
import os
from openpyxl import Workbook

arrayWithNameFiles = []
arrayWithValues = []
folderWithFiles = './XML'

def organizeFilesInFolder(folder):
    for arquivo in sorted(os.listdir(folder)):
        arrayWithNameFiles.append(arquivo)

def chargeFilesToProcess(arrayWithNameFiles):
    for arquivo in arrayWithNameFiles:
        tree = ET.parse(folderWithFiles + '/' + arquivo)
        root = tree.getroot()
        arrayWithValues.append(joinValuesInArray(root))

def joinValuesInArray(root):
    valorNotaFiscal = retornaNotaFiscal(root)
    valorDestinatario = retornaDestinatario(root)
    valorIcms = retornaValorIcms(root)
    
    resultado = valorNotaFiscal + valorDestinatario + valorIcms

    return resultado

def retornaNotaFiscal(root):
    totais = []

    for total in root.iter('{http://www.portalfiscal.inf.br/nfe}ide'):
        nota = total.find('{http://www.portalfiscal.inf.br/nfe}nNF').text
        totais.append(nota)
    
    return totais

def retornaDestinatario(root):
    totais = []

    for total in root.iter('{http://www.portalfiscal.inf.br/nfe}dest'):
        for child in total.iter():
            totais.append(child.text)
        resultado = [totais[1],totais[2],totais[4],totais[5],totais[6],totais[8],totais[9],totais[10]]

    return resultado
    
def retornaValorIcms(root):
    resultado = []

    for total in root.iter('{http://www.portalfiscal.inf.br/nfe}total'):
        
        tagIcms = total.find('{http://www.portalfiscal.inf.br/nfe}ICMSTot')
        
        valorBaseCalculo = float(tagIcms.find('{http://www.portalfiscal.inf.br/nfe}vBC').text)
        valorIcms = float(tagIcms.find('{http://www.portalfiscal.inf.br/nfe}vICMS').text)
        valorNotaFiscal = float(tagIcms.find('{http://www.portalfiscal.inf.br/nfe}vNF').text)
        
 
        resultado.append(valorBaseCalculo)
        resultado.append(valorIcms)
        resultado.append(valorNotaFiscal)

    return resultado

def writeValuesInSheet(arrayWithValues):
    arrayWithColunms = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'L', 'M']
    lineIndex = 1
    book = Workbook()
    sheet = book.active

    for line in arrayWithValues:
        colunmIndex = 0

        for element in line:
            sheet[str(arrayWithColunms[colunmIndex]) + str(lineIndex)] = element
            colunmIndex += 1
        
        lineIndex += 1
    
    book.save('relatorio.xlsx')          

organizeFilesInFolder(folderWithFiles)
chargeFilesToProcess(arrayWithNameFiles)
writeValuesInSheet(arrayWithValues)



