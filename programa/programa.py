# Declaração de imports
import pandas as pd
from openpyxl import load_workbook
import os
from pathlib import Path

#Declaração de variáveis
information_dictionary = {}
list_information = []
path = Path(os.getcwd())
base_excel_path = str(path.parent).replace("\\", "/")
source_excel = 'ORIGEM2.xlsx'
destiny_excel = 'DESTINO2.xlsx'
source_excel_path = base_excel_path + '/excel/' + source_excel
destiny_excel_path =  base_excel_path + '/excel/' + destiny_excel
destiny_column_reference = 'CNPJ'

#Processo responsável por pegar uma lista dos CNPJ's no arquivo excel de Origem.
source = pd.read_excel(source_excel_path, sheet_name = 0)
quantity_cnpj_source = source.CNPJ.count()

for line in range(quantity_cnpj_source):
    cnpj = source.loc[line, "CNPJ"]
    cnpj_replace = int(str(cnpj).replace(".", "").replace("/", "").replace("-", ""))
    status = str(source.loc[line, "STATUS"])
    information_dictionary = {"cnpj": cnpj_replace, "status": status}
    list_information.append(information_dictionary)

list_information.append(information_dictionary)

#Processo responsável por realizar um de > para com o arquivo excel de Destino.
destiny = pd.read_excel(destiny_excel_path, sheet_name = 0)
quantity_customer_destiny = destiny.RESPONSÁVEL.count()

wbDestiny = load_workbook(destiny_excel_path)
wsDestiny = wbDestiny.active

for x in range(quantity_customer_destiny):
    for j in range(quantity_cnpj_source):
        if(int(list_information[j]["cnpj"]) == int(destiny.loc[x, destiny_column_reference])):
            wsDestiny.cell(row=x+2, column=17, value=list_information[j]["status"])
            break
    
#Salvar o excel.
wbDestiny.save(base_excel_path + '/excel/' + destiny_excel)