import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
from pathlib import Path

list_cnpj = []
path = Path(os.getcwd())
base_excel_path = str(path.parent).replace("\\", "/")
source_excel = 'origem.xlsx'
destiny_excel = 'destino.xlsx'
source_excel_path = base_excel_path + '/excel/' + source_excel
destiny_excel_path =  base_excel_path + '/excel/' + destiny_excel
source_column_reference = 'CNPJ'
destiny_column_reference = 'CNPJ'

source = pd.read_excel(source_excel_path, sheet_name = 0)
quantity_cnpj_source = source.CNPJ.count()

for line in range(quantity_cnpj_source):
    cnpj = source.loc[line, source_column_reference]
    cnpj_replace = str(cnpj).replace(".", "").replace("/", "").replace("-", "")
    list_cnpj.append(int(cnpj_replace))

destiny = pd.read_excel(destiny_excel_path, sheet_name = 0)
quantity_cnpj_destiny = destiny.CNPJ.count()

wbDestiny = load_workbook(destiny_excel_path)
wsDestiny = wbDestiny.active

for x in range(quantity_cnpj_destiny):
    for j in range(quantity_cnpj_source):
        if(list_cnpj[j] == destiny.loc[x, destiny_column_reference]):
            wsDestiny.cell(row=x+2, column=2, value="QUENTE").fill = PatternFill(fgColor='00FF00', fill_type = 'solid')
    
wbDestiny.save(base_excel_path + '/excel/' + destiny_excel)