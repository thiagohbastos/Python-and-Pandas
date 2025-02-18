# %%
import os
import pyautogui
import numpy as np
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

import warnings
warnings.filterwarnings('ignore')

# %%
path = r"K:\GSAS\06 - Coordenação Gestão Compras, Logistica e Doctos\03 - Logística\01 - Faturamento\01 - Posições Contratuais"

# %% [markdown]
# ## Rotina

# %%
# Leitura originais

to_append = []
total_len = []

for file_name in os.listdir(path):
    if file_name.startswith('EBCT'):
        f = os.path.join(path, file_name)
        
        try:
            df = pd.read_excel(f, sheet_name='REALIZADO')
        except ValueError:
            try:
                df = pd.read_excel(f, sheet_name='Realizado')
            except ValueError:
                try:
                    df = pd.read_excel(f, sheet_name='Realizado Geral')
                except ValueError:
                    try:
                        df = pd.read_excel(f, sheet_name='EBCT - Multiextrato')
                    except ValueError:
                        try:
                            df = pd.read_excel(f, sheet_name='MENSAL')
                        except ValueError:
                            df = pd.read_excel(f, sheet_name='Detalhamento')
        
        df['Arquivo origem'] = file_name.split('.')[0].split('EBCT ')[-1]
        
        to_append.append(df)
        
        print(file_name, len(df), len(df.columns))

to_append[0]['CONTA'] = ""
to_append[1]['Valor Unitário'] = ""
#to_append[5].drop(columns=['Desconto'], inplace=True)
to_append[6].drop(columns=['Desconto'], inplace=True)
to_append[7].rename(columns={'Custo Unitário': 'Valor Unitário'}, inplace=True)
to_append[8].rename(columns={'Custo Unitário': 'Valor Unitário'}, inplace=True)
to_append[9].drop(columns=['DESCONTO'], inplace=True)
to_append[10].rename(columns={'Valor unitário': 'Valor Unitário'}, inplace=True)
to_append[12].rename(columns={'QUANTIDADE': 'QTDE', 'VALOR UNIT.': 'Valor Unitário'}, inplace=True)
to_append[13].rename(columns={'Unnamed: 6': 'Valor Unitário'}, inplace=True)
to_append[14].drop(columns=['Unnamed: 11'], inplace=True)       
        
ebct = pd.concat(to_append).reset_index(drop=True)

ebct = ebct[ebct.FORNECEDOR.notna()]

print(len(ebct))
ebct.head()

# %%
file = path + r"\CONSOLIDADO - EBCT V2.xlsx"

with pd.ExcelWriter(file) as writer:
    ebct.to_excel(writer, sheet_name='Consolidado EBCT', index=False)

pyautogui.alert('O código foi finalizado. Você já pode utilizar o computador!')