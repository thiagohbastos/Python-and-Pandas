#%%
# ------------- PACOTES -------------
import pandas as pd
import pyodbc
import os
import pyautogui
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL
from time import sleep

# DESATIVAÇÃO DE ALERTAS (WARNINGS)
import warnings
warnings.filterwarnings('ignore')
warnings.simplefilter(action='ignore', category=FutureWarning)



# %%
# ------------- CONECTANDO SQL SERVER -------------
conn_gnu = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server=SQLGDNP;"
    "Database=GNU;"
    "Trusted_Connection=yes;")

conn_mbcorp = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server=SQLMBCORP;"
    "Database=MBCORP;"
    "Trusted_Connection=yes;")

cursor_gnu = conn_gnu.cursor()



#%%
# ------------- CRIANDO CONEXÃO SQLALCHEMY -------------
connection_string = (
  r"Driver=SQL Server Native Client 11.0;"
  r"Server=SQLGDNP;" #
  r"Database=GNU;"
  r"Trusted_Connection=yes;"
)
connection_url = URL.create(
  "mssql+pyodbc", 
  query={"odbc_connect": connection_string}
)
engine = create_engine(connection_url, fast_executemany=True, connect_args={'connect_timeout': 10}, echo=False)
conn_alchemy_gnu = engine.connect()



#%%
# ------------- BASES DE CONFERÊNCIA -------------
transportadoras = ['prosegur', 'brinks', 'protege']
query_cen_transp = """SELECT DISTINCT A.COD_CEN
	,A.DES_CEN
	,B.NOM_TRN
FROM CENTRALIZADORA_NUM A
LEFT JOIN TRANSPORTADORA B
ON B.IDT_TRN = A.IDT_TRN
ORDER BY 1
"""
cen_transp = pd.read_sql_query(query_cen_transp, conn_mbcorp)



#%%
# ------------- INICIANDO SELEÇÃO DE ARQUIVOS -------------
while True:
    print('''GOSTARIA DE EXECUTAR O VALIDADOR COM OS ARQUIVOS NA MESMA PASTA QUE ESTÁ ESTE EXECUTÁVEL?
          
[1] SIM
[2] VALIDAR EM OUTRO CAMINHO
[3] CANCELAR VALIDAÇÃO
''')
    try:
        resposta = int(input('Opção Selecionada: '))
        if resposta not in (1, 2, 3):
            print('Opção invalida! Responda novamente.\n')
        else:
            break
    except:
        print('Opção invalida! Responda novamente.\n')
if resposta == 1:
    pasta_valida = str(os.getcwd()).replace('\\', '/')
    print(f'\nIniciando a validação dos arquivos locais ({pasta_valida}) \n')
elif resposta == 3:
    print('\nVALIDAÇÃO CANCELADA\n')
    exit()
else:
    while True:
        pasta_valida = str(input('''Por favor, informe um caminho válido para validação dos arquivos de faturamento: \n''')).replace('\\', '/')
        if os.path.isdir(pasta_valida):
            print(f'\nIniciando a validação dos arquivos locais ({pasta_valida}) \n')
            pyautogui.alert(f'''A validação dos arquivos de faturamento da pasta {pasta_valida} iniciará em breve...
                            
Antes de proseguir, CERTIFIQUE-SE DE QUE APENAS OS ARQUIVOS PARA VALIDAÇÃO XLSX ESTEJAM NA PASTA INFORMADA.''', title='IMPORTANTE')
            pyautogui.PAUSE = 0.5
            break
        else:
            print('Caminho inválido! Responda novamente. ')

arquivos_pasta = os.listdir(pasta_valida)
arquivos = []
for arquivo in arquivos_pasta:
    if arquivo[-5:].upper() == '.xlsx'.upper() and arquivo[:2] != '~$':
        arquivos.append(pasta_valida + '/' + arquivo)
if len(arquivos) > 3:
    print(f'''Foram encontrados arquivos em excesso. 
Eram esperados 3 arquivos ou menos, porém foram encontrados {len(arquivos)} arquivos válidos na pasta.

O PROCESSO SERÁ ENCERRADO EM 20s caso o executável não seja fechado...''')
    sleep(20)
    exit()



#%%
# ------------- VALIDAÇÃO DE CUSTÓDIA -------------
colunas_custodia = ['COD_SVC', 'SVC', 'DATA', 'COD_CEN', 'DES_CEN', 'VALOR', 'VALOR_TRANSP',
       'CUSTO_PERNOITE', 'ISS', 'TOTAL']

for arquivo in arquivos:
    for transp in transportadoras:

        if transp.upper() in arquivo.upper():
            globals()[f'custodia_{transp}'] = pd.read_excel(arquivo, sheet_name='CUSTODIA')
            globals()[f'custodia_{transp}'] = pd.concat([pd.read_excel(arquivo, sheet_name='INCLUIR CUSTODIA')
                                                         , globals()[f'custodia_{transp}']])
            globals()[f'custodia_{transp}'] = globals()[f'custodia_{transp}'].query('TOTAL != 0 or VALOR_TRANSP !=0')

            globals()[f'filiais_{transp}'] = cen_transp[cen_transp['NOM_TRN'] == transp.upper()]

            temp = []
            for x in globals()[f'custodia_{transp}'].fillna(0).values:

                valida_0 = x[3] in list(globals()[f'filiais_{transp}']['COD_CEN'])
                #0,0116%
                if x[0] == 4:
                    valida_1 = x[6] - x[5] <= 10000 or x[6] - x[5] <= 0
                elif x[0] == 11:
                    valida_1 = x[6] - x[5] <= 1000  or x[6] - x[5] <= 0
                else:
                    temp.append(0)
                    continue

                valida_2 = x[9]<= x[6] * (0.0116/100) #O valor de custódia cobrado é 0.01155%

                if valida_0 and valida_1 and valida_2:
                    temp.append(1)
                else:
                    temp.append(0)
            
            globals()[f'custodia_{transp}']['APROVADO'] = temp
            
            globals()[f'custodia_{transp}_aprovada'] = pd.DataFrame(columns=[colunas_custodia])
            globals()[f'custodia_{transp}_reprovada'] = pd.DataFrame(columns=[colunas_custodia])



# %%
# ------------- EXTRAÇÃO BASE COMPLETA DE TRANSPORTE/ACOMP-TEC -------------
for arquivo in arquivos:
    for transp in transportadoras:
        if transp.upper() in arquivo.upper():
            globals()[f'transporte_{transp}'] = pd.read_excel(arquivo, sheet_name='BASE')
            globals()[f'transporte_{transp}'] = pd.concat([pd.read_excel(arquivo, sheet_name='INCLUIR BASE')
                                                           , globals()[f'transporte_{transp}']])
            globals()[f'transporte_{transp}'].fillna(0, inplace = True)
            globals()[f'transporte_{transp}'] = globals()[f'transporte_{transp}'].query('TOTAL != 0 or VALOR_TRANSP !=0')



# %%
# ------------- TRATANDO SVC 1: SUPRIMENTO -------------
for transp in transportadoras:
    globals()[f'sup_{transp}'] = globals()[f'transporte_{transp}'][globals()[f'transporte_{transp}']['COD_SVC'] == 1]
    


# %%
# ------------- TRATANDO SVC 2: RECOLHIMENTO -------------
for transp in transportadoras:
    globals()[f'rec_{transp}'] = globals()[f'transporte_{transp}'][globals()[f'transporte_{transp}']['COD_SVC'] == 2]



# %%
# ------------- TRATANDO SVC 3: ACOMP_TEC  -------------
for transp in transportadoras:
    globals()[f'acomp_tec_{transp}'] = globals()[f'transporte_{transp}'].query('COD_SVC in (8, 10)')



# %%
# ------------- TRATANDO SVC's 8 e 10 (DSI E INTERBASE) -------------
for transp in transportadoras:
    globals()[f'dsi_interbancario_{transp}'] = globals()[f'transporte_{transp}'].query('COD_SVC in (8, 10)')



# %%
# ------------- TRATANDO SVC's 6, 7 e 9  -------------
for transp in transportadoras:
    globals()[f'demais_transportes_{transp}'] = globals()[f'transporte_{transp}'].query('COD_SVC in (6, 7, 9)')
