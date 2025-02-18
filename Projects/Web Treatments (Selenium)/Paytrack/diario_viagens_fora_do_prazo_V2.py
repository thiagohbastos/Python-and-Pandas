# %%
# ------------- IMPORTANTE BIBLIOTECAS -------------
import selenium
from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {'protocol_handler.excluded_schemes.tel': False})

import os
import time
import shutil
import pyautogui
import numpy as np
import pandas as pd
from unidecode import unidecode
from datetime import datetime, date, timedelta

import warnings
warnings.filterwarnings('ignore')

#print(selenium.__version__)



# %%
project_folder = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Fora do prazo"
#download_folder = r"C:\Users\b043469\Downloads"

# THIAGO - COMENTAR LINHA ACIMA E DESCOMENTAR ABAIXO:
download_folder = r"C:\Users\b042786\Downloads"


# %% 
# ## Funções scraping

# %%
def login_paytrack( username, password ):

    # Página de login Paytrack
    driver.get("https://login.paytrack.com.br/")

    # Digitar usuário
    wait.until(EC.visibility_of_element_located((By.ID, "normal_login_username"))).send_keys(username)
    # Botão próximo
    driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/div/form/div[4]/button").click()

    # Digitar senha
    wait.until(EC.visibility_of_element_located((By.ID, "normal_login_password"))).send_keys(password)
    # Botão Login
    driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/div/form/div[4]/div/button").click()
    #time.sleep(60)

    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    time.sleep(10)

    # Popup
    try:
        driver.find_element(By.XPATH, "/html/body/div[6]/div/div/div[2]/div[1]/span").click()
        time.sleep(5)
    except:
        pass

def baixar_relatorio( 
    relatorio,
    data_inicial, 
    data_final, 
    id_data_inicial, 
    id_data_final 
):
    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(5)
    
    # Pesquisar relatório    
    wait.until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ))).send_keys(Keys.CONTROL + "a")
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(Keys.DELETE)
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(relatorio)
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/button'
                    ).click()

    # Acessar relatório
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[6]/ul/li/div/div/p'
                    ).click()
    time.sleep(2)

    # Inserir Parâmetros
    wait.until(EC.visibility_of_element_located((By.ID, id_data_inicial
                    ))).send_keys(data_inicial)
    driver.find_element(By.ID, id_data_final
                    ).send_keys(data_final)
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[1]/select'
                    ).send_keys("EXCEL (XLSX)")

    # Download
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[2]/button[1]'
                    ).click()

def baixar_relatorio_duas_datas( 
    relatorio,
    data_inicial, 
    data_final, 
    id_data_inicial, 
    id_data_final,
    id_data_inicial_partida, 
    id_data_final_partida
):
    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(5)
    
    # Pesquisar relatório    
    wait.until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ))).send_keys(Keys.CONTROL + "a")
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(Keys.DELETE)
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(relatorio)
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/button'
                    ).click()

    # Acessar relatório
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[6]/ul/li/div/div/p'
                    ).click()
    time.sleep(2)

    # Inserir Parâmetros
    wait.until(EC.visibility_of_element_located((By.ID, id_data_inicial
                    ))).send_keys(data_inicial)
    driver.find_element(By.ID, id_data_final
                    ).send_keys(data_final)
    driver.find_element(By.ID, id_data_inicial_partida
                    ).send_keys(data_inicial)
    driver.find_element(By.ID, id_data_final_partida
                    ).send_keys(data_final)
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[1]/select'
                    ).send_keys("EXCEL (XLSX)")

    # Download
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[2]/button[1]'
                    ).click()

def buscar_historico_download(seconds):
    # Carregar página de histórico
    try:
        wrapper = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="reportList"]/relatorio-async-react/div/div/div[2]/div/div/div/div/div/div/div/div/table'
                                            ))
        )
    except TimeoutException:
        print("Element did not show up")
        
    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(seconds)

    # Acessar relatório
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[1]/div/div/button[2]'
                        ).click()
    try:
        wrapper = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[3]/div/div[2]/div[2]/div/div/div/div/div/div/div/div/table/tbody/tr[1]/td[5]/button'
                                        ))
        )
    except TimeoutException:
        print("Element did not show up")
    
    # Download
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[3]/div/div[2]/div[2]/div/div/div/div/div/div/div/div/table/tbody/tr[1]/td[5]/button'
                        ).click()

    # Voltar para página de relatórios
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[1]/div/div/button[1]'
                        ).click()
    
def mover_para_pasta( relatorio, download_path, destination_path ):
    # Esperar download
    while not os.path.exists(download_path):
        time.sleep(2)
    if os.path.isfile(download_path):
        print(f"{relatorio} downloaded")

    # Mover para pasta
    shutil.move(download_path, destination_path)
    print(f"{relatorio} moved")

# %% 
# # Download Relatórios

# %%
# Data final
yesterday = date.today() - timedelta(days=1)
yesterday = datetime.strftime(yesterday, "%d/%m/%Y")

# Inicializar webdriver
'''
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 240)
pw = "@Mercantil123"

########## Login ##########
login_paytrack("matheus.oandrade@mercantil.com.br", pw)'''

########## AGE01 ##########
### Parâmetros

'''
relatorio = "AGE01"
data_inicial = "01/01/2023"
data_final = yesterday
id_data_inicial = "edtData_data_inicial"
id_data_final = "edtData_data_final"
id_data_inicial_partida = "edtData_dataPartida_i"
id_data_final_partida = "edtData_dataPartida_f"
download_path = download_folder + r"\AGE01 - Emissões aéreo por agência.xlsx"
destination_path = project_folder + r"\AGE01 - Emissões aéreo por agência.xlsx"

baixar_relatorio_duas_datas( 
    relatorio,
    data_inicial, 
    data_final, 
    id_data_inicial, 
    id_data_final,
    id_data_inicial_partida, 
    id_data_final_partida
)

mover_para_pasta(
    relatorio,
    download_path,
    destination_path
)

########## AGE02 ##########
### Parâmetros
relatorio = "AGE02"
data_inicial = "01/01/2023"
data_final = yesterday
id_data_inicial = "edtData_data_inicial"
id_data_final = "edtData_data_final"
id_data_inicial_partida = "edtData_dataPartida_i"
id_data_final_partida = "edtData_dataPartida_f"
download_path = download_folder + r"\AGE02 - Emissões de hotel por agência.xlsx"
destination_path = project_folder + r"\AGE02 - Emissões de hotel por agência.xlsx"

baixar_relatorio_duas_datas( 
    relatorio,
    data_inicial, 
    data_final, 
    id_data_inicial, 
    id_data_final,
    id_data_inicial_partida, 
    id_data_final_partida
)

mover_para_pasta(
    relatorio,
    download_path,
    destination_path
)

########## SOL05 ##########
### Parâmetros
relatorio = "SOL05"
data_inicial = "01/01/2023"
data_final = yesterday
id_data_inicial = "edtData_data_inicio"
id_data_final = "edtData_data_fim"
download_path = download_folder + r"\SOL05 - Relatórios fora do prazo de antecedência.xlsx"
destination_path = project_folder + r"\SOL05 - Relatórios fora do prazo de antecedência.xlsx"

baixar_relatorio(
    relatorio,
    data_inicial,
    data_final,
    id_data_inicial,
    id_data_final
)

mover_para_pasta(
    relatorio,
    download_path,
    destination_path
)

driver.quit()'''

# %% 
# # SOL05

# %%
path = project_folder + r"\SOL05 - Relatórios fora do prazo de antecedência.xlsx"

sol05_raw = pd.read_excel(path, converters={'Código':int},
                      usecols="A,D,E,F,I", header=1) # Coluna N: valor total da viagem

sol05_raw.rename(columns={'Dias de \nantecedência':'Dias de antecedência',
                          'Código':'id',
                          'Justificativa':'Justificativa violação'}, inplace=True)
                          #'Valor':'Valor total viagem'}, inplace=True)
    
### COMBINAR LINHAS CORTADAS
sol05_raw['Rank'] = sol05_raw.groupby('id')['id'].transform('rank', method='first')
sol05_raw['Key'] = sol05_raw['id'].astype(str) + " - " + sol05_raw['Rank'].astype(str)

sol05_raw['Key'].replace('nan - nan', np.nan, inplace=True)
sol05_raw['Key'].ffill(inplace=True)

campos = ['Justificativa violação']
for c in campos:
    sol05_raw[c].fillna('', inplace=True)

sol05 = sol05_raw.groupby('Key').agg({ 'id': 'first'
                                     , 'Data criação': 'first'
                                     , 'Data inicio': 'first'
                                     , 'Dias de antecedência': 'first'
                                     , 'Justificativa violação': ' '.join
                                   }).reset_index(drop=True)

print(len(sol05))
sol05.head()

# %% 
# ## Justificativas

# %%
path = project_folder + r"\Justificativas.xlsx"

just = pd.read_excel(path)
justificativas = list(just.Justificativa)

agosto = sol05[sol05['Data inicio'].dt.to_period('M')=='2023-08']
agosto['Justificativa violação'] = agosto['Justificativa violação'].str.replace('  ', ' ')

for j in justificativas:
    agosto.loc[agosto['Justificativa violação'].str.contains(j, regex=False), 'Justificativa violação'] = j

df_just = agosto.groupby('Justificativa violação').agg({'id':'count'}).reset_index()
just = just.merge(df_just,
                  left_on='Justificativa',
                  right_on='Justificativa violação',
                  how='left'
                 ).drop(columns=['Justificativa violação', 'id_y'])
just.rename(columns={'id_x':'id'}, inplace=True)
just.fillna(0, inplace=True)

sol05_just = sol05.merge(just[['Justificativa', 'Aceitável']]
                         , left_on='Justificativa violação'
                         , right_on='Justificativa'
                         , how='left').drop(columns=['Justificativa'])
sol05_just

# %% 
# # AGE01 (Aéreo)

# %%
path = project_folder + r"\AGE01 - Emissões aéreo por agência.xlsx"

age01_raw = pd.read_excel(path,
                      converters={'Código':int},
                      usecols="A,B,C,D,F,I,J,K,X,Y,AH"
                     )

age01_raw.rename(columns={'Código': 'id'
                     ,'Data partida': 'Data inicio viagem'
                     ,'Valor R$': 'Valor despesa' 
                     }, inplace=True)

age01_raw['Data emissão'] = pd.to_datetime(age01_raw['Data emissão'], format='%d/%m/%y').dt.normalize()

### COMBINAR LINHAS CORTADAS
age01_raw['Rank'] = age01_raw.groupby('id')['id'].transform('rank', method='first')
age01_raw['Key'] = age01_raw['id'].astype(str) + " - " + age01_raw['Rank'].astype(str)

age01_raw['Key'].replace('nan - nan', np.nan, inplace=True)
age01_raw['Key'].ffill(inplace=True)

campos = ['Descrição', 'Motivo', 'Solicitante', 'Origem', 'Destino', 'Centro de custo', 'Projeto']
for c in campos:
    age01_raw[c].fillna('', inplace=True)

age01 = age01_raw.groupby('Key').agg({ 'id': 'first'
                                     , 'Descrição': ' '.join
                                     , 'Motivo': ' '.join
                                     , 'Solicitante': ' '.join
                                     , 'Data emissão': 'first'
                                     , 'Data inicio viagem': 'first'
                                     , 'Origem' : ' '.join
                                     , 'Destino': ' '.join
                                     , 'Centro de custo': ' '.join
                                     , 'Projeto': ' '.join
                                     , 'Valor despesa': 'first'
                                   }).reset_index(drop=True)

for c in ['Origem', 'Destino']:
    age01[c] = age01[c].str.replace('  ', ' ')
    age01[c] = age01[c].str.replace('Guarulhos (GRU)', 'Guarulhos')
    age01[c] = age01[c].str.replace('Congonhas (CGH)', 'Congonhas')
    age01[c] = age01[c].str.replace('Tancredo Neves (CNF)', 'Tancredo Neves')
    age01[c] = age01[c].str.replace('Salgado Filho (POA)', 'Salgado Filho')
    age01[c] = age01[c].str.strip()
    
### NOVAS COLUNAS
age01['Itinerário'] = age01['Origem'] + " - " + age01['Destino']

# age01['Itinerário - Data'] = age01['Origem'] + " - " + \
#                              age01['Destino'] + " - " + \
#                              (age01['Data emissão'].dt.month).astype('str')
age01['Tipo despesa'] = 'Aéreo'

### CONCAT COM HOTEL
age01['Diárias'] = 1
age01['Valor total'] = age01['Valor despesa']

### PARETO
pareto = age01.groupby('Itinerário').agg({'Valor despesa':'sum', 'id':'count'}).reset_index()
pareto.sort_values('id', ascending=False, inplace=True)
pareto['Cumulativo'] = pareto['id'].cumsum()
pareto['Pareto'] = pareto['Cumulativo'] / sum(pareto['id'])

top_trechos = pareto.head(6)
trechos = list(top_trechos['Itinerário'])

age01['Comparativo'] = np.where(age01['Itinerário'].isin(trechos), 1, 0)
age01['Competência'] = age01['Data inicio viagem'].dt.to_period('M')

age01.drop(columns=['Origem'], inplace=True)

print(len(age01))
age01.head()

# %% 
# ## Merge SOL05

# %%
df1 = age01.merge(sol05_just, on='id', how='left')

df1['Política violada'] = np.where(df1['Dias de antecedência'].notna(), 1, 0)

print(len(df1))
df1.head()

# %% 
# ## Comparativo

# %%
comparativo1 = df1[(df1.Comparativo==1)
                   & (df1.Motivo!='Diretoria [2]')
                   & (df1['Data emissão'].dt.month < 7)
                   & (df1['Data inicio viagem'].dt.month < 7)
                  ].groupby(['Itinerário', 'Política violada']).agg({ 'id':'count'
                                                                     ,'Valor despesa':['mean','median']
                                                                    }).reset_index()
comparativo1.columns = ['Itinerário', 'Política violada', 'Viagens', 'Valor médio', 'Valor mediano']

comparativo1['Economia possível'] = comparativo1['Valor médio']
comparativo1.loc[
    comparativo1['Política violada']==0, 'Economia possível'] = comparativo1['Economia possível']*(-1)

economia1 = comparativo1.groupby('Itinerário').agg({'Economia possível':'sum',
                                                    'Valor médio':'last'
                                                 }).reset_index()
economia1.rename(columns={'Valor médio': 'Estimado fora do prazo'}, inplace=True)
economia1['Economia percentual'] = economia1['Economia possível'] / economia1['Estimado fora do prazo']

economia_media_aereo = economia1['Economia percentual'].mean()
economia1

# %% 
# # AGE02 (Hotel)

# %%
path = project_folder + r"\AGE02 - Emissões de hotel por agência.xlsx"

age02 = pd.read_excel(path,
                      converters={'Código':int},
                      usecols="A,B,C,D,F,I,J,K,L,S,T,Y,AC"
                     )

age02.rename(columns={'Código': 'id'
                     ,'Local': 'Destino' 
                     ,'Checkin': 'Data inicio viagem'
                     ,'Tarifa': 'Valor total'
                     ,'Valor': 'Valor líquido'
                     }, inplace=True)

age02['Data emissão'] = pd.to_datetime(age02['Data emissão'], format='%d/%m/%y').dt.normalize()
age02['Data inicio viagem'] = pd.to_datetime(age02['Data inicio viagem'], format='%d/%m/%y').dt.normalize()
age02['Checkout'] = pd.to_datetime(age02['Checkout'], format='%d/%m/%y').dt.normalize()

#age02['Itinerário - Data'] = age02['Destino'] + " - " + (age02['Data emissão'].dt.month).astype('str')
age02['Tipo despesa'] = 'Hotel'

### QUERO COMPARAR O PREÇO DA DIÁRIA
age02['Diárias'] = (age02['Checkout'] - age02['Data inicio viagem']).dt.days
age02['Valor despesa'] = age02['Valor total'] / age02['Diárias']

age02['Hotel'] = age02['Hotel'].str.strip()
hoteis = [
    'NORMANDY HOTEL',
    'HOTEL VIVENZO SAVASSI',
    'GRAN HOTEL MORADA DO SOL',
    'DAYRELL HOTEL & CENTRO DE CONVENCOES',
    'GRANDE HOTEL AMPARO',
    'CASSINO TOWER HOTEL CAMPINAS',
    'Central Palace Hotel',
    'RADISSON BLU BELO HORIZONTE',
    'ESTANCIA AVARE HOTEL',
    'IBIS OURINHOS'
]

age02['Comparativo'] = np.where(age02['Hotel'].isin(hoteis), 1, 0)
age02['Competência'] = age02['Data inicio viagem'].dt.to_period('M')

age02.drop(columns=['Checkout','Valor líquido'], inplace=True)
age02 = age02[age02['id'].notna()]

print(len(age02))
age02.head()

# %% 
# ## Merge SOL05

# %%
df2 = age02.merge(sol05_just, on='id', how='left')

df2['Política violada'] = np.where(df2['Dias de antecedência'].notna(), 1, 0)
df2.rename(columns={'Hotel':'Itinerário'},inplace=True)

print(len(df2))
df2.head()

# %% 
# ## Comparativo

# %%
comparativo2 = df2[(df2.Comparativo==1)
                  & (df2.Motivo!='Diretoria [2]')
                  & (df2['Data inicio viagem'].dt.month < 7)
                 ].groupby(['Itinerário', 'Política violada']).agg({ 'id':'count'
                                                               ,'Valor despesa':['mean','median']
                                                               }).reset_index()
comparativo2.columns = ['Itinerário', 'Política violada', 'Viagens', 'Valor médio', 'Valor mediano']

comparativo2['Economia possível'] = comparativo2['Valor médio']
comparativo2.loc[
    comparativo2['Política violada']==0, 'Economia possível'] = comparativo2['Economia possível']*(-1)

economia2 = comparativo2.groupby('Itinerário').agg({'Economia possível':'sum',
                                                    'Valor médio':'last'
                                                 }).reset_index()
economia2.rename(columns={'Valor médio': 'Estimado fora do prazo'}, inplace=True)
economia2['Economia percentual'] = economia2['Economia possível'] / economia2['Estimado fora do prazo']

economia_media_hotel = economia2['Economia percentual'].mean()
economia2

# %% 
# # Concat

# %%
df = pd.concat([df1, df2])
df.sort_values('id', inplace=True)
df.reset_index(drop=True, inplace=True)

df['Número CC'] = df['Centro de custo'].str[:5]

print(len(df))
df.head()

# %% 
# ## Únicos

# %%
aereos_unicos = df1[df1['Data inicio viagem'] < '2023-08-28'].groupby('id').agg(
    {'Valor total':'sum',
     'Competência':'first',
     'Política violada':'first'}).reset_index()

hoteis_unicos = df2[df2['Data inicio viagem'] < '2023-08-28'].groupby('id').agg(
    {'Valor total':'sum',
    'Competência':'first',
    'Política violada':'first'}).reset_index()

df_dir = df[df.Motivo != 'Diretoria [2]']
viagens_unicas = df_dir[df_dir['Data inicio viagem'] < '2023-08-28'].groupby('id').agg(
    {'Valor total':'sum',
    'Competência':'first',
    'Política violada':'first'}).reset_index()

# %% 
# # Segmentos

# %%
path = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Data Lake\segmentos.xlsx"

segmentos = pd.read_excel(path, sheet_name="Exclusivos", usecols="A,B,E")
segmentos.rename(columns={'CC SOLICITANTE':'Número CC',
                          'NOME CC SOLICITANTE':'Nome CC',
                          'SEGMENTO':'Segmento'}, inplace=True)

segmentos['Número CC'] = segmentos['Número CC'].astype('str')

df_seg = df.merge(segmentos[['Número CC', 'Segmento']], on='Número CC', how='left')

print(len(df_seg))
df_seg.head()

# %% 
# ## Suporte comercial

# %%
df_seg.loc[ ((df_seg['Descrição'].str.lower().str.contains("suporte", regex=False)) |
             (df_seg['Descrição'].str.lower().str.contains("apoio", regex=False))) 
           & (df_seg['Política violada']==1)
          , 'Aceitável'] = 1

suporte = df_seg[(df_seg['Competência'] == '2023-08') & (df_seg['Política violada']==1)]
suporte['Aceitável'] = suporte['Aceitável'].fillna(0)

suporte = suporte[['id', 'Descrição', 'Justificativa violação', 'Aceitável']]

print(len(suporte))
suporte.head()

# %% 
# # Load

# %%
path = project_folder + r"\Viagens fora do prazo.xlsx"

df_aereo = df_seg[df_seg['Tipo despesa'] == 'Aéreo']
df_hotel = df_seg[df_seg['Tipo despesa'] == 'Hotel']

with pd.ExcelWriter(path) as writer:
    df_seg.to_excel(writer, sheet_name='Viagens', index=False)
    df_aereo.to_excel(writer, sheet_name='Aéreo', index=False)
    df_hotel.to_excel(writer, sheet_name='Hotel', index=False)
    aereos_unicos.to_excel(writer, sheet_name='Aéreos Únicos', index=False)
    hoteis_unicos.to_excel(writer, sheet_name='Hotéis Únicos', index=False)
    viagens_unicas.to_excel(writer, sheet_name='Viagens Únicas', index=False)
    comparativo1.to_excel(writer, sheet_name='Comparativo aéreo', index=False)
    comparativo2.to_excel(writer, sheet_name='Comparativo hotel', index=False)
    economia1.to_excel(writer, sheet_name='Economia aéreo', index=False)
    economia2.to_excel(writer, sheet_name='Economia hotel', index=False)
    suporte.to_excel(writer, sheet_name='Aceitáveis agosto', index=False)

# %%
file_to_copy1 = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Fora do prazo\Viagens fora do prazo.xlsx"
destination_directory1 = r"T:\SIC\Seg Patrimonial\Manutenções\Viagens\Viagens fora do prazo.xlsx"
shutil.copy(file_to_copy1, destination_directory1)

file_to_copy2 = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Fora do prazo\Justificativas.xlsx"
destination_directory2 = r"T:\SIC\Seg Patrimonial\Manutenções\Viagens\Justificativas.xlsx"
shutil.copy(file_to_copy2, destination_directory2)

file_to_copy3 = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Data Lake\hierarquia_cc.xlsx"
destination_directory3 = r"T:\SIC\Seg Patrimonial\Manutenções\Viagens\hierarquia_cc.xlsx"
shutil.copy(file_to_copy3, destination_directory3)

# %%
pyautogui.alert('O código foi finalizado. Você já pode utilizar o computador!')