# %%
import os
import time
import shutil
#import pyodbc
#import getpass
import pyautogui
import numpy as np
import pandas as pd
from unidecode import unidecode
from datetime import datetime, date, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {'protocol_handler.excluded_schemes.tel': False})

import warnings
warnings.filterwarnings('ignore')

# %%
project_folder = r"K:\GSAS\00 - Gerência\001 - Atividades e projetos da Gerência\11-2022 - VIAGENS - Relatório Gestão Mensal"
download_folder = r"C:\Users\b043469\Downloads"

# THIAGO - COMENTAR LINHA ACIMA E DESCOMENTAR ABAIXO:
# download_folder = r"C:\Users\b042786\Downloads"

# %%
pw = "@Mercantil123"

# %% [markdown]
# ## Funções Scraping

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

def baixar_relatorio_criterio( 
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
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[1]/table/tbody/tr/td[1]/div/div[4]/div[2]/div[1]/input'
                    ).send_keys("S")
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

# %% [markdown]
# ## Download relatórios

# %%
# Datas
last_month_final = date.today().replace(day=1) - timedelta(days=1)
last_month_initial = last_month_final.replace(day=1)
last_month_final_str = datetime.strftime(last_month_final, "%d/%m/%Y")
last_month_initial_str = datetime.strftime(last_month_initial, "%d/%m/%Y")

# Inicializar webdriver
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 240)

########## Login ##########
login_paytrack("matheus.oandrade@mercantil.com.br", pw)

########## RDV11 ##########
### Parâmetros
relatorio = "RDV11"
data_inicial = last_month_initial_str
data_final = last_month_final_str
id_data_inicial = "edtData_dt_inicial_despesa"
id_data_final = "edtData_dt_final_despesa"
download_path = download_folder + r"\RDV11 - Despesas (Exportação).xlsx"
destination_path = project_folder + fr"\Bases mensais\{str(last_month_final.year)}\{str(last_month_final.month)}\despesas_paytrack.xlsx"

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

########## RDV16 ##########
### Parâmetros
relatorio = "RDV16"
data_inicial = "01/01/2022"
data_final = last_month_final_str
id_data_inicial = "edtData_dt_inicial_despesa"
id_data_final = "edtData_dt_final_despesa"
download_path = download_folder + r"\RDV16 - Despesas com política violada (Exportação).xlsx"
destination_path = project_folder + r"\Bases mensais\Violações\Limite Diário.xlsx"

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

########## GER19 ##########
### Parâmetros
relatorio = "GER19"
data_inicial = "01/01/2022"
data_final = last_month_final_str
id_data_inicial = "edtData_data_inicial"
id_data_final = "edtData_data_final"
download_path = download_folder + r"\GER19 - Saving lost (Exportação).xlsx"
destination_path = project_folder + r"\Bases mensais\Violações\Comparativo aéreo.xlsx"

baixar_relatorio_criterio(
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

########## SOL05 ##########
### Parâmetros
relatorio = "SOL05"
data_inicial = "01/01/2022"
data_final = last_month_final_str
id_data_inicial = "edtData_data_inicio"
id_data_final = "edtData_data_fim"
download_path = download_folder + r"\SOL05 - Relatórios fora do prazo de antecedência.xlsx"
destination_path = project_folder + r"\Bases mensais\Violações\Antecedência.xlsx"

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

########## SOL06 ##########
### Parâmetros
relatorio = "SOL06"
data_inicial = "01/01/2022"
data_final = last_month_final_str
id_data_inicial = "edtData_data_inicio"
id_data_final = "edtData_data_fim"
download_path = download_folder + r"\SOL06 - Relatórios com políticas de aéreo violadas.xlsx"
destination_path = project_folder + r"\Bases mensais\Violações\Aéreo.xlsx"

baixar_relatorio(
    relatorio,
    data_inicial,
    data_final,
    id_data_inicial,
    id_data_final
)

buscar_historico_download(240)

mover_para_pasta(
    relatorio,
    download_path,
    destination_path
)

########## SOL07 ##########
### Parâmetros
relatorio = "SOL07"
data_inicial = "01/01/2022"
data_final = last_month_final_str
id_data_inicial = "edtData_data_inicio"
id_data_final = "edtData_data_fim"
download_path = download_folder + r"\SOL07 - Relatórios com políticas de hotel violadas.xlsx"
destination_path = project_folder + r"\Bases mensais\Violações\Hotel.xlsx"

baixar_relatorio(
    relatorio,
    data_inicial,
    data_final,
    id_data_inicial,
    id_data_final
)

buscar_historico_download(240)

mover_para_pasta(
    relatorio,
    download_path,
    destination_path
)

########## SOL08 ##########
### Parâmetros
relatorio = "SOL08"
data_inicial = "01/01/2022"
data_final = last_month_final_str
id_data_inicial = "edtData_data_inicio"
id_data_final = "edtData_data_fim"
download_path = download_folder + r"\SOL08 - Relatórios com políticas de carro violadas.xlsx"
destination_path = project_folder + r"\Bases mensais\Violações\Carro.xlsx"

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

########## CAD06 ##########
### Parâmetros
relatorio = "CAD06"
data_inicial = "01/01/2022"
data_final = last_month_final_str
id_data_inicial = "edtData_dataInicial"
id_data_final = "edtData_dataFinal"
download_path = download_folder + r"\CAD06 - Prestação de contas aprovadas.xlsx"
destination_path = project_folder + r"\Bases mensais\ajuste_data.xlsx"

baixar_relatorio(
    relatorio,
    data_inicial,
    data_final,
    id_data_inicial,
    id_data_final
)

buscar_historico_download(240)

mover_para_pasta(
    relatorio,
    download_path,
    destination_path
)

# %% [markdown]
# ## 1. Despesas Paytrack

# %%
### 2022 ###
path = project_folder + r"\Bases mensais\2022\despesas_paytrack_2022.xlsx"

paytrack22 = pd.read_excel(path, converters={'#':int, 'Matricula':str, 'CPF':str},
                         usecols="A:I,K:Y,AB:AD")

paytrack22.rename(columns={"#":"id", "Identificador":"Número CC"},inplace=True)
paytrack22['Colaborador'] = paytrack22.Colaborador.apply(unidecode) # Remover acentos
paytrack22['Matricula'] = paytrack22.Matricula.str.replace('.','')
paytrack22['Data de criação'] = pd.to_datetime(paytrack22['Data de criação'], format = '%d/%m/%Y')
paytrack22['Data despesa'] = pd.to_datetime(paytrack22['Data despesa'], format = '%d/%m/%Y')
print(len(paytrack22))

# %%
### 2023 ###
path = project_folder + r"\Bases mensais\2023"
to_append = []
total_len = []

for mes in os.listdir(path):
    f = os.path.join(path, mes, "despesas_paytrack.xlsx")
    
    df = pd.read_excel(f, converters={'#':int, 'Matricula':str, 'CPF':str}, usecols="A:I,K:Y,AB:AD")

    df.rename(columns={"#":"id", "Identificador":"Número CC"},inplace=True)
    df['Colaborador'] = df.Colaborador.apply(unidecode) # Remover acentos
    df['Matricula'] = df.Matricula.str.replace('.','')
    df['Data de criação'] = pd.to_datetime(df['Data de criação'], format = '%d/%m/%Y')
    df['Data despesa'] = pd.to_datetime(df['Data despesa'], format = '%d/%m/%Y')

    to_append.append(df)
    total_len.append(len(df))
    print(mes, len(df))
    
paytrack23 = pd.concat(to_append)
print(len(paytrack23))
paytrack23.head()

# %%
### 2024 ###
path = project_folder + r"\Bases mensais\2024"
to_append = []
total_len = []

for mes in os.listdir(path):
    f = os.path.join(path, mes, "despesas_paytrack.xlsx")
    
    df = pd.read_excel(f, converters={'#':int, 'Matricula':str, 'CPF':str}, usecols="A:I,K:Y,AB:AD")

    df.rename(columns={"#":"id", "Identificador":"Número CC"},inplace=True)
    df['Colaborador'] = df.Colaborador.apply(unidecode) # Remover acentos
    df['Matricula'] = df.Matricula.str.replace('.','')
    df['Data de criação'] = pd.to_datetime(df['Data de criação'], format = '%d/%m/%Y')
    df['Data despesa'] = pd.to_datetime(df['Data despesa'], format = '%d/%m/%Y')

    to_append.append(df)
    total_len.append(len(df))
    print(mes, len(df))
    
paytrack24 = pd.concat(to_append)
print(len(paytrack24))
paytrack24.head()

# %%
### Concatenar
paytrack = pd.concat([paytrack22, paytrack23, paytrack24])
print(len(paytrack))
paytrack.head()

# %%
### Upload na pasta
path = project_folder + r"\Despesas_Paytrack.xlsx"

with pd.ExcelWriter(path) as writer:
    paytrack.to_excel(writer, sheet_name='Despesas Paytrack', index=False)

# %% [markdown]
# ### 1.1. Group By

# %%
### Agrupar base de despesas por id da viagem
grouped = paytrack.groupby('id').agg({'Valor':'sum', 'Data despesa':'max'})
print(len(grouped))

base_id = grouped.merge(
    paytrack[['id', 'Descrição', 'Data de criação', 'Colaborador',
              'Matricula', 'Número CC', 'Centro de custo', 'Projeto']],
    on='id', how='left').drop_duplicates().sort_values('id').reset_index(drop=True)

#base_loc['Id Duplicado'] = base_loc['id'].duplicated(keep=False)

base_id['Count id'] = base_id.groupby('id')['id'].transform('count')
base_id['Rank id'] = base_id.groupby('id')['id'].transform('rank', method='first')

base_id.rename(columns={'Valor':'Valor total'}, inplace=True)
base_id['Valor'] = base_id['Valor total'] / base_id['Count id']

base_id['Matricula'] = pd.to_numeric(base_id['Matricula'], errors='coerce')

base_id['Origem'] = 'Paytrack'

print(len(base_id))
base_id

# %% [markdown]
# ## 2. RH Cadastro

# %%
### Buscar csv do RH Cadastro (Notes)
path = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Data Lake\rh_cadastro.csv"

rh_cadastro = pd.read_csv(path)
rh_cadastro.columns = ['Código Usuário', 'Nome do Usuário', 'Data Admissão', 'Cargo', 'Unidade', 'Centro de Custo', 'Nome Notes', 'Numero Referencial', 'E-mail', 'Matrícula', 'Dígito', 'Empresa']
rh_cadastro.rename(columns={"Nome do Usuário":"Colaborador"},inplace=True)
rh_cadastro['CC Lotação'] = rh_cadastro['Centro de Custo'].astype('str') + " - " + rh_cadastro['Unidade']
rh_cadastro = rh_cadastro.sort_values('Data Admissão').drop_duplicates('Colaborador',keep='last').reset_index(drop=True)

print(len(rh_cadastro))
rh_cadastro.head()

# %%
### Join com a base de viagens
base_rh = base_id.merge(rh_cadastro[['Colaborador', 'Cargo', 'CC Lotação']], on="Colaborador",how="left")
print(len(base_rh))
base_rh.head()

# %% [markdown]
# ## 3. Segmentos

# %%
### Busca segmentos no "data lake". Atualização mensal opcional
path = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Data Lake\segmentos.xlsx"

segmentos = pd.read_excel(path, sheet_name="Exclusivos", usecols="A,B,C:E")
segmentos.rename(columns={'CC SOLICITANTE':'Número CC', 'NOME CC SOLICITANTE':'Centro de custo',
                          'CC GERENCIA / DIRETORIA':'Nº CC Subordinador',
                          'NOME CC GERENCIA / DIRETORIA':'CC Subordinador', 'SEGMENTO':'Segmento'},
                 inplace=True)

conditions = [(segmentos['CC Subordinador'].str.contains('INSS MG', na=False)),
              (segmentos['CC Subordinador'].str.contains('INSS SP', na=False))]
values = ['MG','SP']
segmentos['UF'] = np.select(conditions, values)
segmentos['UF'].replace('0', '', inplace=True)

path2 = project_folder + r"\Bases mensais\ajuste_uf.xlsx"

ajuste_uf = pd.read_excel(path2, sheet_name="Ajuste", usecols="A,C")

segmentos = segmentos.merge(ajuste_uf, on="Número CC", how="left")
segmentos['UF_y'].fillna(segmentos['UF_x'], inplace=True)
segmentos.drop(columns=['UF_x'], inplace=True)
segmentos.rename(columns={'UF_y':'UF'}, inplace=True)

print(len(segmentos))
segmentos.head()

# %%
### Join com a base de viagens
base_seg = base_rh.merge(segmentos[['Número CC', 'Nº CC Subordinador', 'CC Subordinador', 'Segmento', 'UF']],
                         on="Número CC",how="left")

base_seg['Nº CC Subordinador'].fillna(0, inplace=True)
base_seg.fillna("", inplace=True)
base_seg['Nº CC Subordinador'] = base_seg['Nº CC Subordinador'].astype('int64')

print(len(base_seg))
base_seg.head()

# %%
print('ok')

# %% [markdown]
# ## 4. Ajustes Diretoria

# %%
# Responsáveis
path = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Data Lake\centro_custo.csv"
centro_custo = pd.read_csv(path, usecols=[0,1,2,6], sep=',', encoding='latin')
centro_custo.columns = ['Número CC', 'Centro de Custo', 'Colaborador', 'CC Subordinador']
resp = centro_custo[centro_custo['Colaborador']!=0][['Número CC', 'Colaborador']]

centro_custo.head()

# %% [markdown]
# ### 4.1. 2022

# %%
# Primus

## Consolidado 1º Semestre
path1 = project_folder + r"\Faturamento\Primus 22 S1\Consolidado Geral 1º Sem. 2022.xlsx"
s1 = pd.read_excel(path1, sheet_name='Consolidado', usecols="A,B,C,F,N,R,T,U")
s1.columns = ['Data de criação','Passageiro','Descrição','Número CC','Valor','Data despesa','Faixa','Tipo despesa']

for i in [1125, 2758, 2829, 2878]:
    s1.iloc[i, s1.columns.get_loc('Faixa')] = 'DIR - INTERN.'

s1 = s1[(s1['Faixa']=='DIR') | (s1['Faixa']=='DIR - INTERN.') | (s1['Faixa']=='ADM EXT')].reset_index(drop=True)
s1.loc[((s1['Faixa']=='DIR - INTERN.') | (s1['Faixa']=='ADM EXT')) &
       (s1['Tipo despesa'] == 'AEREO'), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
s1.drop(columns=['Faixa', 'Tipo despesa'], inplace=True)

s1['Data despesa'] = s1['Data despesa'] - pd.Timedelta(days=1)

## Mensais 2º Semestre
path2 = project_folder + r"\Faturamento\Primus 22 S2"

to_append = []
total_len = []

for file_name in os.listdir(path2):
    
    f = os.path.join(path2, file_name)
    df = pd.read_excel(f, sheet_name='Consolidado')
    df.columns = df.columns.str.lower()
    df = df[(df['faixa/ cargo']=='DIR') | (df['faixa/ cargo']=='DIR EXT') | (df['faixa/ cargo']=='ADM EXT')]
    df['Referência'] = file_name.split('.')[0]
    
    to_append.append(df)
    total_len.append(len(df))
    print(file_name, len(df))
    
s2 = pd.concat(to_append).reset_index(drop=True)

s2 = s2[['data', 'passageiro', 'fornecedor', 'cc', 'destino', 'total', 'tipo', 'Referência', 'faixa/ cargo']]
s2.columns = ['Data de criação', 'Passageiro', 'Descrição', 'Número CC',
              'Destino', 'Valor', 'Tipo despesa', 'Data despesa', 'Faixa']
s2.loc[(s2['Faixa'].str.contains('EXT')) & (s2['Tipo despesa'] == 'AEREO'), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
s2.loc[(s2['Destino'].str.contains('ATL')) | (s2['Destino'].str.contains('DOH')), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
s2['Descrição'] = s2['Descrição'] + " - " + s2['Destino']
s2.drop(columns=['Faixa', 'Destino', 'Tipo despesa'], inplace=True)

primus22 = pd.concat([s1,s2]).reset_index(drop=True)
#primus22.drop(columns=['Passageiro'], inplace=True)

primus22.loc[primus22['Passageiro'].str.contains('HITOSI', case=False), 'Passageiro'] = 'HITOSI HASSEGAWA'
primus22.loc[primus22['Passageiro'].str.contains('BOFF', case=False), 'Passageiro'] = 'FELIPE LOPES BOFF'
primus22.loc[primus22['Passageiro'].str.contains('PAULINO', case=False), 'Passageiro'] = 'PAULINO RAMOS RODRIGUES'
primus22.loc[primus22['Passageiro'].str.contains('UELQUES', case=False), 'Passageiro'] = 'UELQUESNEURIAN RIBEIRO DE ALMEIDA'
primus22.loc[primus22['Passageiro'].str.contains('GREGORIO', case=False), 'Passageiro'] = 'GREGORIO MOREIRA FRANCO'
primus22.loc[primus22['Passageiro'].str.contains('BRUNO', case=False), 'Passageiro'] = 'BRUNO PINTO SIMAO'
primus22.loc[primus22['Passageiro'].str.contains('ANDERSON', case=False), 'Passageiro'] = 'ANDERSON ADEILSON DE OLIVEIRA'
primus22.loc[(primus22['Passageiro'].str.contains('GUSTAVO', case=False)) &
             (primus22['Passageiro'].str.contains('ARAUJO', case=False)), 'Passageiro'] = 'GUSTAVO HENRIQUE DINIZ DE ARAUJO'

nomes = {'GUERRA/DANIEL' : 'DANIEL GUERRA',
'FIGUEIREDO/CESAR ADRIANO' : 'CESAR ADRIANO FIGUEIREDO',
'FORESTI RIBEIRO/VALERIA' : 'VALERIA DE ARAUJO FORESTI RIBEIRO',
'HORTA/ANDRE RODRIGUES MR' : 'ANDRE RODRIGUES HORTA',
'HORTA/ANDRE RODRIGUES' : 'ANDRE RODRIGUES HORTA',
'HORTA/ANDRE' : 'ANDRE RODRIGUES HORTA',
'MELO DE ARAUJO/GLAUCIA MRS' : 'GLAUCIA MELO DE ARAUJO',
'MELO DE ARAUJO/GLAUCIA MR' : 'GLAUCIA MELO DE ARAUJO',
'DE ARAUJO/LUIZ HENRIQUE MR' : 'LUIZ HENRIQUE ANDRADE DE ARAUJO',
'LUIZ HENRIQUE  DE ARAUJO' : 'LUIZ HENRIQUE ANDRADE DE ARAUJO',
'DUARTE/CAROLINE' : 'CAROLINE DUARTE',
'DE ARAUJO/LUIZ HENRIQUE' : 'LUIZ HENRIQUE ANDRADE DE ARAUJO',
'DE ARAUJO/GLAUCIA' : 'GLAUCIA MELO DE ARAUJO',
'REZENDE/VALCI BRAGA' : 'VALCI BRAGA REZENDE',
'ROHRING/TERESINHA' : 'TERESINHA DA SILVA ROHRIG',
'LUIZ HENRIQUE  ANDRADE DE ARAUJO' : 'LUIZ HENRIQUE ANDRADE DE ARAUJO',
'ARAUJO/PAULO HENRIQUE' : 'PAULO HENRIQUE BRANT DE ARAUJO',
'FERNANDES/MARCOS' : 'MARCOS FERNANDES',
'CRUZ/TAISE CHRISTINE DA MRS' : 'TAISE CHRISTINE DA CRUZ',
'CRUZ/TAISE CHRISTINE DA' : 'TAISE CHRISTINE DA CRUZ',
'MOURA/GUSTAVO HENRIQUE MR' : 'GUSTAVO HENRIQUE CASSIMIRO MOURA',
'SANTIAGO/RICARDO VIEIRA' : 'RICARDO VIEIRA SANTIAGO',
'ROBSON MARCELO MACHADO SANTIAGO' : 'ROBSON MARCELO MACHADO SANTIAGO',
'LEO ADRIANO BORTON' : 'LEO ADRIANO BORTON',
'SANTIAGO/RICARDO' : 'RICARDO VIEIRA SANTIAGO',
'PENIDO/EULER LUIZ' : 'EULER LUIZ DE OLIVEIRA PENIDO',
'ADRIANO BORTON/LEO' : 'LEO ADRIANO BORTON',
'MARCELO MACHADO SANTIAGO/ROBSON' : 'ROBSON MARCELO MACHADO SANTIAGO',
'MARCELO MACHADO SANTIAGO/ROBSON MR' : 'ROBSON MARCELO MACHADO SANTIAGO',
'COSTA FILHO/JOAO MR' : 'JOAO VICENTE BARRETO DA COSTA FILHO',
'GIULIANI/ROBERTO MR' : 'ROBERTO GIULIANI',
'VIEIRA SANTIAGO/RICARDO' : 'RICARDO VIEIRA SANTIAGO',
'GIULIANI/ROBERTO' : 'ROBERTO GIULIANI',
'ZIEGELMEYER/RICARDO' : 'RICARDO ZIEGELMEYER',
'BRANCO/FLAVIO RIO' : 'FLAVIO RIO BRANCO FILHO',
'SANTOS/VINICIUS CUNHA' : 'VINICIUS CUNHA SANTOS',
'HORTA/ANDRÉ RODRIGUES' : 'ANDRE RODRIGUES HORTA',
'SILVA/ROBERTH MACEDO' : 'ROBERTH MACEDO SILVA',
'LOPES KUBIAKI/LUCAS' : 'LUCAS LOPES KUBIAKI',
'SILVA/JEFERSON ALVES DA' : 'JEFERSON ALVES DA SILVA',
'MIRANDA/ANTONIO JOSE COSTA' : 'ANTONIO JOSE COSTA MIRANDA',
'BARROS/CARLA RIBEIRO' : 'CARLA RIBEIRO BARROS',
'ANDRADE DE ARAUJO/LUIZ HENRIQUE' : 'LUIZ HENRIQUE ANDRADE DE ARAUJO',
'DE OLIVEIRA/PEDRO HENRIQUE' : 'PEDRO HENRIQUE DE OLIVEIRA',
'FERREIRA/LEONARDO' : 'LEONARDO FERREIRA',
'VALCI BRAGA REZENDE' : 'VALCI BRAGA REZENDE',
'BRAGA REZENDE/VALCI' : 'VALCI BRAGA REZENDE',
'LOPES LEANDRO/ROBERT' : 'ROBERT LOPES LEANDRO'}

primus22.replace(nomes, inplace=True)

print(len(s1), len(s2), len(primus22))
primus22.head()

# %%
# EBTA

path = project_folder + r"\Faturamento\EBTA 22"

to_append = []
total_len = []

for file_name in os.listdir(path):
    
    f = os.path.join(path, file_name)
    #f = path + f"{file_name}"
    df = pd.read_excel(f, sheet_name='Dados')
    df.columns = df.columns.str.lower()
    df = df[(df['faixa']=='DIR') | (df['faixa']=='DIR EXT') | (df['faixa']=='ADM EXT')]
    df['Referência'] = file_name.split('.')[0]
    
    to_append.append(df)
    total_len.append(len(df))
    print(file_name, len(df))
    
ebta22 = pd.concat(to_append).reset_index(drop=True)
ebta22.columns = ['Data de criação', 'Descrição', 'Número CC', 'Valor', 'faixa', 'Passageiro', 'Data despesa']
ebta22['Tipo despesa'] = np.where(ebta22['faixa'].str.contains('EXT'), 'AEREO INTERNACIONAL', 'AEREO')
ebta22.loc[(ebta22['Descrição'].str.contains('SFO')) | 
           (ebta22['Descrição'].str.contains('DOH')) |
           (ebta22['Descrição'].str.contains('IST')), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
ebta22.drop(columns=['faixa', 'Tipo despesa'], inplace=True)

ebta22.loc[ebta22['Passageiro'].str.contains('PAULINO', case=False), 'Passageiro'] = 'PAULINO RAMOS RODRIGUES'
ebta22.loc[ebta22['Passageiro'].str.contains('RAMOS RODRIGUES', case=False), 'Passageiro'] = 'PAULINO RAMOS RODRIGUES'
ebta22.loc[ebta22['Passageiro'].str.contains('BOFF', case=False), 'Passageiro'] = 'FELIPE LOPES BOFF'
ebta22.loc[ebta22['Passageiro'].str.contains('VALCI', case=False), 'Passageiro'] = 'VALCI BRAGA REZENDE'
ebta22.loc[ebta22['Passageiro'].str.contains('TAISE', case=False), 'Passageiro'] = 'TAISE CHRISTINE DA CRUZ'
ebta22.loc[ebta22['Passageiro'].str.contains('UELQUES', case=False), 'Passageiro'] = 'UELQUESNEURIAN RIBEIRO DE ALMEIDA'
ebta22.loc[ebta22['Passageiro'].str.contains('GREGORIO', case=False), 'Passageiro'] = 'GREGORIO MOREIRA FRANCO'
ebta22.loc[ebta22['Passageiro'].str.contains('GREG', case=False), 'Passageiro'] = 'GREGORIO MOREIRA FRANCO'
ebta22.loc[ebta22['Passageiro'].str.contains('BRUNO', case=False), 'Passageiro'] = 'BRUNO PINTO SIMAO'
ebta22.loc[ebta22['Passageiro'].str.contains('ANDERSON', case=False), 'Passageiro'] = 'ANDERSON ADEILSON DE OLIVEIRA'
ebta22.loc[(ebta22['Passageiro'].str.contains('GUSTAVO', case=False)) &
             (ebta22['Passageiro'].str.contains('ARAUJO', case=False)), 'Passageiro'] = 'GUSTAVO HENRIQUE DINIZ DE ARAUJO'
ebta22.loc[(ebta22['Passageiro'].str.contains('GUS', case=False)) &
             (ebta22['Passageiro'].str.contains('ARAUJO', case=False)), 'Passageiro'] = 'GUSTAVO HENRIQUE DINIZ DE ARAUJO'
ebta22.loc[ebta22['Passageiro'].str.contains('BORTON', case=False), 'Passageiro'] = 'LEO ADRIANO BORTON'
ebta22.loc[ebta22['Passageiro'].str.contains('GIULIANI', case=False), 'Passageiro'] = 'ROBERTO GIULIANI'
ebta22.loc[ebta22['Passageiro'].str.contains('FLAVIO', case=False), 'Passageiro'] = 'FLAVIO RIO BRANCO'
ebta22.loc[ebta22['Passageiro'].str.contains('ROBERTH', case=False), 'Passageiro'] = 'ROBERTH MACEDO SILVA'
ebta22.loc[ebta22['Passageiro'].str.contains('ANDRE', case=False), 'Passageiro'] = 'ANDRE RODRIGUES HORTA'
ebta22.loc[ebta22['Passageiro'].str.contains('MARCELO MACHADO', case=False), 'Passageiro'] = 'ROBSON MARCELO MACHADO SANTIAGO'
ebta22.loc[ebta22['Passageiro'].str.contains('COSTA FILHO', case=False), 'Passageiro'] = 'JOAO VICENTE BARRETO DA COSTA FILHO'
ebta22.loc[ebta22['Passageiro'].str.contains('SANTIAGO RICARDO', case=False), 'Passageiro'] = 'RICARDO VIEIRA SANTIAGO'
ebta22.loc[ebta22['Passageiro'].str.contains('HITOSI', case=False), 'Passageiro'] = 'HITOSI HASSEGAWA'
ebta22.loc[ebta22['Passageiro'].str.contains('SALDO', case=False), 'Passageiro'] = 'Não informado'
ebta22.loc[ebta22['Passageiro'].str.contains('BOLETO', case=False), 'Passageiro'] = 'Não informado'

print(len(ebta22))
ebta22.head()

# %% [markdown]
# ### 4.2. 2023 e 2024

# %%
# Primus
dfs_primus = []

for year in ['2023', '2024']:

        folder = fr"K:\GSAS\06 - Coordenação Gestão Compras, Logistica e Doctos\03 - Logística\01 - Faturamento\02 - Detalhamento de Faturas\Primus\{year}"
        to_append = []
        total_len = []

        for mes in os.listdir(folder):
                if os.path.isdir(folder+r'\\'+mes) and mes != str(datetime.today().month)+'.'+str(datetime.today().year):
                        ref = datetime.strptime((mes.split('.')[1] + '-' + mes.split('.')[0] + '-' + '01'), '%Y-%m-%d')
                        
                        f1 = os.path.join(folder, mes, "V1", "Consolidado V1.xlsx")
                        df1 = pd.read_excel(f1, sheet_name='Consolidado')
                        df1 = df1[(df1['FAIXA/ CARGO'] == 'DIR') | 
                                (df1['FAIXA/ CARGO'] == 'DIR EXT') | 
                                (df1['FAIXA/ CARGO'] == 'ADM EXT')].reset_index(drop=True)
                        df1['Referência'] = ref
                        
                        to_append.append(df1)
                        total_len.append(len(df1))

                        try:
                                f2 = os.path.join(folder, mes, "V2", "Consolidado V2.xlsx")
                                df2 = pd.read_excel(f2, sheet_name='Consolidado')
                                df2 = df2[(df2['FAIXA/ CARGO'] == 'DIR') | 
                                        (df2['FAIXA/ CARGO'] == 'DIR EXT') | 
                                        (df2['FAIXA/ CARGO'] == 'ADM EXT')].reset_index(drop=True)
                                df2['Referência'] = ref

                                to_append.append(df2)
                                total_len.append(len(df2))
                        except:
                                continue

                        print(mes, len(df1), len(df2))
        
        base_primus = pd.concat(to_append)
        base_primus.loc[(base_primus['FAIXA/ CARGO'].str.contains('EXT')) &
                (base_primus['TIPO'] == 'AEREO'), 'TIPO'] = 'AEREO INTERNACIONAL'

        base_primus['FORNECEDOR'] = base_primus['FORNECEDOR'] + " - " + base_primus['DESTINO']

        base_primus = base_primus[['DATA', 'PASSAGEIRO', 'FORNECEDOR', 'CC', 'TOTAL', 'TIPO', 'Referência']]
        base_primus.columns = ['Data de criação', 'Passageiro', 'Descrição', 'Número CC', 'Valor', 'Tipo despesa', 'Data despesa']
        base_primus.drop(columns=['Tipo despesa'], inplace=True)

        base_primus['Data de criação'] = pd.to_datetime(base_primus['Data de criação']).dt.strftime('%Y-%m-%d')
        base_primus['Data despesa'] = pd.to_datetime(base_primus['Data despesa']).dt.strftime('%Y-%m-%d')

        dfs_primus.append(base_primus)
        print(year, "ok")

primus23, primus24 = dfs_primus[0], dfs_primus[1]

#primus23 = primus23[~((primus23['Passageiro'] == 'LOPES BOFF/FELIPE') & (primus23['Número CC'] == 52000))]
#primus23.drop(columns=['Passageiro'], inplace=True)

nomes = {'RODRIGUES/PAULINO' : 'PAULINO RAMOS RODRIGUES',
'LOPES BOFF/FELIPE' : 'FELIPE LOPES BOFF',
'SIMÃO/BRUNO PINTO' : 'BRUNO PINTO SIMAO',
'Mariano da Fonseca/André Ranalli' : 'ANDRE RANALLI MARIANO DA FONSECA',
'MOREIRA FRANCO/GREGORIO' : 'GREGORIO MOREIRA FRANCO',
'RAMOS RODRIGUES/PAULINO' : 'PAULINO RAMOS RODRIGUES',
'FELIPE LOPES BOFF' : 'FELIPE LOPES BOFF',
'BRANCO/FLAVIO RIO' : 'FLAVIO RIO BRANCO FILHO',
'BOFF/FELIPE' : 'FELIPE LOPES BOFF',
'FORESTI RIBEIRO/VALERIA' : 'VALERIA DE ARAUJO FORESTI RIBEIRO',
'BRAGA REZENDE/VALCI MR' : 'VALCI BRAGA REZENDE',
'BRAGA REZENDE/VALCI' : 'VALCI BRAGA REZENDE',
'RAMOS RODRIGUES/PAULINO MR' : 'PAULINO RAMOS RODRIGUES',
'Lima Pereira Ruffo/Munir Amer' : 'MUNIR AMER LIMA PEREIRA RUFFO',
'SIMAO/BRUNO PINTO' : 'BRUNO PINTO SIMAO',
'HASSEGAWA/HITOSI' : 'HITOSI HASSEGAWA',
'ZIEGELMEYER/RICARDO' : 'RICARDO ZIEGELMEYER',
'DE OLIVEIRA SANTOS/MATEUS MORAES' : 'MATEUS MORAES DE OLIVEIRA SANTOS',
'FELIPE BOFF' : 'FELIPE LOPES BOFF',
'RICARDO ZIEGELMEYER' : 'RICARDO ZIEGELMEYER',
'CESAR ADRIANO FIGUEIREDO' : 'CESAR ADRIANO FIGUEIREDO',
'DINIZ DE ARAUJO/GUSTAVO HENRIQUE' : 'GUSTAVO HENRIQUE DINIZ DE ARAUJO',
'RELATORIO DE OS CANCELADAS' : 'Não informado',
'DE OLIVEIRA SOUZA/GILBERTO' : 'GILBERTO DE OLIVEIRA SOUZA',
'MAGALHAES/MARINA MRS' : 'MARINA DE AGUIAR MAGALHAES',
'MAGALHAES/MARINA' : 'MARINA DE AGUIAR MAGALHAES',
'LEONARDO CERQUEIRA' : 'LEONARDO MAURICIO CERQUEIRA',
'RODRIGO ARAUJO SIMOES' : 'RODRIGO DE ARAUJO SIMOES',
'OLIVEIRA/ANDERSON ADEILSON DE' : 'ANDERSON ADEILSON DE OLIVEIRA',
'ARAUJO/GUSTAVO HENRIQUE' : 'GUSTAVO HENRIQUE DINIZ DE ARAUJO',
'ALMEIDA/UELQUESNEURIAN RIBEIRO DE' : 'UELQUESNEURIAN RIBEIRO DE ALMEIDA',
'VIEIRA/ADILSON SANTOS' : 'ADILSON SANTOS VIEIRA',
'GUSTAVO HENRIQUE DE ARAUJO' : 'GUSTAVO HENRIQUE DINIZ DE ARAUJO',
'DINIZ DE ARAUJO/GUSTAVO HENRIQUE MR' : 'GUSTAVO HENRIQUE DINIZ DE ARAUJO',
'ARAUJO/GUSTAVO HENRIQUE MR' : 'GUSTAVO HENRIQUE DINIZ DE ARAUJO',
'VILLANI DE CASTRO/RAPHAEL MR' : 'RAPHAEL VILLANI DE CASTRO',
'KUBIAKI/LUCAS MR' : 'LUCAS LOPES KUBIAKI',
'HORTA/ANDRE' : 'ANDRE RODRIGUES HORTA',
'HORTA/LIZIANE' : 'LIZIANE HORTA',
'LUCAS KUBIAKI' : 'LUCAS LOPES KUBIAKI',
'ARAUJO/PAULO HENRIQUE' : 'PAULO HENRIQUE BRANT DE ARAUJO',
'LOPES KUBIAKI/LUCAS' : 'LUCAS LOPES KUBIAKI',
'PINTO SIMAO/BRUNO' : 'BRUNO PINTO SIMAO',
'COLLODORO/RENAN' : 'RENAN MOREIRA COLLODORO',
'DINIZ DE ARAUJO/GUSTAVO MR' : 'GUSTAVO HENRIQUE DINIZ DE ARAUJO',
'DINIZ DE ARAUJO/GUSTAVO' : 'GUSTAVO HENRIQUE DINIZ DE ARAUJO',
'RANALLI MARIANO DA FONSECA/ANDRE' : 'ANDRE RANALLI MARIANO DA FONSECA',
'ALVARENGA/GUSTAVO' : 'GUSTAVO DINIZ ALVARENGA',
'RAPHAEL CASTRO' : 'RAPHAEL VILLANI DE CASTRO',
'Kubiaki/Lucas' : 'LUCAS LOPES KUBIAKI',
'RAMOS RODRIGUES  /PAULINO' : 'PAULINO RAMOS RODRIGUES',
'GUSTAVO ALVARENGA' : 'GUSTAVO DINIZ ALVARENGA',
'BRUNO SIMAO' : 'BRUNO PINTO SIMAO',
'RENAN COLLODORO' : 'RENAN MOREIRA COLLODORO'}

primus23.replace(nomes, inplace=True)
primus24.replace(nomes, inplace=True)

print(len(primus23), len(primus24))

# %%
# EBTA
path = project_folder + r"\Faturamento\EBTA 23"

to_append = []
total_len = []

for file_name in os.listdir(path):
    
    f = os.path.join(path, file_name)
    df = pd.read_excel(f, sheet_name='Dados')
    df.columns = df.columns.str.lower()
    df = df[(df['faixa']=='DIR') | (df['faixa']=='DIR EXT') | (df['faixa']=='ADM EXT')]
    df['Referência'] = file_name.split('.')[0]
    
    to_append.append(df)
    total_len.append(len(df))
    print(file_name, len(df))
    
ebta23 = pd.concat(to_append).reset_index(drop=True)
ebta23.columns = ['Data de criação', 'Descrição', 'Número CC', 'Valor', 'faixa', 'Passageiro', 'Data despesa']
ebta23['Tipo despesa'] = np.where(ebta23['faixa'].str.contains('EXT'), 'AEREO INTERNACIONAL', 'AEREO')
ebta23.loc[(ebta23['Descrição'].fillna("").str.contains('TLV')), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
ebta23.drop(columns=['faixa'], inplace=True)

ebta23.loc[ebta23['Passageiro'].str.contains('PAULINO', case=False), 'Passageiro'] = 'PAULINO RAMOS RODRIGUES'
ebta23.loc[ebta23['Passageiro'].str.contains('VALCI', case=False), 'Passageiro'] = 'VALCI BRAGA REZENDE'
ebta23.loc[(ebta23['Passageiro'].str.contains('GUSTAVO', case=False)) &
           (ebta23['Passageiro'].str.contains('ARAUJO', case=False)), 'Passageiro'] = 'GUSTAVO HENRIQUE DINIZ DE ARAUJO'
ebta23.loc[ebta23['Passageiro'].str.contains('MARINA', case=False), 'Passageiro'] = 'MARINA DE AGUIAR MAGALHAES'
ebta23.loc[ebta23['Passageiro'].str.contains('KUBIAKI', case=False), 'Passageiro'] = 'LUCAS LOPES KUBIAKI'
ebta23.loc[ebta23['Passageiro'].str.contains('BOFF', case=False), 'Passageiro'] = 'FELIPE LOPES BOFF'
ebta23.loc[ebta23['Passageiro'].str.contains('SALDO', case=False), 'Passageiro'] = 'Não informado'
ebta23.loc[ebta23['Passageiro'].str.contains('BOLETO', case=False), 'Passageiro'] = 'Não informado'

print(len(ebta23))
ebta23.head()

# %% [markdown]
# ### 4.3. Primus e EBTA

# %%
primus = pd.concat([primus22, primus23, primus24]).reset_index(drop=True)

primus['Data de criação'] = pd.to_datetime(primus['Data de criação'], format = '%Y-%m-%d')
primus['Data despesa'] = pd.to_datetime(primus['Data despesa'], format = '%Y-%m-%d')

primus['Descrição'] = "Fornecedor: " + primus['Descrição']
primus.rename(columns={'Passageiro':'Colaborador'}, inplace=True)
primus['Origem'] = 'Primus'

print(len(primus22), len(primus23), len(primus24), len(primus))
primus.head()

# %%
ebta = pd.concat([ebta22, ebta23]).reset_index(drop=True)
ebta = ebta[ebta['Número CC'].notna()]
ebta['Data de criação'] = pd.to_datetime(ebta['Data de criação'], format = '%d/%m/%Y')
ebta['Data despesa'] = pd.to_datetime(ebta['Data despesa'], format = '%Y-%m-%d')
ebta['Descrição'] = "Trechos: " + ebta['Descrição']
ebta.rename(columns={'Passageiro':'Colaborador'}, inplace=True)
ebta['Origem'] = 'EBTA'

print(len(ebta22), len(ebta23), len(ebta))
ebta

# %% [markdown]
# ### 4.4. Diretoria

# %%
diretoria = pd.concat([primus, ebta[ebta['Número CC']!='                         ']])
diretoria['Número CC'] = diretoria['Número CC'].astype('int64')
diretoria['Número CC'].replace({13629:13269}, inplace=True)
diretoria.loc[(diretoria['Colaborador'] == 'FELIPE LOPES BOFF') & (diretoria['Número CC'] == 52000), 'Número CC'] = 13199
diretoria = diretoria.merge(segmentos[['Número CC', 'Centro de custo',
                                       'Nº CC Subordinador', 'CC Subordinador', 'Segmento']],
                            on='Número CC', how='left')

diretoria['id'] = [i for i in range(len(diretoria))]
diretoria['Valor total'] = diretoria['Valor']
#diretoria['Descrição'] = "Contábil - Diretoria"
#diretoria = diretoria.merge(resp, on='Número CC', how='left')
diretoria['Projeto'] = "10 - Outros"
diretoria['Count id'], diretoria['Rank id'] = 1, 1
diretoria = diretoria.merge(rh_cadastro[['Colaborador', 'Cargo', 'CC Lotação']], on='Colaborador', how='left')

print(len(diretoria))
diretoria.head()

# %%
path = project_folder + r"\Despesas_Diretoria.xlsx"

with pd.ExcelWriter(path) as writer:
    diretoria.to_excel(writer, sheet_name='Despesas Diretoria', index=False)

# %%
filtro_dir = base_seg[base_seg.Segmento!='DIRETORIA E PRESIDÊNCIA'].reset_index(drop=True)
base_dir = pd.concat([filtro_dir, diretoria]).reset_index(drop=True)

# base_dir['Tipo despesa'].replace({
#     'AEREO': 'Aéreo', 'HOTEL': 'Hotel', 'CARRO':'Carro', 'PLANTÃO':'Plantão', 'SEGURO DE VIAGEM':'Seguro',
#     'SEGURO':'Seguro', 'SEGURO VIAGEM':'Seguro', 'GESTÃO':'Gestão', 'TAXA':'Taxa'}, inplace=True)

print(len(base_dir))
base_dir.head()

# %%
base_dir['Data despesa'] = base_dir['Data despesa'] + pd.Timedelta(days=1)
base_dir.rename(columns={'Data despesa': 'Data despesa antiga'}, inplace=True)

# %% [markdown]
# ### 4.5. Ajuste Data

# %%
path = project_folder + r"\Bases mensais\ajuste_data.xlsx"

ajuste_data = pd.read_excel(path, header=2, usecols="C,F")
ajuste_data.columns = ['Viagem', 'Data despesa']
ajuste_data['id'] = ajuste_data.Viagem.str.split("-", n=1, expand=True)[0]

ajuste_data = ajuste_data[pd.to_numeric(ajuste_data['id'], errors='coerce').notnull()].reset_index(drop=True)
ajuste_data['id'] = ajuste_data['id'].astype('int64')

print(len(ajuste_data))
ajuste_data.head()

# %%
base_dta = base_dir.merge(ajuste_data[['Data despesa', 'id']], on='id', how='left')
base_dta['Data despesa'].fillna(base_dta['Data despesa antiga'], inplace=True)
base_dta.drop(columns=['Data despesa antiga'], inplace=True)

print(len(base_dta))
base_dta.head()

# %% [markdown]
# ## 5. Orçamentos

# %%
### 2022 ###
path = project_folder + r"\Bases mensais\Orçamento 2022.xlsx"

orcamento22 = pd.read_excel(path, sheet_name='Orçamento')
print(len(orcamento22))
orcamento22.head()

# %%
### 2023 ###
path = project_folder + r"\Bases mensais\Orçamento 2023.xlsx"

orcamento23 = pd.read_excel(path, sheet_name='Orçamento')
print(len(orcamento23))
orcamento23.head()

# %%
### 2024 ###
path = project_folder + r"\Bases mensais\Orçamento 2024.xlsx"

orcamento24 = pd.read_excel(path, sheet_name='Orçamento')
print(len(orcamento24))
orcamento24.head()

# %%
orcamento = pd.concat([orcamento22, orcamento23, orcamento24])
orcamento['Key'] = (orcamento['Centro de custo'].astype('str') + '-' +
                    orcamento['Mês'].dt.year.astype('str') + '-' +
                    orcamento['Mês'].dt.month.astype('str'))

orcamento.rename(columns={'Centro de custo':'Número CC'}, inplace=True)

print(len(orcamento))
orcamento.head()

# %% [markdown]
# ### 5.1. Agrupar por Key

# %%
orcamento_mes = orcamento.groupby('Key', as_index=False).agg({'Valor Orçado':'sum',
                                                              'Número CC':'first',
                                                              'Mês':'first',
                                                              'Grupo Orçamentário':'first'})

print(len(orcamento_mes))
orcamento_mes.head()

# %%
base_dta['Key'] = (base_dta['Número CC'].astype('str') + '-' +
                   base_dta['Data despesa'].dt.year.astype('str') + '-' +
                   base_dta['Data despesa'].dt.month.astype('str'))

base_orc = base_dta.merge(orcamento_mes[['Key','Valor Orçado']], on='Key', how='left')
base_orc['Valor Orçado'].fillna(0, inplace=True)

# base_orc['Count Key'] = base_orc.groupby('Key')['Key'].transform('count')
# base_orc['Orçado Ajustado'] = base_orc['Valor Orçado'] / base_orc['Count Key']
#base_orc.drop(columns=['Valor Orçado', 'Count Key'], inplace=True)
#base_orc.rename(columns={'Orçado Ajustado': 'Valor Orçado'}, inplace=True)

base_orc['Mês'] = pd.to_datetime(base_dta['Data despesa'].dt.year.astype('str') + '-' +
                                 base_dta['Data despesa'].dt.month.astype('str') + '-' + '01')

print(len(base_orc))
base_orc.head()

# %% [markdown]
# ### 5.2. Compor base de orçamento

# %%
orcamento_mes_full = orcamento_mes.merge(base_orc.groupby('Key', as_index=False).agg(
                                         {'Valor':'sum', 'Número CC':'first', 'Mês':'first'}),
                                         on='Key', how='outer')
base_orc.drop(columns=['Mês'], inplace=True)

orcamento_mes_full['Número CC_x'].fillna(orcamento_mes_full['Número CC_y'],inplace=True)
orcamento_mes_full['Mês_x'].fillna(orcamento_mes_full['Mês_y'],inplace=True)
orcamento_mes_full.drop(columns=['Número CC_y','Mês_y'], inplace=True)

orcamento_mes_full.rename(columns={'Valor':'Valor Realizado', 'Número CC_x':'Número CC', 'Mês_x':'Mês'}, inplace=True)
orcamento_mes_full['Número CC'] = orcamento_mes_full['Número CC'].astype('int64')

orcamento_mes_full = orcamento_mes_full.merge(segmentos[['Número CC', 'Segmento', 'UF']],
                                              on='Número CC', how='left')

orcamento_mes_full['Valor Realizado'].fillna(0, inplace=True)
orcamento_mes_full['Valor Orçado'].fillna(0, inplace=True)
orcamento_mes_full['UF'].fillna('', inplace=True)

print(len(orcamento_mes_full))
orcamento_mes_full.head()

# %% [markdown]
# ### 5.3 Orçamento + RH

# %%
rh_merge = rh_cadastro[['Colaborador', 'Centro de Custo']]
rh_merge.columns = ['Colaborador', 'Número CC']

# %%
# rh_cadastro 3438
# orcamento_mes_full 4400
# 11172

orcamento_rh = rh_merge.merge(orcamento_mes_full, on='Número CC', how='outer')

orcamento_rh['Count Key'] = orcamento_rh.groupby('Key')['Key'].transform('count')
orcamento_rh['Orçado Ajustado'] = orcamento_rh['Valor Orçado'] / orcamento_rh['Count Key']
orcamento_rh['Realizado Ajustado'] = orcamento_rh['Valor Realizado'] / orcamento_rh['Count Key']

orcamento_rh

# %%
### EM TESTE

group = base_orc.groupby('Key').agg({'Número CC':'first', 'Centro de custo':'first', 'Valor Orçado':'first'}).reset_index()

group2 = group.groupby('Número CC').agg({'Centro de custo':'first', 'Valor Orçado':'sum'})

group[group['Número CC']==13150].sort_values('Valor Orçado', ascending=False).reset_index(drop=True)

# %% [markdown]
# ## 6. Comparativo aéreo

# %%
path = project_folder + r"\Bases mensais\Violações\Comparativo aéreo.xlsx"
comparativo_raw = pd.read_excel(path, header=1)

# Ler arquivo
comparativo = pd.read_excel(path, usecols="A,H,P,X,Y", header=1)
comparativo.rename(columns={'Relatório':'id', 'Justificativa':'Justificativa maior preço', 
                            'Valor':'Valor Selecionado', 'Valor.1':'Melhor preço', 'Diferença':'Diferença preço aéreo'},
                   inplace=True)
comparativo['Justificativa maior preço'].fillna('', inplace=True)

# SE PRECISAR PARA A API, SE NÃO TIVER O CAMPO DE DIFERENÇA PREÇO AÉREO
#comparativo['Diferença preço aéreo'] = comparativo['Valor Selecionado'] - comparativo['Melhor preço']

# Padronizar id
comparativo.dropna(subset=['id'], inplace=True)
comparativo = comparativo[pd.to_numeric(comparativo['id'], errors='coerce').notnull()]
comparativo['id'] = comparativo['id'].astype('int64')

comparativo['Valor Selecionado'] = comparativo['Valor Selecionado'].astype('float64')
comparativo['Melhor preço'] = comparativo['Melhor preço'].astype('float64')
comparativo['Diferença preço aéreo'] = comparativo['Diferença preço aéreo'].astype('float64')

# Remover canceladas
comparativo = comparativo[comparativo['Valor Selecionado'] != 0]

# Remover duplicatas (forçar apenas 1 ocorrência do id juntando as colunas de string)
comparativo.drop_duplicates(inplace=True)
comparativo = comparativo.groupby(['id'], as_index = False).agg({'Justificativa maior preço': ' / '.join,
                                                                'Valor Selecionado':'sum', 
                                                                'Melhor preço':'sum',
                                                                'Diferença preço aéreo':'sum'})
print(len(comparativo))
comparativo.head()

# %%
base_comp = base_orc.merge(comparativo, on='id', how='left')
print(len(base_comp))
base_comp.head()

# %% [markdown]
# ## 7. Violação antecedência

# %%
path = project_folder + r"\Bases mensais\Violações\Antecedência.xlsx"
antecedencia_raw = pd.read_excel(path, header=1)

antecedencia = pd.read_excel(path, usecols="A,F,I,J", header=1)
antecedencia.rename(
    columns={'Código':'id', 'Dias de \nantecedência':'Dias de antecedência',
             'Justificativa':'Justificativa antecedência', 'Aprovador':'Aprovador antecedência'}, inplace=True)

antecedencia.dropna(subset=['id'], inplace=True)
antecedencia['id'] = antecedencia['id'].astype('int64')

print(len(antecedencia))
antecedencia.head()

# %%
base_ant = base_comp.merge(antecedencia, on='id', how='left')
print(len(base_ant))
base_ant.head()

# %% [markdown]
# ## 8. Aéreo

# %%
path = project_folder + r"\Bases mensais\Violações\Aéreo.xlsx"
aereo_raw = pd.read_excel(path, header=1)

# Ler arquivo
aereo = pd.read_excel(path, usecols="A,F,N,R,U", header=1)
aereo.rename(
    columns={'Solicitação':'id', 'Política violada':'Política aéreo violada',
             'Justificativa':'Justificativa aéreo', 'Aprovador':'Aprovador aéreo'}, inplace=True)
aereo.fillna('', inplace=True)

# Padronizar id
aereo.dropna(subset=['id'], inplace=True)
aereo = aereo[pd.to_numeric(aereo['id'], errors='coerce').notnull()]
aereo['id'] = aereo['id'].astype('int64')

# Remover canceladas
aereo = aereo[aereo.Status != 'Cancelada']
aereo.drop(columns=['Status'], inplace=True)

# Remover duplicatas (forçar apenas 1 ocorrência do id juntando as colunas de string)
aereo.drop_duplicates(inplace=True)
aereo = aereo.groupby(['id'], as_index = False).agg({'Política aéreo violada': ' / '.join,
                                                     'Justificativa aéreo': ' / '.join,
                                                     'Aprovador aéreo': ' / '.join})

print(len(aereo))
aereo.head()

# %%
base_aer = base_ant.merge(aereo, on='id', how='left')
print(len(base_aer))
base_aer.head()

# %% [markdown]
# ## 9. Hotel

# %%
path = project_folder + r"\Bases mensais\Violações\Hotel.xlsx"
hotel_raw = pd.read_excel(path, header=1)

# Ler arquivo
hotel = pd.read_excel(path, usecols="A,E,K,O,Q", header=1)
hotel.rename(
    columns={'Solicitação':'id', 'Política violada':'Política hotel violada',
             'Justificativa':'Justificativa hotel', 'Aprovador':'Aprovador hotel'}, inplace=True)
hotel.fillna('', inplace=True)

# Padronizar id
hotel.dropna(subset=['id'], inplace=True)
hotel = hotel[pd.to_numeric(hotel['id'], errors='coerce').notnull()]
hotel['id'] = hotel['id'].astype('int64')

# Remover canceladas
hotel = hotel[hotel.Status != 'Cancelada']
hotel.drop(columns=['Status'], inplace=True)

# Remover duplicatas (forçar apenas 1 ocorrência do id juntando as colunas de string)
hotel.drop_duplicates(inplace=True)
hotel = hotel.groupby(['id'], as_index = False).agg({'Política hotel violada': ' / '.join,
                                                     'Justificativa hotel': ' / '.join,
                                                     'Aprovador hotel': ' / '.join})
print(len(hotel))
hotel.head()

# %%
base_hot = base_aer.merge(hotel, on='id', how='left')
print(len(base_hot))
base_hot.head()

# %% [markdown]
# ## 10. Carro

# %%
path = project_folder + r"\Bases mensais\Violações\Carro.xlsx"
carro_raw = pd.read_excel(path, header=1)

# Ler arquivo
carro = pd.read_excel(path, usecols="A,D,L,O,R", header=1)
carro.rename(
    columns={'Solicitação':'id', 'Política violada':'Política carro violada',
             'Justificativa':'Justificativa carro', 'Aprovador':'Aprovador carro'}, inplace=True)
carro.fillna('', inplace=True)

# Padronizar id
carro.dropna(subset=['id'], inplace=True)
carro = carro[pd.to_numeric(carro['id'], errors='coerce').notnull()]
carro['id'] = carro['id'].astype('int64')

# Remover canceladas
carro = carro[carro.Status != 'Cancelada']
carro.drop(columns=['Status'], inplace=True)

# Remover duplicatas (forçar apenas 1 ocorrência do id juntando as colunas de string)
carro.drop_duplicates(inplace=True)
carro = carro.groupby(['id'], as_index = False).agg({'Política carro violada': ' / '.join,
                                                     'Justificativa carro': ' / '.join,
                                                     'Aprovador carro': ' / '.join})
print(len(carro))
carro.head()

# %%
base_car = base_hot.merge(carro, on='id', how='left')
print(len(base_car))
base_car.head()

# %% [markdown]
# ## 11. Limite diário

# %%
path = project_folder + r"\Bases mensais\Violações\Limite diário.xlsx"
lim_diario_raw = pd.read_excel(path)

# Ler arquivo
lim_diario = pd.read_excel(path, usecols="A,C,F,M,N,V,W")
lim_diario.rename(
    columns={'#':'id', 'Política violada':'Política limite diário violada', 'Situação':'Status',
             'Justificativa':'Justificativa limite diário', 'Aprovador':'Aprovador limite diário'}, inplace=True)
lim_diario.fillna('', inplace=True)
lim_diario['Política limite diário violada'] = lim_diario['Política limite diário violada'].apply(unidecode)

# Padronizar id
lim_diario.dropna(subset=['id'], inplace=True)
lim_diario = lim_diario[pd.to_numeric(lim_diario['id'], errors='coerce').notnull()]
lim_diario['id'] = lim_diario['id'].astype('int64')

# Remover canceladas
lim_diario = lim_diario[lim_diario.Status != 'Cancelada'].reset_index(drop=True)
lim_diario.drop(columns=['Status'], inplace=True)

# Remover duplicatas (forçar apenas 1 ocorrência do id juntando as colunas de string)
#lim_diario.drop_duplicates(inplace=True)

##lim_diario = lim_diario.groupby(['id'], as_index = False).agg({'Política limite diário violada': ' / '.join,
##                                                     'Justificativa limite diário': ' / '.join,
##                                                     'Aprovador limite diário': ' / '.join})

print(len(lim_diario))
lim_diario.head()

# %%
# Extrair limites diários e valores excedidos da política violada
regex = lim_diario['Política limite diário violada'].str.extractall('R\$\ (0|[1-9][0-9]{0,2})(.\d{3})*(\,\d{1,2})?')
regex['valores'] = regex[0]+regex[2]
regex.drop(columns=[0,1,2], inplace=True)
regex = regex.unstack(level=-1).reset_index(drop=True)
regex.columns = regex.columns.to_flat_index()
regex.rename({('valores', 0):'Limite diário', ('valores', 1):'Valor excedido'}, axis='columns', inplace=True)

# Transformar valores em número
regex['Limite diário'] = regex['Limite diário'].str.replace('.','')
regex['Limite diário'] = regex['Limite diário'].str.replace(',','.')
regex['Valor excedido'] = regex['Valor excedido'].str.replace('.','')
regex['Valor excedido'] = regex['Valor excedido'].str.replace(',','.')

regex['Limite diário'] = regex['Limite diário'].astype('float64')
regex['Valor excedido'] = regex['Valor excedido'].astype('float64')

print(len(regex))
regex.head()

# %%
# Acrescentar colunas à base de limite diário
lim_diario_merged = lim_diario.merge(regex, left_index=True, right_index=True, how='left')

# Agrupar por id, tipo (hospedagem/alimentação) e data --> O valor excedido é igual para os mesmos id, tipos e datas
lim_diario_merged['Key'] = (lim_diario_merged['id'].astype('str') + '-' +
                            lim_diario_merged['Tipo despesa'].astype('str') + '-' +
                            lim_diario_merged['Data despesa'].astype('str'))

lim_diario_merged = lim_diario_merged.groupby('Key', as_index=False).agg({'id':'first',
                                                                          'Aprovador limite diário':'first',
                                                                          'Tipo despesa':'first',
                                                                          'Data despesa':'first',
                                                                          'Política limite diário violada':'first',
                                                                          'Justificativa limite diário':' / '.join,
                                                                          'Valor excedido':'first'})

# Somar valor excedido total por viagem
lim_diario_merged = lim_diario_merged.groupby('id', as_index=False).agg({'Aprovador limite diário':'first',
                                                                         'Política limite diário violada':' / '.join,
                                                                         'Justificativa limite diário':' / '.join,
                                                                         'Valor excedido':'sum'})
print(len(lim_diario_merged))
lim_diario_merged.head()

# %%
base_lim = base_car.merge(lim_diario_merged, on='id', how='left')
print(len(base_lim))
base_lim.head()

# %% [markdown]
# ## 12. Preço e Prazo

# %%
conditions_preco = [(base_lim['Política aéreo violada'].str.contains('preço', na=False)),
              (base_lim['Política hotel violada'].str.contains('preço', na=False)),
              (base_lim['Política carro violada'].str.contains('preço', na=False)),
              (base_lim['Política limite diário violada'].notna())
             ]
values_preco = [1,1,1,1]
base_lim['Fora do preço'] = np.select(conditions_preco, values_preco)

conditions_prazo = [(base_lim['Política aéreo violada'].str.contains('prazo', na=False)),
              (base_lim['Política hotel violada'].str.contains('prazo', na=False)),
              (base_lim['Política carro violada'].str.contains('prazo', na=False)),
              (base_lim['Dias de antecedência'].notna())
             ]
values_prazo = [1,1,1,1]
base_lim['Fora do prazo'] = np.select(conditions_prazo, values_prazo)

# %% [markdown]
# ## 13. Ajustes finais

# %% [markdown]
# 13. Hierarquia INSS
# 
# path = project_folder + r"\Bases mensais\hierarquia_inss.sql"
# 
# conn = pyodbc.connect('Driver={SQL Server};'
#                       'Server=SWDVMA1383;'
#                       'Database=DBZB099D;'
#                       'Trusted_Connection=yes;')
# 
# with open(path, 'r') as file:
#     query = file.read()
# 
# hier_inss = pd.read_sql(query, conn)
# hier_inss.columns = ['Número CC', 'Centro de custo', 'Gerência', 'Regional', 'Superintendência']
# 
# hier_inss = hier_inss.merge(segmentos[['Número CC', 'Segmento']], on='Número CC', how='left')
# 
# print(len(hier_inss))
# hier_inss.head()

# %% [markdown]
# ## 14. Exportar

# %%
path = project_folder + r"\Gestão de Viagens_v6.xlsx"

with pd.ExcelWriter(path) as writer:
    base_lim.to_excel(writer, sheet_name='Despesas Paytrack', index=False)
    orcamento_mes_full.to_excel(writer, sheet_name='Orçamento', index=False)
    orcamento_rh.to_excel(writer, sheet_name='Orçamento RH', index=False)
    #hier_inss.to_excel(writer, sheet_name='Hierarquias INSS', index=False)
    
    rh_cadastro.to_excel(writer, sheet_name='RH Cadastro', index=False)
    comparativo_raw.to_excel(writer, sheet_name='Comparativo preços', index=False)
    antecedencia_raw.to_excel(writer, sheet_name='Antecedência', index=False)
    aereo_raw.to_excel(writer, sheet_name='Aéreo', index=False)
    hotel_raw.to_excel(writer, sheet_name='Hotel', index=False)
    carro_raw.to_excel(writer, sheet_name='Carro', index=False)
    lim_diario_raw.to_excel(writer, sheet_name='Limite diário', index=False)

# %% [markdown]
# ## 15. Copiar arquivos para rede

# %%
file_to_copy = r"K:\GSAS\00 - Gerência\001 - Atividades e projetos da Gerência\11-2022 - VIAGENS - Relatório Gestão Mensal\Gestão de Viagens_v6.xlsx"
destination_directory = r"T:\SIC\Seg Patrimonial\Manutenções\Viagens\Gestão de Viagens_v6.xlsx"
shutil.copy(file_to_copy, destination_directory)

file_to_copy2 = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Data Lake\hierarquia_cc.xlsx"
destination_directory2 = r"T:\SIC\Seg Patrimonial\Manutenções\Viagens\hierarquia_cc.xlsx"
shutil.copy(file_to_copy2, destination_directory2)

file_to_copy3 = r"K:\GSAS\12. Núcleo de Informações\6.05-MATHEUS\Data Lake\segmentos.xlsx"
destination_directory3 = r"T:\SIC\Seg Patrimonial\Manutenções\Viagens\segmentos.xlsx"
shutil.copy(file_to_copy3, destination_directory3)

pyautogui.alert('O código foi finalizado. Você já pode utilizar o computador!')