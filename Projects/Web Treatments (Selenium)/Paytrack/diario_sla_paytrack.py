# %%
# ------------- IMPORTANDO BIBLIOTECAS -------------
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

import sqlalchemy
from sqlalchemy.engine import URL

from funcoes_paytrack import login_paytrack, baixar_relatorio, buscar_historico_download, mover_para_pasta

import warnings
warnings.filterwarnings('ignore')

#print(selenium.__version__)



# %%
project_folder = r"K:\GSAS\00 - Gerência\001 - Atividades e projetos da Gerência\11-2022 - VIAGENS - Relatório Gestão Mensal\SLA Datas"

matricula = 'b042786'
#matricula = input('Informe sua mattricula: \n')
download_folder = f"C:/Users/{matricula}/Downloads"

##### GAMBIARRA DEMANDAS ENGENHARIA #####
file_to_copy = r"K:\GSAS\10 - Gestão de Relacionamento com Fornecedores\08.Conservação e Limpeza\Manutenções\Data.xlsx"
destination_directory = r"T:\SIC\Seg Patrimonial\Manutenções\Data.xlsx"
shutil.copy(file_to_copy, destination_directory)

file_to_copy2 = r"K:\GSAS\10 - Gestão de Relacionamento com Fornecedores\08.Conservação e Limpeza\Manutenções\Lista de OS.xlsx"
destination_directory2 = r"T:\SIC\Seg Patrimonial\Manutenções\Full Connection\Lista de OS.xlsx"
shutil.copy(file_to_copy2, destination_directory2)

file_to_copy4 = r"K:\GSAS\10 - Gestão de Relacionamento com Fornecedores\08.Conservação e Limpeza\Manutenções\Controlão Full Connection.xlsx"
destination_directory4 = r"T:\SIC\Seg Patrimonial\Manutenções\Full Connection\Controlão Full Connection.xlsx"
shutil.copy(file_to_copy4, destination_directory4)



# %%
# ------------- Definindo Funções -------------
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
# ------------- INICIANDO NAVEGADOR -------------
# Parâmetros
yesterday = date.today() - timedelta(days=1)
yesterday = datetime.strftime(yesterday, "%d/%m/%Y")

# Inicializar webdriver
#driver = webdriver.Chrome()
#driver.maximize_window()
#wait = WebDriverWait(driver, 240)
#
#
#
##%%
## ------------- INICIANDO DOWNLOADS -------------
## Login
pw = "@Mercantil123"
#login_paytrack( "matheus.oandrade@mercantil.com.br", pw )
#
########### RDV11 ##########
#### Parâmetros
#relatorio = "RDV11"
#data_inicial = "01/01/2023"
#data_final = yesterday
#id_data_inicial = "edtData_dt_inicial_despesa"
#id_data_final = "edtData_dt_final_despesa"
#download_path = download_folder + r"\RDV11 - Despesas (Exportação).xlsx"
#destination_path = project_folder + r"\RDV11 - Despesas (Exportação).xlsx"
#
#baixar_relatorio(
#    relatorio,
#    data_inicial,
#    data_final,
#    id_data_inicial,
#    id_data_final
#)
#
#mover_para_pasta(
#    relatorio,
#    download_path,
#    destination_path
#)
#
########### SOL03 ##########
#### Parâmetros
#relatorio = "SOL03"
#data_inicial = "01/01/2023"
#data_final = yesterday
#id_data_inicial = "edtData_data_inicial"
#id_data_final = "edtData_data_final"
#download_path = download_folder + r"\SOL03 - Roteiro do relatório.xlsx"
#destination_path = project_folder + r"\SOL03 - Roteiro do relatório.xlsx"
#
#baixar_relatorio(
#    relatorio,
#    data_inicial,
#    data_final,
#    id_data_inicial,
#    id_data_final
#)
#
#mover_para_pasta(
#    relatorio,
#    download_path,
#    destination_path
#)
#
########### GER27 ##########
#### Parâmetros
#relatorio = "GER27"
#data_inicial = "01/01/2023"
#data_final = yesterday
#id_data_inicial = "edtData_data_inicial_viagem"
#id_data_final = "edtData_data_final_viagem"
#download_path = download_folder + r"\GER27 - Relatório de reprovações.xlsx"
#destination_path = project_folder + r"\GER27 - Relatório de reprovações.xlsx"
#
#baixar_relatorio(
#    relatorio,
#    data_inicial,
#    data_final,
#    id_data_inicial,
#    id_data_final
#)
#
#mover_para_pasta(
#    relatorio,
#    download_path,
#    destination_path
#)
#
########### CAD06 ##########
#### Parâmetros
#relatorio = "CAD06"
#data_inicial = "01/01/2023"
#data_final = yesterday
#id_data_inicial = "edtData_dataInicial"
#id_data_final = "edtData_dataFinal"
#download_path = download_folder + r"\CAD06 - Prestação de contas aprovadas.xlsx"
#destination_path = project_folder + r"\CAD06 - Prestação de contas aprovadas.xlsx"
#
#baixar_relatorio(
#    relatorio,
#    data_inicial,
#    data_final,
#    id_data_inicial,
#    id_data_final
#)
#
#buscar_historico_download(180)
#
#mover_para_pasta(
#    relatorio,
#    download_path,
#    destination_path
#)
#
#driver.quit()
#


# %%
# ------------- TRAT 1 -------------
path = project_folder + r"\RDV11 - Despesas (Exportação).xlsx"

rdv11 = pd.read_excel(path, converters={'#':int}, usecols='A,C,D,F,I,O')
rdv11.columns = ['id', 'Status', 'Data_prestação_contas', 'Data_conferencia', 'Colaborador', 'Data_despesa']

rdv11['Colaborador'] = rdv11['Colaborador'].apply(unidecode)

print(len(rdv11))
rdv11.head()



# %%
# ------------- TRAT 2 -------------
df1 = rdv11.groupby('id').agg({'Data_prestação_contas':'first'}).reset_index().sort_values('id')
df2 = rdv11.groupby('id').agg({'Data_prestação_contas':'last'}).reset_index().sort_values('id')
validate = df1 == df2

print(df1.isna().sum())
print(df2.isna().sum())
validate[validate['Data_prestação_contas']==False]



# %%
# ------------- TRAT 3 -------------
df1 = rdv11.groupby('id').agg({'Data_conferencia':'first'}).reset_index().sort_values('id')
df2 = rdv11.groupby('id').agg({'Data_conferencia':'last'}).reset_index().sort_values('id')
validate = df1 == df2

print(df1.isna().sum())
print(df2.isna().sum())
validate[validate['Data_conferencia']==False]



# %%
# ------------- TRAT 4 -------------
df_prest = rdv11.groupby(['id', 'Colaborador', 'Status']).agg(
    {'Data_prestação_contas':'last', 'Data_conferencia':'last', 'Data_despesa':['min','max']}
    ).reset_index().sort_values('id')
df_prest.columns = ['id', 'Colaborador', 'Status', 'Data_prestação_contas', 'Data_conferencia', 'Data_inicio', 'Data_final']

print(len(df_prest))
df_prest.head()



# %% 
# ------------- IGNORAR -------------
# path = project_folder + r"\RDV14 - Prestação de contas finalizadas (Exportação).xlsx"
# 
# rdv14 = pd.read_excel(path, converters={'#':int}, usecols='A,E')
# rdv14.columns = ['id', 'Data_inicio_rdv14']
# 
# print(len(rdv14))
# rdv14.head()

# %% 
# df1 = rdv14.groupby('id').agg({'Data_inicio_rdv14':'first'}).reset_index().sort_values('id')
# df2 = rdv14.groupby('id').agg({'Data_inicio_rdv14':'last'}).reset_index().sort_values('id')
# validate = df1 == df2
# 
# print(df1.isna().sum())
# print(df2.isna().sum())
# validate[validate['Data_inicio_rdv14']==False]

# %% 
# df_inicio = rdv14.groupby('id').agg({'Data_inicio_rdv14':'last'}).reset_index().sort_values('id')
# df_inicio = df_prest.merge(df_inicio, on='id', how='left')
# 
# print(len(df_inicio))
# df_inicio.head()

# %% 
# ## 2. Data fim



# %%
# ------------- TRAT 5 -------------
path = project_folder + r"\SOL03 - Roteiro do relatório.xlsx"
sol03 = pd.read_excel(path, header=5, converters={'Unnamed: 9':int}, usecols='J,P,U')
sol03.columns = ['id', 'Data_inicio_roteiro', 'Data_fim_roteiro']

sol03.dropna(subset='id', inplace=True)
sol03['Data_inicio_roteiro'] = pd.to_datetime(sol03['Data_inicio_roteiro'], format='%d/%m/%y %H:%M')
sol03['Data_fim_roteiro'] = pd.to_datetime(sol03['Data_fim_roteiro'], format='%d/%m/%y %H:%M')

print(len(sol03))
sol03.head()



# %%
# ------------- TRAT 6 -------------
df1 = sol03.groupby('id').agg({'Data_fim_roteiro':'first'}).reset_index().sort_values('id')
df2 = sol03.groupby('id').agg({'Data_fim_roteiro':'last'}).reset_index().sort_values('id')
validate = df1 == df2

print(df1.isna().sum())
print(df2.isna().sum())
validate[validate['Data_fim_roteiro']==False].head()

### Alguns ids estão duplicados com datas finais diferentes. Nesses casos, a viagem foi para mais de um destino



# %%
# ------------- TRAT 7 -------------
inicio = sol03.groupby('id').agg({'Data_inicio_roteiro':'first'}).reset_index().sort_values('id')
fim = sol03.groupby('id').agg({'Data_fim_roteiro':'last'}).reset_index().sort_values('id')

df_fim = inicio.merge(fim, on='id', how='left')
df_fim = df_prest.merge(df_fim, on='id', how='left')

df_fim['Data_fim'] = df_fim['Data_fim_roteiro']
df_fim['Data_fim'].fillna(df_fim['Data_final'], inplace=True)
df_fim.loc[df_fim.Data_final > df_fim.Data_fim_roteiro, 'Data_fim'] = df_fim['Data_final']
df_fim.drop(columns=['Data_final', 'Data_fim_roteiro', 'Data_inicio_roteiro'], inplace=True)

print(len(df_fim))
df_fim.head()



# %%
# ------------- TRAT 8 -------------
path = project_folder + r"\CAD06 - Prestação de contas aprovadas.xlsx"

cad06 = pd.read_excel(path, header=2, usecols="C,F")
cad06.columns = ['Viagem', 'Data_finalizacao']

cad06['id'] = cad06.Viagem.str.split("-", n=1, expand=True)[0]
cad06 = cad06[pd.to_numeric(cad06['id'], errors='coerce').notnull()].reset_index(drop=True)
cad06['id'] = cad06['id'].astype('int64')

cad06.drop(columns=['Viagem'], inplace=True)

print(len(cad06))
cad06.head()

# %%
# ------------- TRAT 9 -------------
df1 = cad06.groupby('id').agg({'Data_finalizacao':'first'}).reset_index().sort_values('id')
df2 = cad06.groupby('id').agg({'Data_finalizacao':'last'}).reset_index().sort_values('id')
validate = df1 == df2

print(df1.isna().sum())
print(df2.isna().sum())

print(len(validate[validate['Data_finalizacao']==False]))
validate[validate['Data_finalizacao']==False].head()



# %%
# ------------- TRAT 10 -------------
df_conf = cad06.groupby('id').agg({'Data_finalizacao':'last'}).reset_index().sort_values('id')
df_conf = df_fim.merge(df_conf, on='id', how='left')

print(len(df_conf))
df_conf.head()



# %%
# ------------- TRAT 11 -------------
path = project_folder + r"\GER27 - Relatório de reprovações.xlsx"

ger27 = pd.read_excel(path, converters={'#':int}, usecols="A,L")
ger27.columns = ['id', 'Data_reprovacao']

ger27['Data_reprovacao'] = pd.to_datetime(ger27['Data_reprovacao'])

print(len(ger27))
ger27.head()



# %%
# ------------- TRAT 12 -------------
df1 = ger27.groupby('id').agg({'Data_reprovacao':'first'}).reset_index().sort_values('id')
df2 = ger27.groupby('id').agg({'Data_reprovacao':'last'}).reset_index().sort_values('id')
validate = df1 == df2

print(df1.isna().sum())
print(df2.isna().sum())

print(len(validate[validate['Data_reprovacao']==False]))
validate[validate['Data_reprovacao']==False].head()

### Regra de negócio: considerar apenas a primeira reprovação



# %%
# ------------- TRAT 13 -------------
df_repr = ger27.groupby('id').agg({'Data_reprovacao':'first'}).reset_index().sort_values('id')
df_repr = df_conf.merge(df_repr, on='id', how='left')

# Excluir casos com reprovação antes ou durante a viagem
df_repr.loc[df_repr.Data_reprovacao <= df_repr.Data_fim, 'Data_reprovacao'] = np.nan

print(len(df_repr))
df_repr.head()




# %%
# ------------- TRAT 14 -------------
connection_string_db2p = (
    r"Driver=SQL Server Native Client 11.0;"
    r"Server=SWDVMA1383;"
    r"Database=DBZB099D;"
    r"Trusted_Connection=yes;"
)
connection_url_db2p = URL.create(
    "mssql+pyodbc", 
    query={"odbc_connect": connection_string_db2p}
)
engine_db2p = sqlalchemy.create_engine(connection_url_db2p,
                                       fast_executemany=True,
                                       connect_args={'connect_timeout': 10},
                                       echo=False)
conn_db2p = engine_db2p.connect()

try:
    with engine_db2p.connect() as con:
        con.execute(sqlalchemy.text("SELECT 1"))
    print('engine is valid')
except Exception as e:
    print(f'Engine invalid: {str(e)}')



# %%
# ------------- TRAT 15 -------------
query = """
        WITH func AS (
            SELECT
            PES.NOM_PES AS 'NOME',
            EMP.NUM_CEN_CST AS 'CENTRO_CUSTO',
            ROW_NUMBER() OVER (PARTITION BY PES.NOM_PES ORDER BY COALESCE(EMP.DTA_DMS, getdate()) DESC) AS rn

            FROM EMPREGADO EMP

            LEFT JOIN PESSOA AS PES
            ON EMP.NUM_PES = PES.NUM_PES
        ),

        cc AS (
            SELECT
            CEC.COD_CEN_CST AS 'CENTRO_CUSTO',
            CEC.NOM_CEN_CST AS 'NOME_CC',
            PES.NOM_PES AS 'RESPONSAVEL',
            ROW_NUMBER() OVER (PARTITION BY CEC.COD_CEN_CST ORDER BY COALESCE(CEC.DTA_ALT, getdate()) DESC) AS rncc

            FROM CENTRO_CUSTO_CEC CEC

            LEFT JOIN EMPREGADO EMP
            ON CEC.NUM_MAT_EPG_RSP = EMP.NUM_MAT_EPG

            LEFT JOIN PESSOA PES
            ON EMP.NUM_PES = PES.NUM_PES
            
            WHERE CEC.DTA_DTV IS NULL
        )

        SELECT
        func.NOME,
        func.CENTRO_CUSTO,
        cc.NOME_CC,
        cc.RESPONSAVEL

        FROM func

        LEFT JOIN cc
        ON func.CENTRO_CUSTO = cc.CENTRO_CUSTO

        WHERE func.rn=1
        AND cc.rncc=1
        
        ORDER BY func.NOME
        """



# %%
# ------------- TRAT 16 -------------
cc = pd.read_sql(query, con=conn_db2p)
cc.columns = ['Colaborador', 'Num_CC', 'Nome_CC', 'Responsavel_CC']

cc['Num_CC'] = pd.to_numeric(cc['Num_CC'], errors='coerce')
cc['Num_CC'] = cc['Num_CC'].astype('Int64')
cc['Responsavel_CC'].fillna("", inplace=True)

print(len(cc))
cc.head()



# %%
# ------------- TRAT 17 -------------
df_cc = df_repr.merge(cc, on='Colaborador', how='left')

print(len(df_cc))
df_cc.tail()



# %%
# ------------- TRAT 18 -------------
reprovadas = df_cc[df_cc['Data_reprovacao'].notna()]
reprovadas['SLA_rep_prestacao'] = np.where(reprovadas['Data_reprovacao'] > reprovadas['Data_prestação_contas'],
    np.nan, (reprovadas['Data_prestação_contas'] - reprovadas['Data_reprovacao']).dt.days) # 5 dias
reprovadas['SLA_rep_conferencia'] = np.where(reprovadas['Data_reprovacao'] > reprovadas['Data_conferencia'],
    np.nan, (reprovadas['Data_conferencia'] - reprovadas['Data_prestação_contas']).dt.days) # 5 dias

df_cc['SLA_prestacao'] = (df_cc['Data_prestação_contas'] - df_cc['Data_fim']).dt.days # 10 dias
df_cc['SLA_aprovacao'] = (df_cc['Data_finalizacao'] - df_cc['Data_conferencia']).dt.days # 5 dias
df_cc['SLA_reprovacao'] = (df_cc['Data_reprovacao'] - df_cc['Data_conferencia']).dt.days # 5 dias

df_final = df_cc.merge(reprovadas[['id', 'SLA_rep_prestacao', 'SLA_rep_conferencia']], on='id', how='left')

print(len(df_final))
df_final.head()



# %% 
# ------------- IGNORAR -------------
# df_cc['Prestacao_atrasada'] = np.where(df_cc['SLA_prestacao'] > 10, "Sim", "Não")
# df_cc['Aprovacao_atrasada'] = np.where(df_cc['SLA_aprovacao'] > 5, "Sim", "Não")
# df_cc['Reprovacao_atrasada'] = np.where(df_cc['SLA_reprovacao'] > 5, "Sim", "Não")
# df_cc['Prestacao_pos_rep_atrasada'] = np.where(df_cc['SLA_rep_prestacao'] > 5, "Sim", "Não")
# df_cc['Conferencia_pos_rep_atrasada'] = np.where(df_cc['SLA_rep_conferencia'] > 5, "Sim", "Não")
# 
# ['Prestação de contas atrasada', 'Aprovação atrasada','Reprovação atrasada',
# 'Prestação pós reprovação atrasada', 'Conferência pós reprovação atrasada']
# 
# print(len(df_cc))
# df_cc.head()



# %%
# ------------- TRAT 19 -------------
for col in df_final.select_dtypes(include=['datetime64']).columns.tolist():
    df_final[col] = df_final[col].dt.strftime('%d/%m/%Y %H:%M:%S')

df_final.columns = ['id', 'Colaborador', 'Status', 'Data prestação de contas', 'Data conferência', 'Data início',
                 'Data fim', 'Data finalização', 'Data reprovação', 'Número CC', 'Nome CC', 'Responsável CC',
                 'SLA prestação de contas', 'SLA aprovação', 'SLA reprovação', 'SLA prestação pós reprovação',
                 'SLA conferência pós reprovação']

#df_final.columns = ['ID', 'COLABORADOR', 'DT_PRESTACAO', 'DT_CONFERENCIA', 'DT_INICIO', 'DT_FIM', 'DT_FINALIZACAO', 
#                    'DT_REPROVACAO', 'NUMERO_CC', 'NOME_CC', 'RESPONSAVEL_CC', 'SLA_PRESTACAO', 'SLA_APROVACAO',
#                    'SLA_REPROVACAO', 'SLA_PRESTACAO_POS_REPROVACAO', 'SLA_CONFERENCIA_POS_REPROVACAO']
#df_final['DT_INGESTAO'] = pd.Timestamp.now()

df_final



# %%
# ------------- TRAT 20 -------------
path = project_folder + r"\SLA Relatórios.xlsx"

with pd.ExcelWriter(path) as writer:
    df_final.to_excel(writer, sheet_name='Relatório', index=False)



# %%
# ------------- TRAT 21 -------------
file_to_copy3 = path
destination_directory3 = r"T:\SIC\Seg Patrimonial\Manutenções\SLA Relatórios.xlsx"
shutil.copy(file_to_copy3, destination_directory3)

pyautogui.alert('O código foi finalizado!')



# %%
