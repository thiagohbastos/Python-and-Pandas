#%%
# ------------- PACOTES -------------
import pandas as pd
import os
import pyodbc
import warnings
from datetime import date
from time import sleep
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL
warnings.filterwarnings('ignore')
warnings.simplefilter(action='ignore', category=FutureWarning)


#%%
# ------------- CONECTANDO SQL SERVER -------------
conn_gnu = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server=SQLGDNP;"
    "Database=GNU;"
    "Trusted_Connection=yes;")


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
# ------------- COLETANDO INFO ATMS -------------
cursor_gnu = conn_gnu.cursor()
query_info_atms = """SELECT A.COD_CEN
	,B.RESPONSAVEL
	,A.DES_CEN
	,A.NUM_DND
	,A.NOME_AGENCIA
	,A.IDT_TML
    ,A.NOM_TRANSP
FROM [MERCANTIL\B040466].TC_INFO_AG_TSR_TML A
LEFT JOIN [MERCANTIL\Y038650].TC_RESPONSAVEL_TESOURARIA B
ON A.COD_CEN = B.COD_TESOURARIA
WHERE A.IDT_TML IS NOT NULL
"""
info_atms = pd.read_sql_query(query_info_atms, conn_gnu)


#%%
# ------------- COLETANDO NOMES DOS ARQUIVOS HST -------------
resultado = []
meses = ['0'
         ,'JANEIRO'
         , 'FEVEREIRO'
         , 'MARÇO'
         , 'ABRIL'
         , 'MAIO'
         , 'JUNHO'
         , 'JULHO'
         , 'AGOSTO'
         , 'SETEMBRO'
         , 'OUTUBRO'
         , 'NOVEMBRO'
         , 'DEZEMBRO']
mes_atual = meses[date.today().month]
ano = date.today().year
atual = f'K:/GSAS/09 - Coordenacao Gestao Numerario/00 Supervisão de Numerário/P.A. POSTO ATENDIMENTO/PROGRAMAÇÃO ATM/{ano}'
try:
    pastas_dir_atual = os.listdir(atual)
except Exception as erro:
    print(erro)
    sleep(60)
for x in pastas_dir_atual:
    if mes_atual in x:
        temp = os.path.join(atual, x)
        dias = os.listdir(temp)
        for dia in dias:
            #if dia in ('25.08', '28.08', '29.08'):
                temp_1 = os.path.join(temp, dia)
                arquivos = os.listdir(temp_1)
                for arquivo in arquivos:
                    if 'EMERGENCIAL' in arquivo.upper() and arquivo not in resultado:
                        resultado.append(os.path.join(temp_1, arquivo))
                    else:
                        continue
    else:
        continue


# %%
# ------------- EXTRAINDO HISTÓRICO PARA VARIÁVEL -------------
cont = 0
colunas = ['Código Banco'
,'Num. Ordem de Serviço'
,'Data Programação'
,'Data Entrega'
,'Tipo Entrega'
,'Ponto/ATM'
,'Nome da Agência'
,'Horário'
,'K7-A R$ 100'
,'K7-A R$ 50'
,'K7-A R$ 20'
,'K7-A R$ 10'
,'K7-A R$ 5'
,'K7-A R$ 2'
,'K7-B R$ 200'
,'K7-B R$ 100'
,'K7-B R$ 50'
,'K7-B R$ 20'
,'K7-B R$ 10'
,'K7-B R$ 5'
,'K7-B R$ 2'
,'K7-C R$ 100'
,'K7-C R$ 50'
,'K7-C R$ 20'
,'K7-C R$ 10'
,'K7-D R$ 200'
,'K7-D R$ 100'
,'K7-D R$ 50'
,'K7-D R$ 20'
,'Valor Total']
sftp_var = pd.DataFrame(columns=[colunas])
sftp_var = ''
for arquivo_prog in resultado:
    if '.xlsx'.upper() == arquivo_prog[-5:].upper() and 'ATM' in os.path.split(arquivo_prog)[1].upper():
                try:
                    temp_xlsx = pd.read_excel(arquivo_prog)
                    if len(temp_xlsx) >= 1 and len(sftp_var) >= 1:
                        sftp_var = pd.concat([sftp_var, temp_xlsx])
                    elif len(temp_xlsx) >= 1 and len(sftp_var) == 0:
                        sftp_var = temp_xlsx
                except Exception as erro:
                    print(f'''ERRO NA LEITURA DO ARQUIVO {arquivo_prog}

    ERRO:
    {erro}\n''')
                    print(f'''Em caso de arquivo aberto com outro usuário, solicite sua liberação antes de rodar o processo novamente.
                    
    Em caso de erros diversos, favor printar o erro acima e enviar para o responsável pelo programa (Thiago Bastos)''')
                    input('\nPrecione qualquer tecla para sair...\n')
                    exit()
    cont += 1
    print(cont)
bkp1 = sftp_var.copy()


#%%
# ------------- TRATANDO TABELA -------------             
sftp_var = sftp_var[sftp_var['Código Banco'] == 389].reset_index(drop=True)
sftp_var['DUPLICADO'] = sftp_var.duplicated()

duplicados = sftp_var[sftp_var['DUPLICADO'] == True].drop_duplicates()

sftp_var = sftp_var.drop_duplicates()
sftp_var['Horário'] = sftp_var['Horário'].fillna('19:00')
sftp_var = sftp_var.drop(columns=['valida'
                            , 'Código Banco'
                            , 'Tipo Entrega'
                            , 'DUPLICADO'
                            #, 1
                            ]
                            )
sftp_var = sftp_var.fillna(0)
sftp_var = sftp_var.reset_index(drop=True)
bkp2 = sftp_var.copy()


#%%
# ------------- TRATANDO VALORES -------------
disc_colunas = list(sftp_var.columns)
del disc_colunas[:6]
for x in disc_colunas:
    if sftp_var[x].dtype not in('float64','int64'):
        for indice in range(0, len(sftp_var)):
            if not str(sftp_var[x][indice]).replace('.','').isnumeric():
                sftp_var.at[indice, x] = 0
        sftp_var[x] = sftp_var[x].astype('float64')

sftp_var['Num. Ordem de Serviço'] = sftp_var['Num. Ordem de Serviço'].astype('int64').astype('object')
sftp_var['Data Entrega'] = sftp_var['Data Entrega'].astype('datetime64[ns]')
sftp_var['Data Programação'] = sftp_var['Data Programação'].astype('datetime64[ns]')
sftp_var.rename(columns={'Ponto/ATM':'IDT_TML', 'Nome da Agência': 'Nome_da_Agencia'}, inplace='True')
sftp_var = pd.merge(sftp_var, info_atms, how='left', on='IDT_TML')


# %%
# ------------- TRATANDO COLUNAS PARA INSERÇÃO -------------
enviado_sftp = sftp_var.copy()
enviado_sftp.rename(columns={
    'Num. Ordem de Serviço': 'OS',
    'Data Programação' : 'DTA_PROG',
    'Data Entrega' : 'DTA_ENTREGA',
    'Horário' : 'HORA',
    'Valor Total' : 'TOTAL',
    'K7-A R$ 100': 'K7A_100',
    'K7-A R$ 50': 'K7A_50',
    'K7-A R$ 20': 'K7A_20',
    'K7-A R$ 10': 'K7A_10',
    'K7-A R$ 5': 'K7A_5',
    'K7-A R$ 2': 'K7A_2',
    'K7-B R$ 200': 'K7B_200',
    'K7-B R$ 100': 'K7B_100',
    'K7-B R$ 50': 'K7B_50',
    'K7-B R$ 20': 'K7B_20',
    'K7-B R$ 10': 'K7B_10',
    'K7-B R$ 5': 'K7B_5',
    'K7-B R$ 2': 'K7B_2',
    'K7-C R$ 100': 'K7C_100',
    'K7-C R$ 50': 'K7C_50',
    'K7-C R$ 20': 'K7C_20',
    'K7-C R$ 10': 'K7C_10',
    'K7-D R$ 200': 'K7D_200',
    'K7-D R$ 100': 'K7D_100',
    'K7-D R$ 50': 'K7D_50',
    'K7-D R$ 20': 'K7D_20'
}, inplace='True'
)
enviado_sftp = enviado_sftp.fillna(0)
enviado_sftp = enviado_sftp.reset_index(drop=True)
enviado_sftp = enviado_sftp[['NOM_TRANSP', 'DTA_PROG', 'DTA_ENTREGA', 'HORA', 
       'IDT_TML', 'COD_CEN', 'DES_CEN', 'NUM_DND', 'NOME_AGENCIA', 'OS', 
       'K7A_100', 'K7A_50', 'K7A_20', 'K7A_10', 'K7A_5', 'K7A_2', 'K7B_200',
       'K7B_100', 'K7B_50', 'K7B_20', 'K7B_10', 'K7B_5', 'K7B_2', 'K7C_100',
       'K7C_50', 'K7C_20', 'K7C_10', 'K7D_200', 'K7D_100', 'K7D_50', 'K7D_20',
       'TOTAL']]


#%%
# ------------- TRATANDO VALORES² -------------
disc_colunas = list(enviado_sftp.columns)
for x in disc_colunas:
    if enviado_sftp[x].dtype == 'float64':
        enviado_sftp[x] = enviado_sftp[x].astype('int64')


#%%
# ------------- INSERÇÕES SQLALCHEMY -------------
conn_gnu.execute('''TRUNCATE TABLE [MERCANTIL\B042786].TI_PROGRAMACAO_EMERGENCIAL''')
conn_gnu.commit()
conn_alchemy_gnu = engine.connect()
with conn_alchemy_gnu as conn:
    with conn.begin() as beg:
        enviado_sftp.to_sql(name="TI_PROGRAMACAO_EMERGENCIAL"
                , con=conn
                , if_exists= 'replace' #'append'
                , index=False
                , chunksize=1000
                , schema='MERCANTIL\B042786')
        beg.commit()


# %%
# ------------- TRATANDO DADOS NÃO ENCONTRADOS -------------
conn_gnu.execute('''UPDATE [MERCANTIL\B042786].TC_PROGRAMACAO_ENVIADA_K7_HST

SET NOM_TRANSP = C.NOM_TRANSP , COD_CEN = C.COD_CEN , DES_CEN = C.DES_CEN , NUM_DND = C.NUM_DND , NOME_AGENCIA = B.NOME_DEPENDENCIA

--SELECT *
FROM [MERCANTIL\B042786].TC_PROGRAMACAO_ENVIADA_K7_HST A
LEFT JOIN
(
	SELECT  DISTINCT B.IDT_TML
	       ,B.NOME_DEPENDENCIA
	       ,B.NUM_DND
	FROM [MERCANTIL\B038660].TI_PRINCIPAL_TERMINAIS_HST B
) B
ON A.IDT_TML = B.IDT_TML
LEFT JOIN
(
	SELECT  DISTINCT C.NUM_DND
	       ,C.COD_CEN
	       ,C.NOM_TRANSP
	       ,C.DES_CEN
	FROM [MERCANTIL\B040466].TC_INFO_AG_TSR_TML C
) C
ON B.NUM_DND = C.NUM_DND
WHERE A.NOM_TRANSP NOT IN ('BRINKS', 'PROSEGUR', 'PROTEGE') 
    OR A.NOM_TRANSP IS NULL''')
conn_gnu.commit()


#%%
# ------------- FINALIZANDO PACOTE -------------
cursor_gnu.close()
