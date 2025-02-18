#%%
# ------------- PACOTES -------------
#<PASTA DO PROJETO>\venv\Scripts\activate
#pyinstaller --name="VALIDA PROGRAMAÇÃO" --onefile Valida_programacao.py --icon icone.ico
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
# ------------- COLETANDO OPÇÃO DO USUÁRIO -------------
while True:
    print('''GOSTARIA DE EXECUTAR A VALIDAÇÃO DO SFTP OU DA PASTA INICIAL?
[1] PARA VALIDA SFTP
[2] PARA VALIDA PASTA
''')
    try:
        resposta = int(input('Opção Selecionada: '))
        if resposta not in (1, 2):
            print('Opção invalida! Responda novamente.\n')
        else:
            break
    except:
        print('Opção invalida! Responda novamente.\n')
if resposta == 1:
    print('\nIniciando a validação do SFTP... \n')
else:
    print('\nIniciando a validação dos ARQUIVOS DA PASTA... \n')


#%%
# ------------- PROGRAMADOS GNU -------------
query_prog = '''SELECT  AG.IDT_TML
       ,AB.DTA_HOR_PRG AS DTA_HOR_GNU
       ,AB.DTA_HOR_PRG AS DTA_HOR_MERGE
       ,AB.VLR_ABT_PRG
FROM ABASTEC_PROGRAMADO AB
INNER JOIN TML_AGENCIA AG
ON AB.IDT_TML_DND = AG.IDT_TML_DND
WHERE cast(DTA_HOR_PRG AS date) > CAST(GETDATE() AS date)
--AND LEFT(CAST(AB.DTA_HOR_PRG AS TIME) , 2) NOT IN (19)'''
prog_gnu = pd.read_sql_query(query_prog, conn_mbcorp)

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

prog_gnu = pd.merge(prog_gnu, info_atms, how='left', on='IDT_TML')


#%%
# ------------- CRIANDO VARIÁVEIS AUXILIARES DA LEITURA -------------
transportadoras = ['BRINKS', 'PROSEGUR', 'PROTEGE']
if resposta == 1:
    loc_brinks = f'T:\gnu\SFTP\{transportadoras[0]}\ENVIADOS'
    loc_prosegur = f'T:\gnu\SFTP\{transportadoras[1]}\ENVIADOS'
    loc_protege = f'T:\gnu\SFTP\{transportadoras[2]}\ENVIADOS'
    locais = [loc_brinks, loc_prosegur, loc_protege]
elif resposta == 2:
    locais = ['K:/GSAS/09 - Coordenacao Gestao Numerario/09 Prototipos SSIS/B042786/SISS/TI_REMESSAS_ATM/ATM']

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


#%%
# ------------- COLETANDO VALORES INSERIDOS NO DISCO/PASTA -------------
for transp in locais:
    arquivos_pasta = os.listdir(transp)
    for arquivo_prog in arquivos_pasta:
        temp_caminho = transp + '/' + arquivo_prog
        if '.xlsx'.upper() == arquivo_prog[-5:].upper() and 'ATM' in arquivo_prog.upper():
            try:
                temp_xlsx = pd.read_excel(temp_caminho)
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


#%%
# ------------- TRATAMENTO INICIAL SFTP -------------
sftp_var = sftp_var[sftp_var['Código Banco'] == 389].reset_index(drop=True)
sftp_var['DUPLICADO'] = sftp_var.duplicated()

duplicados = sftp_var[sftp_var['DUPLICADO'] == True].drop_duplicates()

sftp_var = sftp_var.drop_duplicates()
sftp_var['Horário'] = sftp_var['Horário'].fillna('19:00')

#CRIANDO CÓPIA DA BASE PARA HISTÓRICO
enviado_sftp = sftp_var.copy()

#TRATANDO DADOS SFTP
sftp_var = sftp_var.drop(columns=['valida'
                            , 'Código Banco'
                            , 'Num. Ordem de Serviço'
                            , 'Tipo Entrega'
                            ])
sftp_var = sftp_var.fillna(0)


#%%
# ------------- TRATANDO VALORES INCOMPATÍVEIS PLANS -------------
disc_colunas = list(sftp_var.columns)
if len(disc_colunas) == 28:
    del disc_colunas[:5]
for x in disc_colunas:
    if sftp_var[x].dtype not in('float64','int64'):
        for indice in range(0, len(sftp_var)):
            if not str(sftp_var[x][indice]).replace('.','').isnumeric():
                sftp_var.at[indice, x] = 0
        sftp_var[x] = sftp_var[x].astype('float64')

sftp_var.rename(columns={'Ponto/ATM':'IDT_TML', 'Nome da Agência': 'Nome_da_Agencia'}, inplace='True')
sftp_var = pd.merge(sftp_var, info_atms, how='left', on='IDT_TML')


#%%
# ------------- TRATANDO DATA/HORA DE ENTREGA NO SFTP -------------
sftp_var['DTA_HOR_DSC'] = sftp_var['Data Entrega'].astype(str) + ' ' + sftp_var['Horário'].astype(str) + ':00'
sftp_var['DTA_HOR_DSC'] = [x[:19] for x in sftp_var['DTA_HOR_DSC']]
sftp_var['DTA_HOR_DSC'] = pd.to_datetime(sftp_var['DTA_HOR_DSC'])


#%%
# ------------- VALIDANDO SOMATÓRIA DOS K7s e DUPLICADOS -------------
sftp_var['TOT_CALC'] = sftp_var['K7-A R$ 100'] + sftp_var['K7-A R$ 50'] + sftp_var['K7-A R$ 20'] + sftp_var['K7-A R$ 10'] + sftp_var['K7-A R$ 5'] + sftp_var['K7-A R$ 2'] + sftp_var['K7-B R$ 200'] + sftp_var['K7-B R$ 100'] + sftp_var['K7-B R$ 50'] + sftp_var['K7-B R$ 20'] + sftp_var['K7-B R$ 10'] + sftp_var['K7-B R$ 5'] + sftp_var['K7-B R$ 2'] + sftp_var['K7-C R$ 100'] + sftp_var['K7-C R$ 50'] + sftp_var['K7-C R$ 20'] + sftp_var['K7-C R$ 10'] + sftp_var['K7-D R$ 200'] + sftp_var['K7-D R$ 100'] + sftp_var['K7-D R$ 50'] + sftp_var['K7-D R$ 20']
sftp_var['VALIDA_SOMA'] = sftp_var['TOT_CALC'] == sftp_var['Valor Total']
soma_errada = sftp_var[sftp_var['VALIDA_SOMA'] == False]

sftp_var['DTA_HOR_MERGE'] = sftp_var['DTA_HOR_DSC']

sftp_var = sftp_var[['Data Programação', 'Data Entrega', 'IDT_TML', 'Nome_da_Agencia',
       'Horário', 'Valor Total', 'DTA_HOR_DSC', 'COD_CEN',
       'RESPONSAVEL', 'DES_CEN', 'NUM_DND', 'NOME_AGENCIA', 'TOT_CALC',
       'VALIDA_SOMA', 'DUPLICADO']]


#%%
# ------------- CRIANDO DATAFRAME QUE 'VALIDA SFTP E GNU' -------------
nao_disco = 0
nao_gnu = 0
valida = pd.merge(sftp_var, prog_gnu, how='outer', on=['IDT_TML', 'COD_CEN', 'RESPONSAVEL', 'DES_CEN', 'NUM_DND', 'NOME_AGENCIA'])
valida = valida[['COD_CEN'
                 , 'DES_CEN'
                 , 'RESPONSAVEL'
                 , 'NUM_DND'
                 , 'NOME_AGENCIA'
                 , 'IDT_TML'
                 , 'DTA_HOR_DSC'
                 , 'TOT_CALC'
                 , 'DTA_HOR_GNU'
                 , 'VLR_ABT_PRG'
                 , 'VALIDA_SOMA'
                 ]]


#%%
# ------------- INSERINDO RESULTADOS NO SQLGDNP -------------
cursor_gnu.execute("""TRUNCATE TABLE [MERCANTIL\B042786].TC_VALIDA_DISCO""")
valida.fillna(0, inplace=True)
for index, row in valida.iterrows():
    if row.VALIDA_SOMA:
        somatorio = 'OK'
    elif not row.VALIDA_SOMA and row.TOT_CALC != 0:
        somatorio = 'SOMA DIVERGENTE'
    else:
        somatorio = '-'

    if row.TOT_CALC == row.VLR_ABT_PRG and row.VALIDA_SOMA:
        status = 'OK'
    elif row.TOT_CALC == row.VLR_ABT_PRG and not row.VALIDA_SOMA:
        status = 'AVALIAR SOMATORIO'
    elif row.TOT_CALC == 0:
        status = 'NÃO ESTÁ NO DISCO'
        nao_disco += 1
    elif row.VLR_ABT_PRG == 0:
        status = 'NÃO ESTÁ NO GNU'
        nao_gnu += 1
    elif row.TOT_CALC != row.VLR_ABT_PRG:
        status = 'VALOR DIVERGENTE'

    cursor_gnu.execute(f"""INSERT INTO [MERCANTIL\B042786].TC_VALIDA_DISCO 
        (COD_CEN
        ,RESPONSAVEL
        ,DES_CEN
        ,NUM_DND
        ,NOME_AGENCIA
        ,IDT_TML
        ,DATA_GNU
        ,SUPRIMENTO_PROGRAMADO_D1
        ,VALOR_DSC
        ,DATA_DSC
        ,STATUS
        ,SOMATORIO) 
        VALUES 
        (
        ?
        , ?
        , ?
        , ?
        , ?
        , ?
        , ?
        , ? 
        , ? 
        , ?
        , ?
        , ?
        )
        """
        ,int(row.COD_CEN)
        ,row.RESPONSAVEL
        ,row.DES_CEN
        ,int(row.NUM_DND)
        ,row.NOME_AGENCIA
        ,int(row.IDT_TML)
        ,row.DTA_HOR_GNU
        ,row.VLR_ABT_PRG
        ,row.TOT_CALC
        ,row.DTA_HOR_DSC
        ,status
        ,somatorio
        )
    
cursor_gnu.execute("""UPDATE [MERCANTIL\B042786].[TC_VALIDA_DISCO]
SET DATA_DSC = NULL, VALOR_DSC = NULL
FROM [MERCANTIL\B042786].[TC_VALIDA_DISCO] A
WHERE A.VALOR_DSC = 0

UPDATE [MERCANTIL\B042786].[TC_VALIDA_DISCO]
SET DATA_GNU = NULL, SUPRIMENTO_PROGRAMADO_D1 = NULL
FROM [MERCANTIL\B042786].[TC_VALIDA_DISCO] A
WHERE A.SUPRIMENTO_PROGRAMADO_D1 = 0""")
cursor_gnu.commit()


#%%
# ------------- INÍCIO DO TRATAMENTO PARA HST DA PROG -------------
enviado_sftp = enviado_sftp.drop(columns=['valida'
                            , 'Código Banco'
                            , 'Tipo Entrega'
                            , 'DUPLICADO'
                            ])
enviado_sftp = enviado_sftp.fillna(0)
enviado_sftp = enviado_sftp.reset_index(drop=True)

# ------------- TRATANDO VALORES COLUNAS HST -------------
hst_colunas = list(enviado_sftp.columns)
del hst_colunas[:6]
for x in hst_colunas:
    if enviado_sftp[x].dtype not in('float64','int64'):
        for indice in range(0, len(enviado_sftp)):
            if not str(enviado_sftp[x][indice]).replace('.','').isnumeric():
                enviado_sftp.at[indice, x] = 0
        enviado_sftp[x] = enviado_sftp[x].astype('float64')

enviado_sftp['Num. Ordem de Serviço'] = enviado_sftp['Num. Ordem de Serviço'].astype('object')
enviado_sftp.rename(columns={'Ponto/ATM':'IDT_TML', 'Nome da Agência': 'Nome_da_Agencia'}, inplace='True')
enviado_sftp = pd.merge(enviado_sftp, info_atms, how='left', on='IDT_TML')


#%%
# ------------- TRATANDO DISPOSIÇÃO DAS COLUNAS HST -------------
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
# ------------- TRATANDO VALORES HST² -------------
disc_colunas = list(enviado_sftp.columns)
for x in disc_colunas:
    if enviado_sftp[x].dtype == 'float64':
        enviado_sftp[x] = enviado_sftp[x].astype('int64')


#%%
# ------ LIMPANDO POSSÍVEIS DUPLICATAS E INSERINDO NOVOS VALORES ------
cursor_gnu.execute('''DELETE FROM 
    [MERCANTIL\B042786].TC_PROGRAMACAO_ENVIADA_K7_HST
    WHERE CAST(DTA_PROG AS DATE) = CAST(GETDATE() AS DATE)'''
)
cursor_gnu.commit()

with conn_alchemy_gnu as conn:
    with conn.begin() as beg:
        enviado_sftp.to_sql(name="TC_PROGRAMACAO_ENVIADA_K7_HST"
                , con=conn
                , if_exists='append'
                , index=False
                , chunksize=500
                , schema='MERCANTIL\B042786')
        beg.commit()


# %%
# ------------- TRATANDO DADOS NÃO ENCONTRADOS -------------
cursor_gnu.execute('''UPDATE [MERCANTIL\B042786].TC_PROGRAMACAO_ENVIADA_K7_HST

SET NOM_TRANSP = C.NOM_TRANSP 
, COD_CEN = C.COD_CEN 
, DES_CEN = C.DES_CEN 
, NUM_DND = C.NUM_DND 
, NOME_AGENCIA = B.NOME_DEPENDENCIA

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
cursor_gnu.commit()


#%%
# ------------- FINALIZANDO PACOTE -------------
print('Validação finalizada com ÊXITO!\n')
print('\n----------- RELATÓRIO GERAL -----------')
print(f'Atms que NÃO ESTÃO no GNU      : {nao_gnu}')
print(f'Atms que NÃO ESTÃO no DISCO    : {nao_disco}')
print(f'Atms com o somatório divergente: {len(soma_errada)}')
print(f'Atms duplicados                : {len(duplicados)}')

if len(duplicados) >= 1:
    print('Atms duplicados nos arquivos de abastecimento:\n')
    print(duplicados, end='\n\n')




###################################################################################
#%% DAQUI EM DIANTE COMEÇA A EXTRAÇÃO DOS EMERGENCIAIS DO MÊS
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
print(f'''------------------------------------------
INICIANDO A EXTRAÇÃO DOS EMERGENCIAIS DO MÊS DE {mes_atual}
POR FAVOR NÃO FECHE O PROGRAMA!
------------------------------------------''')


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
ano = date.today().year
atual = f'K:/GSAS/09 - Coordenacao Gestao Numerario/00 Supervisão de Numerário/P.A. POSTO ATENDIMENTO/PROGRAMAÇÃO ATM/{ano}'
pastas_dir_atual = os.listdir(atual)
for x in pastas_dir_atual:
    #if mes_atual in x:
        temp = os.path.join(atual, x)
        dias = os.listdir(temp)
        for dia in dias:
            temp_1 = os.path.join(temp, dia)
            arquivos = os.listdir(temp_1)
            for arquivo in arquivos:
                if 'EMERGENCIAL' in arquivo.upper() and arquivo not in resultado:
                    resultado.append(os.path.join(temp_1, arquivo))
                else:
                    continue
    #else:
    #    continue


# %%
# ------------- EXTRAINDO HISTÓRICO PARA VARIÁVEL -------------
if len(resultado) >= 1:
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
                    , chunksize=100
                    , schema='MERCANTIL\B042786')
            beg.commit()


    # %%
    # ------------- TRATANDO DADOS NÃO ENCONTRADOS -------------
    conn_gnu.execute('''
    MERGE [MERCANTIL\B042786].TC_PROGRAMACAO_EMERGENCIAL AS DESTINO 
    USING [MERCANTIL\B042786].TI_PROGRAMACAO_EMERGENCIAL AS ORIGEM
    ON ORIGEM.OS = DESTINO.OS 
    AND ORIGEM.NUM_DND = DESTINO.NUM_DND
    AND ORIGEM.IDT_TML = DESTINO.IDT_TML
    AND ORIGEM.TOTAL = DESTINO.TOTAL

    --HÁ REGISTRO NO DESTINO E NA ORIGEM: 
    /*WHEN MATCHED AND DESTINO.TOTAL != ORIGEM.TOTAL
    THEN UPDATE
    SET QTD_LOG = ORIGEM.QTD_LOG, QTD_PLANS = ORIGEM.QTD_PLANS*/

    --NÃO HÁ REGISTROS NO DESTINO, PORÉM HÁ NA ORIGEM:
    WHEN NOT MATCHED THEN
    INSERT (NOM_TRANSP, DTA_PROG, DTA_ENTREGA, HORA, IDT_TML, COD_CEN, DES_CEN, NUM_DND, NOME_AGENCIA, OS, K7A_100, K7A_50, K7A_20, K7A_10, K7A_5, K7A_2, K7B_200, K7B_100, K7B_50, K7B_20, K7B_10, K7B_5, K7B_2, K7C_100, K7C_50, K7C_20, K7C_10, K7D_200, K7D_100, K7D_50, K7D_20, TOTAL)

    VALUES (
    ORIGEM.NOM_TRANSP
    , ORIGEM.DTA_PROG
    , ORIGEM.DTA_ENTREGA
    , ORIGEM.HORA
    , ORIGEM.IDT_TML
    , ORIGEM.COD_CEN
    , ORIGEM.DES_CEN
    , ORIGEM.NUM_DND
    , ORIGEM.NOME_AGENCIA
    , ORIGEM.OS
    , ORIGEM.K7A_100
    , ORIGEM.K7A_50
    , ORIGEM.K7A_20
    , ORIGEM.K7A_10
    , ORIGEM.K7A_5
    , ORIGEM.K7A_2
    , ORIGEM.K7B_200
    , ORIGEM.K7B_100
    , ORIGEM.K7B_50
    , ORIGEM.K7B_20
    , ORIGEM.K7B_10
    , ORIGEM.K7B_5
    , ORIGEM.K7B_2
    , ORIGEM.K7C_100
    , ORIGEM.K7C_50
    , ORIGEM.K7C_20
    , ORIGEM.K7C_10
    , ORIGEM.K7D_200
    , ORIGEM.K7D_100
    , ORIGEM.K7D_50
    , ORIGEM.K7D_20
    , ORIGEM.TOTAL
    )
    ;
    ''')
    conn_gnu.commit()

    conn_gnu.execute('''
    MERGE [MERCANTIL\B042786].TC_PROGRAMACAO_ENVIADA_K7_HST AS DESTINO 
    USING [MERCANTIL\B042786].TC_PROGRAMACAO_EMERGENCIAL AS ORIGEM
    ON ORIGEM.OS = DESTINO.OS 
    AND ORIGEM.NUM_DND = DESTINO.NUM_DND
    AND ORIGEM.IDT_TML = DESTINO.IDT_TML
    AND ORIGEM.TOTAL = DESTINO.TOTAL

    --HÁ REGISTRO NO DESTINO E NA ORIGEM: 
    /*WHEN MATCHED AND DESTINO.TOTAL != ORIGEM.TOTAL
    THEN UPDATE
    SET QTD_LOG = ORIGEM.QTD_LOG, QTD_PLANS = ORIGEM.QTD_PLANS*/

    --NÃO HÁ REGISTROS NO DESTINO, PORÉM HÁ NA ORIGEM:
    WHEN NOT MATCHED THEN
    INSERT (NOM_TRANSP, DTA_PROG, DTA_ENTREGA, HORA, IDT_TML, COD_CEN, DES_CEN, NUM_DND, NOME_AGENCIA, OS, K7A_100, K7A_50, K7A_20, K7A_10, K7A_5, K7A_2, K7B_200, K7B_100, K7B_50, K7B_20, K7B_10, K7B_5, K7B_2, K7C_100, K7C_50, K7C_20, K7C_10, K7D_200, K7D_100, K7D_50, K7D_20, TOTAL)

    VALUES (
    ORIGEM.NOM_TRANSP
    , ORIGEM.DTA_PROG
    , ORIGEM.DTA_ENTREGA
    , ORIGEM.HORA
    , ORIGEM.IDT_TML
    , ORIGEM.COD_CEN
    , ORIGEM.DES_CEN
    , ORIGEM.NUM_DND
    , ORIGEM.NOME_AGENCIA
    , ORIGEM.OS
    , ORIGEM.K7A_100
    , ORIGEM.K7A_50
    , ORIGEM.K7A_20
    , ORIGEM.K7A_10
    , ORIGEM.K7A_5
    , ORIGEM.K7A_2
    , ORIGEM.K7B_200
    , ORIGEM.K7B_100
    , ORIGEM.K7B_50
    , ORIGEM.K7B_20
    , ORIGEM.K7B_10
    , ORIGEM.K7B_5
    , ORIGEM.K7B_2
    , ORIGEM.K7C_100
    , ORIGEM.K7C_50
    , ORIGEM.K7C_20
    , ORIGEM.K7C_10
    , ORIGEM.K7D_200
    , ORIGEM.K7D_100
    , ORIGEM.K7D_50
    , ORIGEM.K7D_20
    , ORIGEM.TOTAL
    )
    ;
    ''')
    conn_gnu.commit()


#%%
input('\nPrecione qualquer tecla para sair...\n')


#%%
conn_gnu.close()
conn_mbcorp.close()

# %%
