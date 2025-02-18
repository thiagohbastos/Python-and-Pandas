#%%
# ---------- IMPORTANDO PACOTES ----------
import pandas as pd
from math import isnan
import pyodbc
import warnings
warnings.simplefilter("ignore")


#%%
# ------------- CONECTANDO BANCOS DE DADOS -------------
with open("C:/Users/b042786/Documents/Projetos-MB/___ÚTEIS___/Usuário e Senha.txt", 'r') as info_pessoal:
   usuario = info_pessoal.readline().replace('\n', '')
   senha_db2 = info_pessoal.readline()
   info_pessoal.close()

conn_db2 = pyodbc.connect(
        'Driver={IBM DB2 ODBC DRIVER - IBMDBCL1}; '
        'Hostname=db2gatev11; '
        'Port=50000; '
        #'Protocol=TCPIP; '
        'Database=db2PMVS; '
        #'CurrentSchema=schema; '
        f'UID={usuario}; '
        f'PWD={senha_db2};'
        )

conn_gdnp = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server=sqlgdnp;"
    "Database=GNU;"
    "Trusted_Connection=yes;")
cursor = conn_gdnp.cursor()


# %%
#---------- EXTRAINDO DADOS EM SEUS CAMINHOS ----------
credores = pd.read_excel('//fsclt01grps.mercantil.com.br/grupos/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/DEVEDORES E CREDORES/GESTÃO CREDORES - 7914-5.xlsm', skiprows=3, sheet_name='Credores')
erro_saque = pd.read_excel('//fsclt01grps.mercantil.com.br/grupos/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/ERRO DE SAQUE E DEPOSITO/NOVO_ERRO DE SAQUE.xlsm', skiprows=4, sheet_name='ERROS DE SAQUE')
erro_deposito = pd.read_excel('//fsclt01grps.mercantil.com.br/grupos/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/ERRO DE SAQUE E DEPOSITO/NOVO_ERRO DE DEPÓSITO.xlsm', skiprows=4, sheet_name='ERROS DE DEPÓSITO')
devedores = pd.read_excel('//fsclt01grps.mercantil.com.br/grupos/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/DEVEDORES E CREDORES/GESTÃO DEVEDORES - 6864-3.xlsm', skiprows=4, sheet_name='Devedores')


#%%
#---------- TRATANDO COLUNAS E REORDENANDO ----------
credores['COD_TIP_LCT'] = 1
credores['NUM_DMD'] = 0
credores['COD_TSR'] = credores['Tesouraria']
credores['COD_STT_LCT'] = [1 if x[11] > 0 else 2 for x in credores.values]
credores = credores.rename(columns={'Data Diferença': 'DTA_LCT', 
        'Nº dependência':'NUM_DND' , 
        'ATM': 'COD_TML',
        'Valor' :'VLR_LCT', 
        'Tesouraria':'IDT_CTA_IDZ', 
        'Diferença':'VLR_DPN', 
        'Documento':'NUM_DOC'})
credores.fillna(0, inplace=True)
#credores = credores[credores['ATM'] != 0]

devedores['COD_TIP_LCT'] = 2
devedores['COD_TSR'] = devedores['Tesouraria']
devedores['NUM_DMD'] = 0
devedores['COD_STT_LCT'] = [1 if x[11] > 0 else 2 for x in devedores.values]
devedores = devedores.rename(columns={'Data Diferença': 'DTA_LCT', 
        'Nº dependência':'NUM_DND' , 
        'ATM': 'COD_TML',
        'Valor' :'VLR_LCT', 
        'Tesouraria':'IDT_CTA_IDZ', 
        'Diferença':'VLR_DPN', 
        'Número documento CAB':'NUM_DOC'})
devedores.fillna(0, inplace=True)

erro_saque['NUM_DOC'] = 0
erro_saque['COD_TSR'] = erro_saque['TECE']
erro_saque['COD_STT_LCT'] = [1 if x[17] > 0 else 2 for x in erro_saque.values]
erro_saque = erro_saque.rename(columns={'DATA DO SAQUE': 'DTA_LCT', 
        'COD_DEPENDENCIA':'NUM_DND' ,
        'DEMANDA':'NUM_DMD' , 
        'ATM': 'COD_TML',
        'VALOR DA DIFERENÇA' :'VLR_LCT', 
        'TECE':'IDT_CTA_IDZ', 
        'DEVEDORES DEP.5100-5           FX.549':'VLR_DPN'})
erro_saque.fillna(0, inplace=True)
erro_saque = erro_saque[erro_saque['Controle']==549]
erro_saque = erro_saque[erro_saque['MOTIVO']!='02 - CEDULA MUTILADA'][erro_saque['MOTIVO']!='04 - CEDULA SUSPEITA'][erro_saque['MOTIVO']!='05 - CEDULA ENTINTADA']
erro_saque['COD_TIP_LCT'] = [3 if x == '01 - ERROS DE SAQUE' else 6 for x in erro_saque['MOTIVO']]

erro_deposito['COD_TIP_LCT'] = 4
erro_deposito['NUM_DOC'] = 0
erro_deposito['COD_TSR'] = erro_deposito['TECE']
erro_deposito['COD_STT_LCT'] = [1 if x[15] > 0 else 2 for x in erro_deposito.values]
erro_deposito = erro_deposito.rename(columns={'DATA DO DEPÓSITO': 'DTA_LCT', 
        'COD_DEPENDENCIA':'NUM_DND' ,
        'DEMANDA':'NUM_DMD' , 
        'ATM': 'COD_TML',
        'VALOR DA DIFERENÇA' :'VLR_LCT', 
        'TECE':'IDT_CTA_IDZ', 
        'DEVEDORES DEP.5100-5           FX.6864-3    CI.716-4':'VLR_DPN'})
erro_deposito.fillna(0, inplace=True)

query_gnu_db2 = ''
with open('GNU_DB2.sql', 'r') as arquivo_sql:
    for linha in arquivo_sql:
        query_gnu_db2 += linha
    arquivo_sql.close()
gnu_db2 = pd.read_sql_query(query_gnu_db2, conn_db2).fillna(0).astype({'NUM_DMD': int})
gnu_db2.fillna(0, inplace=True)
gnu_db2['VLR_LCT'] = gnu_db2['VLR_LCT'] / 100
gnu_db2['VLR_DPN'] = gnu_db2['VLR_DPN'] / 100

# %%
#---------- AJUSTANDO LAYOUT E UNINDO OS CONTROLES ----------
lista_campos = ['DTA_LCT', 'NUM_DND', 'COD_TML', 'COD_TIP_LCT', 'VLR_LCT',
       'IDT_CTA_IDZ', 'COD_TSR', 'COD_STT_LCT', 'VLR_DPN', 'NUM_DOC',
       'NUM_DMD']

lista_campos_2 = ['DTA_LCT', 'NUM_DND', 'COD_TML', 'COD_TIP_LCT', 'VLR_LCT',
       'IDT_CTA_IDZ', 'COD_TSR', 'COD_STT_LCT', 'VLR_DPN', 'NUM_DOC',
       'NUM_DMD', 'IDT_LCT']

gnu_db2 = gnu_db2[lista_campos_2]
credores = credores[lista_campos]
devedores = devedores[lista_campos]
erro_saque = erro_saque[lista_campos]
erro_deposito = erro_deposito[lista_campos]
resumo_controle = pd.concat([credores, devedores, erro_deposito, erro_saque]).astype({
        'COD_TML': int, 'NUM_DOC': int, 'NUM_DMD': int})
resumo_controle = resumo_controle.query('NUM_DMD!=0 | NUM_DOC!=0')
resumo_controle = resumo_controle.reset_index(drop=True)


#%%
#---------- ANÁLISE BASE MANUAL E SISTEMA ----------

resumo_controle['ID'] = [x + 1000000 for x in range(len(resumo_controle))]
resumo_controle['PRESENTE_SISTEMA'] = 0
resumo_controle['VALOR_GNU'] = 0 

df1 = gnu_db2[gnu_db2['COD_TIP_LCT']==1]
df2 = gnu_db2[gnu_db2['COD_TIP_LCT']==2]
df3 = gnu_db2[gnu_db2['COD_TIP_LCT']==3]
df4 = gnu_db2[gnu_db2['COD_TIP_LCT']==4]

cont1 = cont2 = cont3 = cont4 = avaliar = 0
for registro_manual in resumo_controle.values:
     id = registro_manual[-3]
     tipo_lct = registro_manual[3]
     doc = int(registro_manual[-5])
     demanda = registro_manual[-4]
     atm = registro_manual[2]
     controle_data = registro_manual[0]
     valor_lct = registro_manual[4]
     tece = registro_manual[5]
     if tipo_lct in (1, 2) and doc in(list(globals()[f'df{tipo_lct}']['NUM_DOC'])):
        globals()[f'df{tipo_lct}_fim'] = globals()[f'df{tipo_lct}'][globals()[f'df{tipo_lct}']['NUM_DOC']==doc]
        globals()[f'df{tipo_lct}_fim'] = globals()[f'df{tipo_lct}_fim'][globals()[f'df{tipo_lct}_fim']['VLR_LCT']==valor_lct]
        globals()[f'df{tipo_lct}_fim'] = globals()[f'df{tipo_lct}_fim'][globals()[f'df{tipo_lct}_fim']['IDT_CTA_IDZ']==tece]
        globals()[f'df{tipo_lct}_fim'] = globals()[f'df{tipo_lct}_fim'][globals()[f'df{tipo_lct}_fim']['COD_TML']==atm]
        if len(globals()[f'df{tipo_lct}_fim']) == 0:
           continue
        gnu_vlr_dpn = globals()[f'df{tipo_lct}_fim'][globals()[f'df{tipo_lct}_fim']['NUM_DOC']==doc].values[0][8]
        gnu_data = globals()[f'df{tipo_lct}_fim'][globals()[f'df{tipo_lct}_fim']['NUM_DOC']==doc].values[0][0]
        gnu_id = globals()[f'df{tipo_lct}_fim'][globals()[f'df{tipo_lct}_fim']['NUM_DOC']==doc].values[0][11]
        data_dif = pd.Timestamp(gnu_data) - controle_data
        data_dif = pd.Timedelta(data_dif).days
        if -4 <= data_dif <= 4 and len(globals()[f'df{tipo_lct}_fim'])!= 0:
             globals()[f'cont{tipo_lct}'] += 1
             resumo_controle.at[resumo_controle.index[resumo_controle['ID'] == id].tolist()[0]
                     , 'PRESENTE_SISTEMA'] = 1
          
             resumo_controle.at[resumo_controle.index[resumo_controle['ID'] == id].tolist()[0]
                     , 'VALOR_GNU'] = gnu_vlr_dpn

             resumo_controle.at[resumo_controle.index[resumo_controle['ID'] == id].tolist()[0]
             , 'IDT_LCT'] = gnu_id

        if len(globals()[f'df{tipo_lct}_fim']) >= 2:
           print(f'Avaliar>> doc: {doc} - {tipo_lct}')
           avaliar += 1
           continue

     elif tipo_lct in (3, 4) and demanda in(list(globals()[f'df{tipo_lct}']['NUM_DMD'])):
        gnu_vlr_dpn = globals()[f'df{tipo_lct}'][globals()[f'df{tipo_lct}']['NUM_DMD']==demanda].values[0][8]
        gnu_id = globals()[f'df{tipo_lct}'][globals()[f'df{tipo_lct}']['NUM_DMD']==demanda].values[0][11]
        globals()[f'cont{tipo_lct}'] += 1
        resumo_controle.at[resumo_controle.index[resumo_controle['ID'] == id].tolist()[0]
                , 'PRESENTE_SISTEMA'] = 1

        resumo_controle.at[resumo_controle.index[resumo_controle['ID'] == id].tolist()[0]
                        , 'VALOR_GNU'] = gnu_vlr_dpn

        resumo_controle.at[resumo_controle.index[resumo_controle['ID'] == id].tolist()[0]
        , 'IDT_LCT'] = gnu_id

#print(f'''TIPO 1: {cont1}
#TIPO 2: {cont2}
#TIPO 3: {cont3}
#TIPO 4: {cont4}
#avaliar: {avaliar}''')

    
#%%
#---------- COLUNAS DE COMPARAÇÃO ----------
valida = list()
for registro in resumo_controle.values:
    if registro[12] == 1 and registro[13] == registro[8]:
        valida.append('OK')
    elif registro[12] == 1 and registro[13] != registro[8]:
        valida.append('NÃO OK')
    else:
        valida.append('')
resumo_controle['VALIDA'] = valida

dif = list()
for registro in resumo_controle.values:
    if registro[14] == 'NÃO OK':
        dif.append(registro[8] - registro[13])
    else:
        dif.append(0)

resumo_controle['DIF'] = dif

resumo_controle.drop(columns=['ID', 'DIF'], inplace=True)


#%%
#---------- INSERINDO RESULTADO SQLGDNP ----------
cursor.execute('TRUNCATE TABLE [MERCANTIL\B042786].[TC_DIVERGENCIAS_GNU]')

insercao = resumo_controle[resumo_controle['VALIDA']=='NÃO OK']
for index, row in insercao.iterrows():
   try:
        if row.COD_TIP_LCT == 1:
           tip_lct = 'Credores'
        elif row.COD_TIP_LCT == 2:
           tip_lct = 'Devedores'
        elif row.COD_TIP_LCT == 3:
           tip_lct = 'Erro de Saque'
        elif row.COD_TIP_LCT == 4:
           tip_lct = 'Erro de Depósito'
        else:
           tip_lct = 'Não Identificado'
        cursor.execute('''INSERT INTO [MERCANTIL\B042786].[TC_DIVERGENCIAS_GNU]
        (IDT_LCT, DTA_LCT, NUM_DND, COD_TML, TIP_LCT, VLR_LCT, COD_TSR
        ,VLR_DPN, NUM_DOC, NUM_DMD, VALOR_GNU, ULT_ATU)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,GETDATE())''', 
        row.IDT_LCT ,row.DTA_LCT, row.NUM_DND, row.COD_TML, tip_lct, row.VLR_LCT, row.COD_TSR
        ,row.VLR_DPN, row.NUM_DOC, row.NUM_DMD, row.VALOR_GNU) 
   except Exception as erro:
        print(erro)
        pass
cursor.commit()

#%%
#---------- SALVANDO ARQUIVO RESULTADO ----------
base_tratada = resumo_controle
#base_tratada.drop(columns=['PRESENTE_SISTEMA', 'VALOR_GNU', 'VALIDA'], inplace=True)
arquivo = pd.ExcelWriter(path='Base Tratada.xlsx',engine='xlsxwriter')
base_tratada.to_excel(arquivo, sheet_name="MANUAL", index=False)
gnu_db2.to_excel(arquivo, sheet_name="GNU", index=False)
arquivo.save()
arquivo.close()


# %%
#---------- FECHANDO CONEXÃO ----------
conn_db2.close()
cursor.close()
conn_gdnp.close()
