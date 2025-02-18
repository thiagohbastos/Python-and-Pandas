# %%
import os
import time
import shutil
#import getpass
import pyautogui
import numpy as np
import pandas as pd
from unidecode import unidecode
from datetime import datetime, date, timedelta

import warnings
warnings.filterwarnings('ignore')

# %%
project_folder = r"K:\GSAS\00 - Gerência\001 - Atividades e projetos da Gerência\11-2022 - VIAGENS - Relatório Gestão Mensal"

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
# Concatenar
paytrack = pd.concat([paytrack22, paytrack23, paytrack24])
print(len(paytrack))
paytrack.head()

# %%
paytrack.loc[(paytrack['Tipo despesa'] == 'Outros') &
             (
                 (paytrack['Fornecedor'].str.lower() == 'azul') |
                 (paytrack['Fornecedor'].str.lower().str.contains('azul linhas', case=False)) |
                 (paytrack['Fornecedor'].str.lower() == 'gol') |
                 (paytrack['Fornecedor'].str.lower().str.contains('gol linhas', case=False)) | 
                 (paytrack['Fornecedor'].str.lower() == 'latam') |
                 (paytrack['Justificativa'].str.lower().str.contains('despacho', case=False)) |
                 (paytrack['Justificativa'].str.lower().str.contains('bagagem', case=False))
             )
            #][['Justificativa', 'Fornecedor', 'Tipo despesa']]
             , 'Tipo despesa'] = 'Aéreo'
    
paytrack.loc[(paytrack['Tipo despesa'] == 'Outros') &
             (
                 (paytrack['Justificativa'].str.lower() == 'refeição') |
                 (paytrack['Justificativa'].str.lower().str.contains('não estou conseguindo anexar comprovantes', case=False)) |
                 (paytrack['Justificativa'].str.lower().str.contains('supermercado', case=False)) |
                 (paytrack['Justificativa'].str.lower() == 'cereais') |
                 (paytrack['Justificativa'].str.lower() == 'compra para alimentação durante o final de semana.') |
                 (paytrack['Justificativa'].str.lower() == 'água mineral')
             )
            #][['Justificativa', 'Fornecedor', 'Tipo despesa']]
             , 'Tipo despesa'] = 'Alimentação'

paytrack.loc[(paytrack['Tipo despesa'] == 'Outros') &
             (
                 (paytrack['Fornecedor'].str.lower().str.contains('jb', case=False)) |
                 (paytrack['Fornecedor'].str.lower().str.contains('localiza', case=False)) |
                 (paytrack['Fornecedor'].str.lower().str.contains('movida', case=False)) |
                 (paytrack['Justificativa'].str.lower().str.contains('carro alugado', case=False))
             )
            #][['Justificativa', 'Fornecedor', 'Tipo despesa']]
             , 'Tipo despesa'] = 'Carro'

paytrack.loc[(paytrack['Tipo despesa'] == 'Outros') &
             (
                 (paytrack['Justificativa'].str.lower().str.contains('carona', case=False)) |
                 (paytrack['Justificativa'].str.lower().str.contains('caronas', case=False))
             )
            #][['Justificativa', 'Fornecedor', 'Tipo despesa']]
             , 'Tipo despesa'] = 'Combustível-Veículo Particular'

paytrack.loc[(paytrack['Tipo despesa'] == 'Outros') &
             (
                 (paytrack['Fornecedor'].str.lower().str.contains('hotel', case=False)) |
                 (paytrack['Justificativa'].str.lower().str.contains('frigobar', case=False)) |
                 (paytrack['Justificativa'].str.lower().str.contains('hotel', case=False))
             )
            #][['Justificativa', 'Fornecedor', 'Tipo despesa']]
             , 'Tipo despesa'] = 'Hotel'

paytrack.loc[(paytrack['Tipo despesa'] == 'Outros') &
             (
                 (paytrack['Fornecedor'].str.lower().str.contains('uber', case=False)) |
                 (paytrack['Fornecedor'].str.lower().str.contains('táxi', case=False)) |
                 (paytrack['Justificativa'].str.lower().str.contains('uber', case=False))
             )
            #][['Justificativa', 'Fornecedor', 'Tipo despesa']]
             , 'Tipo despesa'] = 'Táxi'

# %%
path = project_folder + r"\Despesas_Paytrack.xlsx"

with pd.ExcelWriter(path) as writer:
    paytrack.to_excel(writer, sheet_name='Despesas Paytrack', index=False)

# %% [markdown]
# ### 1.1. JOIN

# %%
#grouped = paytrack.groupby('id').agg({'Valor':'sum', 'Data despesa':'max'})
#print(len(grouped))

base_loc = paytrack[['id', 'Valor', 'Data despesa', 'Descrição', 'Data de criação', 'Colaborador',
                     'Tipo despesa', 'Matricula', 'Número CC', 'Centro de custo', 'Projeto'
                   ]].sort_values('id').reset_index(drop=True)

#base_loc['Id Duplicado'] = base_loc['id'].duplicated(keep=False)

base_loc['Count id'] = base_loc.groupby('id')['id'].transform('count')
base_loc['Rank id'] = base_loc.groupby('id')['id'].transform('rank', method='first')

#base_loc.rename(columns={'Valor':'Valor total'}, inplace=True)
#base_loc['Valor'] = base_loc['Valor total'] / base_loc['Count id']

base_loc['Matricula'] = pd.to_numeric(base_loc['Matricula'], errors='coerce')

base_loc['Origem'] = 'Paytrack'

print(len(base_loc))
base_loc

# %% [markdown]
# ## 2. RH Cadastro

# %%
path = project_folder + r"\Bases mensais\rh_cadastro.xlsx"

rh_cadastro = pd.read_excel(path)
rh_cadastro.rename(columns={"Nome do Usuário":"Colaborador"},inplace=True)
rh_cadastro['CC Lotação'] = rh_cadastro['Centro de Custo'].astype('str') + " - " + rh_cadastro['Unidade']
rh_cadastro = rh_cadastro.sort_values('Data Admissão').drop_duplicates('Colaborador',keep='last')

print(len(rh_cadastro))
rh_cadastro.head()

# %%
base_rh = base_loc.merge(rh_cadastro[['Colaborador', 'Cargo', 'CC Lotação']], on="Colaborador",how="left")
print(len(base_rh))
base_rh.head()

# %% [markdown]
# ## 3. Segmentos

# %%
path = project_folder + r"\Bases mensais\segmentos.xlsx"

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
base_seg = base_rh.merge(segmentos[['Número CC', 'Nº CC Subordinador', 'CC Subordinador', 'Segmento', 'UF']],
                         on="Número CC",how="left")

base_seg['Nº CC Subordinador'].fillna(0, inplace=True)
base_seg.fillna("", inplace=True)
base_seg['Nº CC Subordinador'] = base_seg['Nº CC Subordinador'].astype('int64')

print(len(base_seg))
base_seg.head()

# %% [markdown]
# ## 4. Ajustes Diretoria

# %%
# Responsáveis
path = project_folder + r"\Bases mensais\responsaveis.xlsx"
centro_custo = pd.read_excel(path, usecols="A,B,C,G")
centro_custo.columns = ['Número CC', 'Centro de Custo', 'Colaborador', 'CC Subordinador']
resp = centro_custo[centro_custo['Colaborador']!=0][['Número CC', 'Colaborador']]

# %% [markdown]
# ### 4.1. 2022

# %%
# Primus

## Consolidado 1º Semestre
path1 = project_folder + r"\Faturamento\Primus 22 S1\Consolidado Geral 1º Sem. 2022.xlsx"
s1 = pd.read_excel(path1, sheet_name='Consolidado', usecols="A,B,C,F,N,R,T,U")
s1Raw = pd.read_excel(path1, sheet_name='Consolidado', usecols="A,B,C,F,N,R,T,U")
s1.columns = ['Data de criação','Passageiro','Descrição','Número CC','Valor','Data despesa','Faixa','Tipo despesa']

for i in [1125, 2758, 2829, 2878]:
    s1.iloc[i, s1.columns.get_loc('Faixa')] = 'DIR - INTERN.'

s1 = s1[(s1['Faixa']=='DIR') | (s1['Faixa']=='DIR - INTERN.') | (s1['Faixa']=='ADM EXT')].reset_index(drop=True)
s1.loc[((s1['Faixa']=='DIR - INTERN.') | (s1['Faixa']=='ADM EXT')) &
       (s1['Tipo despesa'] == 'AEREO'), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
s1.drop(columns=['Faixa'], inplace=True)

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
s2Raw = pd.concat(to_append).reset_index(drop=True)

s2 = s2[['data', 'passageiro', 'fornecedor', 'cc', 'destino', 'total', 'tipo', 'Referência', 'faixa/ cargo']]
s2.columns = ['Data de criação', 'Passageiro', 'Descrição', 'Número CC',
              'Destino', 'Valor', 'Tipo despesa', 'Data despesa', 'Faixa']
s2.loc[(s2['Faixa'].str.contains('EXT')) & (s2['Tipo despesa'] == 'AEREO'), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
s2.loc[(s2['Destino'].str.contains('ATL')) | (s2['Destino'].str.contains('DOH')), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
s2['Descrição'] = s2['Descrição'] + " - " + s2['Destino']
s2.drop(columns=['Faixa', 'Destino'], inplace=True)

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
ebta22Raw = pd.concat(to_append).reset_index(drop=True)

ebta22.columns = ['Data de criação', 'Descrição', 'Número CC', 'Valor', 'faixa', 'Passageiro', 'Data despesa']
ebta22['Tipo despesa'] = np.where(ebta22['faixa'].str.contains('EXT'), 'AEREO INTERNACIONAL', 'AEREO')
ebta22.loc[(ebta22['Descrição'].str.contains('SFO')) | 
           (ebta22['Descrição'].str.contains('DOH')) |
           (ebta22['Descrição'].str.contains('IST')), 'Tipo despesa'] = 'AEREO INTERNACIONAL'
ebta22.drop(columns=['faixa'], inplace=True)

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
# ### 4.2. 2023

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
ebta23Raw = pd.concat(to_append).reset_index(drop=True)

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
base_dir['Tipo despesa'].replace({
    'AEREO': 'Aéreo', 'AEREO INTERNACIONAL': 'Aéreo internacional','HOTEL': 'Hotel', 'CARRO':'Carro', 'PLANTÃO':'Plantão',
    'SEGURO DE VIAGEM':'Seguro', 'SEGURO':'Seguro', 'SEGURO VIAGEM':'Seguro', 'GESTÃO':'Gestão', 'TAXA':'Taxa'}, 
    inplace=True)

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

# %% [markdown]
# ### 5.4 Salvar sem violação

# %%
path = project_folder + r"\Gestão de Viagens Despesas.xlsx"

final_df = base_orc.drop(columns=['Data de criação', 'Matricula', 'Count id', 'Rank id', 'Valor total'])

with pd.ExcelWriter(path) as writer:
    final_df.to_excel(writer, sheet_name='Despesas Paytrack', index=False)
    orcamento_mes_full.to_excel(writer, sheet_name='Orçamento', index=False)
    rh_cadastro.to_excel(writer, sheet_name='RH Cadastro', index=False)

# %% [markdown]
# ## 6. Copiar arquivos

# %%
file_to_copy = r"K:\GSAS\00 - Gerência\001 - Atividades e projetos da Gerência\11-2022 - VIAGENS - Relatório Gestão Mensal\Gestão de Viagens Despesas.xlsx"
destination_directory = r"T:\SIC\Seg Patrimonial\Manutenções\Viagens\Gestão de Viagens Despesas.xlsx"
shutil.copy(file_to_copy, destination_directory)

pyautogui.alert('O código foi finalizado. Você já pode utilizar o computador!')


