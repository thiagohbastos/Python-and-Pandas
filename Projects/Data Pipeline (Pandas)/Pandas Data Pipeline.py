#%%
# ------------- PACOTES -------------
#<PASTA DO PROJETO>\venv\Scripts\activate
#pyinstaller --name="AUTOMACAO PROGRAMACAO" --onefile Pregera_plans.py
import pandas as pd
import xlsxwriter
import openpyxl
import pyodbc
import warnings
from datetime import date
import datetime
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL
from time import sleep
warnings.filterwarnings('ignore')


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
# ------------- FUNÇÃO DE CRIAR QUERIES -------------
nome_queries = ['STATUS_CARD.sql'
    ,'CUSTODIA_COMPLETA.sql'
    ,'PROGRAMADO_GNU.sql'
    ,'ATM_INFO.sql'
    ,'COMPOSICAO.sql'
    ,'SALDOS_TRATADOS.sql'
    ,'SUGESTAO.sql'
    ,'ALTERACAO_MANUAL.sql'
    ,'PERCENTUAL_CED_ATM.sql'
    ]

def cria_query(caminho_query:str):
    query = ''
    with open(caminho_query, 'r', encoding='utf-8') as x:
        for linha_atm in x:
            query += linha_atm
    x.close()
    return query

def remove_repetidos(lista):
    l = []
    for i in lista:
        if i not in l:
            l.append(i)
    l.sort()
    return l

for x in nome_queries:
    query_nome = 'query_' + x.lower().replace('.sql', '')
    globals()[f'{query_nome}'] = cria_query(x)

agora = datetime.datetime.now()
ano = agora.year
mes = str(agora.month) if len(str(agora.month)) == 2 else f'0{agora.month}' 
dia = str(agora.day) if len(str(agora.day)) == 2 else f'0{agora.day}' 
hora = str(agora.hour) if len(str(agora.hour)) == 2 else f'0{agora.hour}'
minuto = str(agora.minute) if len(str(agora.minute)) == 2 else f'0{agora.minute}'


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
engine = create_engine(connection_url
                       ,fast_executemany=True
                       ,connect_args={'connect_timeout': 10}
                       , echo=False)
conn_alchemy_gnu = engine.connect()


#%%
# ------------- CRIANDO ARQUIVO DE RESUMO DO PROCESSO -------------
caminho_resumo = f'K:/GSAS/09 - Coordenacao Gestao Numerario/09 Prototipos SSIS/B042786/__PYTHON__/__EXECUTAVEIS__/PROGRAMACAO AUT/RESUMOS/Resumo ({date.today().day}-{date.today().month}-{date.today().year}).txt'
with open (caminho_resumo, 'w') as resumo:
    resumo.write('-------- INICIANDO PROCESSO DE AUTOMAÇÃO DA PROGRAMAÇÃO --------\n\n')
    resumo.close()
print('-------- INICIANDO PROCESSO DE AUTOMAÇÃO DA PROGRAMAÇÃO --------\n\n')


#%%
# ------------- TABELA DE STATUS CARD -------------
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nCARREGANDO TABELA DE STATUS CARD (INDICADORES DE CUSTÓDIA)\n...')
    resumo.close()
print('\nCARREGANDO TABELA DE STATUS CARD (INDICADORES DE CUSTÓDIA)...')
status_card = pd.read_sql_query(query_status_card, conn_gnu)[['TECE', 'AJUSTE', 'DIMINUI_TROCO']]


#%%
# ------------- TABELA DE PERCENTUAL DE ABASTECIMENTO -------------
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nCARREGANDO VALORES IDEAIS DE CUSTÓDIA\n...')
    resumo.close()
print('\nCARREGANDO VALORES IDEAIS DE CUSTÓDIA...')
perc_ced = pd.read_sql_query(query_percentual_ced_atm, conn_gnu)


#%%
# ------------- TABELA DE ABASTECIMENTO PROGRAMADO -------------
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nCARREGANDO TABELA DE ABASTECIMENTO PROGRAMADO/SUGERIDO')
    resumo.close()
print('\nCARREGANDO TABELA DE ABASTECIMENTO PROGRAMADO/SUGERIDO...')
info_atm = pd.read_sql_query(query_atm_info, conn_gnu)
abastec_prog = pd.read_sql_query(query_programado_gnu, conn_mbcorp)
abastec_prog = abastec_prog.merge(info_atm, how='left', on='IDT_TML')[
    ['DTA_HOR_PRG', 'TECE', 'IDT_TML', 'NUM_DND', 'VLR_ABT_PRG']
    ].sort_values(by=['TECE', 'VLR_ABT_PRG', 'IDT_TML'], 
    ignore_index=True)
usando_prog = 1

if len(abastec_prog) == 0:
    with open (caminho_resumo, 'a') as resumo:
        resumo.write('Utilizando as sugestões para alteração\n...')
        resumo.close()
    print('Utilizando as sugestões para alteração.')
    abastec_prog = pd.read_sql_query(query_sugestao, conn_gnu)
    usando_prog = 0
else:
    with open (caminho_resumo, 'a') as resumo:
        resumo.write('Utilizando os terminais inseridos no GNU\n...')
        resumo.close()
    print('Utilizando os terminais inseridos no GNU.')


#%%
# -------- AVALIANDO VALORES CADASTRADOS NO GNU --------
composicoes = pd.read_sql_query(query_composicao, conn_gnu)
interromper = 0
if usando_prog == 1:
    for x in abastec_prog.values:
        valor = x[4]
        atm = x[2]
        dnd = x[3]
        if valor not in(composicoes['VLR_ABT_PRG'].values):
            print(f'ERRO EM VALOR CADASTRADO: ATM {atm}, DND {dnd}, VALOR R${valor}')
            interromper = 1
            with open (caminho_resumo, 'a') as resumo:
                resumo.write(f'\nERRO EM VALOR CADASTRADO: ATM {atm}, DND {dnd}, VALOR R${valor}...')
                resumo.close()

    if interromper == 1:
        with open (caminho_resumo, 'a') as resumo:
            resumo.write(f'\nO programa será encerrado em 10 segundos devidos erros de perfis não cadastrados citados acima...')
            resumo.close()
        print('O programa será encerrado em 10 segundos devidos erros de perfis não cadastrados citados acima...')
        sleep(10)
        conn_gnu.close()
        conn_mbcorp.close()
        exit()


#%%
# -------- TABELA DE INDICAÇÃO DE AJUSTES COMPLETA --------
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nCARREGANDO TABELA DE INDICAÇÃO DE AJUSTES NECESSÁRIOS\n...')
    resumo.close()
print('\nCARREGANDO TABELA DE INDICAÇÃO DE AJUSTES NECESSÁRIOS...')
abastec_prog = abastec_prog.merge(composicoes, how='left', on='VLR_ABT_PRG').rename(
    columns={'COD_CEN': 'TECE'})
abastec_prog = abastec_prog.merge(status_card, how='left', on='TECE')
abastec_prog['VLR_ABT_ORI'] = abastec_prog['VLR_ABT_PRG']
abastec_prog = abastec_prog[['DTA_HOR_PRG', 'TECE', 'IDT_TML', 'NUM_DND', 'VLR_ABT_ORI', 'VLR_ABT_PRG', 'COMP',
       'CED_2', 'CED_5', 'CED_10', 'CED_20', 'CED_50', 'CED_100', 'CED_200',
       'AJUSTE', 'DIMINUI_TROCO']]


#%%
#ALTERANDO VALORES DE ALTERAÇÃO MANUAL
alteracoes_manuais = pd.read_sql_query(query_alteracao_manual, conn_gnu)
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nAVALIANDO ALTERAÇÕES MANUAIS')
    if len(alteracoes_manuais) == 0:
        resumo.write('NÃO HÁ ALTERAÇÕES MANUAIS\n...')
    else:
        resumo.write('ALTERAÇÕES MANUAIS ENCONTRADAS E CONSIDERADAS.\nGENTILEZA AVALIAR\n...')
    resumo.close()
print('\nAVALIANDO ALTERAÇÕES MANUAIS...')
sleep(1.5)
if len(alteracoes_manuais) == 0:
    print('NÃO HÁ ALTERAÇÕES MANUAIS...')
else:
    print('ALTERAÇÕES MANUAIS ENCONTRADAS E CONSIDERADAS.\nGENTILEZA AVALIAR...')

for registro in alteracoes_manuais.values:
    atm = registro[2]
    atm_ced2 = registro[3]
    atm_ced5 = registro[4]
    atm_ced10 = registro[5]
    atm_ced20 = registro[6]
    atm_ced50 = registro[7]
    atm_ced100 = registro[8]
    atm_ced200 = registro[9]
    atm_valor = registro[11]
    atm_comp = registro[12]

    index = abastec_prog.index[abastec_prog['IDT_TML'] == atm].tolist()
    for caso in index:
        abastec_prog.at[caso, 'CED_2'] = atm_ced2
        abastec_prog.at[caso, 'CED_5'] = atm_ced5
        abastec_prog.at[caso, 'CED_10'] = atm_ced10
        abastec_prog.at[caso, 'CED_20'] = atm_ced20
        abastec_prog.at[caso, 'CED_50'] = atm_ced50
        abastec_prog.at[caso, 'CED_100'] = atm_ced100
        abastec_prog.at[caso, 'CED_200'] = atm_ced200
        abastec_prog.at[caso, 'VLR_ABT_ORI'] = atm_valor
        abastec_prog.at[caso, 'VLR_ABT_PRG'] = atm_valor
        abastec_prog.at[caso, 'COMP'] = atm_comp
        break

alterando_atm = abastec_prog.copy()
alterando_atm['ALTERADO'] = 0


#%%
# ------------- TABELA DE SALDOS TRATADOS -------------
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nCARREGANDO VALORES DE CUSTÓDIA E RECOLHIMENTO DISPONÍVEIS\n...\n\n')
    resumo.close()
print('\nCARREGANDO VALORES DE CUSTÓDIA E RECOLHIMENTO DISPONÍVEIS...')
saldos_tratados = pd.read_sql_query(query_saldos_tratados, conn_gnu)
saldos_tratados = saldos_tratados.merge(status_card, how='left', on='TECE')

alterando_saldo = saldos_tratados.copy()
alterando_saldo = saldos_tratados.reset_index(drop=True)
alterando_saldo['VALOR_ORIGINAL'] = alterando_saldo['VALOR_TOTAL']
alterando_saldo = alterando_saldo[['TECE', 'TIPO_NUMERARIO', 'CEDULA', 'VALOR_TOTAL', 'VALOR_ORIGINAL'
    , 'AJUSTE', 'DIMINUI_TROCO']]
alterando_saldo['AJUSTE_REALIZADO'] = 0


#%%
# ------------- REGRAS DOS K7s E SEUS VALORES -------------
# Para cada cédula temos uma lista com:
# Primeiro elemento: salto de alteração do k7 da cédula.
# Segundo elemento: valor máximo para a cédula.
# Terceiro elemento: 
config_ced = {2:[1000, 4000, 1000]
    ,5:[2500, 10000, 2500]
    ,10:[5000, 20000, 5000]
    ,20:[10000, 40000, 10000]
    ,50:[5000, 100000, 5000]
    ,100:[10000, 210000, 10000]
    ,200:[20000, 400000, 20000]}



#%%
# ------------- TRATANDO CED GERAIS -------------
cont = 0
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\n----- Iniciando o primeiro processo de adaptação de custódia -----\n')
    resumo.close()
print('\nIniciando o primeiro processo de adaptação de custódia')
for x in status_card[status_card['AJUSTE'] == 1].values:
    stop = 0
    tece = x[0]
    ceds_negativas = alterando_saldo[alterando_saldo['VALOR_TOTAL'] < 0][
        alterando_saldo['TECE'] == tece].sort_values(by = 'VALOR_TOTAL')[['CEDULA']].values
    ceds_negativas = [int(x[0]) for x in ceds_negativas.tolist()]

    ceds_posi = alterando_saldo[alterando_saldo['VALOR_TOTAL'] > 0][
        alterando_saldo['TECE'] == tece].sort_values(
            by = 'VALOR_TOTAL', ascending=False)[['CEDULA']].values
    ceds_posi = [int(x[0]) for x in ceds_posi.tolist()]
    with open (caminho_resumo, 'a') as resumo:
        resumo.write(f'------ Iniciando TECE {tece} ------\n')
        resumo.close()
    print(f'------ Iniciando TECE {tece} ------')
    for ced_neg in ceds_negativas:
        if stop == 1:
            stop = 0
            continue
        
        atms = alterando_atm[alterando_atm['TECE'] == tece][
            alterando_atm[f'CED_{ced_neg}'] > 0]
        
        for atm in atms.values:
            if stop == 1:
                break
            tml = atm[2]
            dta_prog = atm[0]
            composicao = [int(z.strip()) for z in atm[6].split('-')]
           
            config_ced_temp = {}
            for k, v in config_ced.items():
                config_ced_temp[k] = v
            config_ced_temp[ced_neg] = [config_ced_temp[ced_neg][0], config_ced_temp[ced_neg][1] * composicao.count(ced_neg)
                                        , config_ced_temp[ced_neg][2] * composicao.count(ced_neg)]

            linha_atm = [x for x in alterando_atm[
                alterando_atm['IDT_TML'] == tml][
                alterando_atm['DTA_HOR_PRG'] == dta_prog
                ].index][0]

            ceds_posi_atm = [x for x in composicao if x in ceds_posi]

            usadas = []
            for ced_posi in ceds_posi_atm:
                if stop == 1:
                    break

                ocorrencias = ceds_posi_atm.count(ced_posi)
                if ocorrencias >= 2 and ced_posi not in usadas:
                    config_ced_temp[ced_posi] = [config_ced_temp[ced_posi][0], 
                        config_ced_temp[ced_posi][1] * ocorrencias]
                    usadas.append(ced_posi)
                while True:
                    #Quantidade de alteração da ced negativa >= da ced positiva
                    parametro_alt = max([config_ced_temp[ced_neg][0], config_ced_temp[ced_posi][0]])

                    #Validando se ao alterar o k7 com cédula negativa o valor do k7 extrapola o máximo
                    valida_1 = alterando_atm.loc[linha_atm][f'CED_{ced_posi}'] + parametro_alt <= config_ced_temp[ced_posi][1]
                    
                    #Validando se o saldo da cédula na TECE permanece negativo
                    valida_2 = alterando_saldo[alterando_saldo['TECE'] == tece][
                        alterando_saldo['CEDULA'] == ced_neg].values[0][3] < 0 and alterando_saldo[alterando_saldo['TECE'] == tece][
                        alterando_saldo['CEDULA'] == ced_posi].values[0][3] > 0

                    #Valida se é possível tirar saldo da cédula positiva
                    valida_3 = alterando_saldo[alterando_saldo['TECE'] == tece][
                        alterando_saldo['CEDULA'] == ced_posi].values[0][3] - parametro_alt >= 0

                    #Valida se a alteração vai zerar o k7 da ced positiva
                    valida_4 = alterando_atm.loc[linha_atm][f'CED_{ced_neg}'] - parametro_alt > 0

                    valida_5 = alterando_atm.loc[linha_atm][f'CED_{ced_neg}'] - parametro_alt >= config_ced_temp[ced_neg][2]

                    if not valida_2:
                        stop = 1
                        break
                    
                    if  valida_1 and valida_2 and valida_3 and valida_4 and valida_5:

                        linha_saldo_posi = alterando_saldo[alterando_saldo['TECE'] == tece][
                            alterando_saldo['CEDULA'] == ced_posi].index[0]

                        linha_saldo_neg = alterando_saldo[alterando_saldo['TECE'] == tece][
                            alterando_saldo['CEDULA'] == ced_neg].index[0]

                        alterando_atm.at[linha_atm, f'CED_{ced_neg}'] -= parametro_alt
                        alterando_atm.at[linha_atm, f'CED_{ced_posi}'] += parametro_alt
                        alterando_atm.at[linha_atm, 'ALTERADO'] = 1
                        alterando_saldo.at[linha_saldo_posi, 'VALOR_TOTAL'] -= parametro_alt
                        alterando_saldo.at[linha_saldo_neg, 'VALOR_TOTAL'] += parametro_alt

                        if alterando_saldo[alterando_saldo['TECE'] == tece][alterando_saldo['CEDULA'] == ced_neg].values[0][3] > 0:
                            alterando_saldo.at[linha_saldo_neg, 'AJUSTE_REALIZADO'] = 1
                            break
                    else:
                        break
qtd_alterados_1 = sum(alterando_atm['ALTERADO'])
with open (caminho_resumo, 'a') as resumo:
    resumo.write(f'Ao fim da etapa foram alterados {qtd_alterados_1} atms.\n\n')
    resumo.close()
print(f'Ao fim da etapa foram alterados {qtd_alterados_1} atms.')
alterando_saldo['VALOR_AJUSTE'] = alterando_saldo['VALOR_TOTAL'] - alterando_saldo['VALOR_ORIGINAL']


#%%
# ------------- TRATANDO CED TROCO (R$2) -------------
if usando_prog == 1:
    for x in status_card[status_card['DIMINUI_TROCO'] == 1].values:
        tece = x[0]
        ajuste = x[1]
        diminuir = x[2]

        linhas = abastec_prog[abastec_prog['TECE']== tece][
            abastec_prog['DIMINUI_TROCO']==1][
            abastec_prog['CED_2'] > 0].index

        cont = 0
        while True:
            linha_ced2 = alterando_saldo[alterando_saldo['CEDULA'] == 2][
                alterando_saldo['TECE']== tece]
            index_saldo = linha_ced2.index[0]
            saldo_ced2 = linha_ced2.values[0][3]
            for z in linhas:
                alterando_atm.at[z, 'CED_2'] -= 1000
                alterando_atm.at[z, 'VLR_ABT_PRG'] -= 1000
                alterando_atm.at[z, 'ALTERADO'] = 1
                alterando_saldo.at[index_saldo, 'VALOR_TOTAL'] += 1000

                linha_ced2 = alterando_saldo[alterando_saldo['CEDULA'] == 2][
                alterando_saldo['TECE']== tece]
                row_ced_2 = linha_ced2.index[0]
                saldo_ced2 = linha_ced2.values[0][3]

                if saldo_ced2 >= 0:
                    alterando_saldo.at[row_ced_2, 'AJUSTE_REALIZADO'] = 1
                    break
            cont += 1
            if cont >= 4 or saldo_ced2 >= 0:
                break


#%%
# ------------- OTIMIZANDO CUSTÓDIA -------------
metas = {}
cedulas_base = [10, 20, 50, 100, 200]
#if usando_prog == 0:
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\n----- Iniciando processo de otimização de custódia ----- \n')
    resumo.close()
print('\nIniciando processo de otimização de custódia')
for x in perc_ced.values:
    stop = 0
    if len(alterando_atm[alterando_atm['TECE'] == x[0]]) == 0:
        continue

    tece = x[0]
    with open (caminho_resumo, 'a') as resumo:
        resumo.write(f'------ Iniciando TECE {tece} ------\n')
        resumo.close()
    print(f'------ Iniciando TECE {tece} ------')

    try:
        ced_10_atual = int(saldos_tratados[saldos_tratados['TECE'] == tece][
            saldos_tratados['CEDULA'] == 10]['VALOR_TOTAL'])
    except:
        ced_10_atual = 0
    ced_10_meta = x[2]

    try:
        ced_20_atual = int(saldos_tratados[saldos_tratados['TECE'] == tece][
            saldos_tratados['CEDULA'] == 20]['VALOR_TOTAL'])
    except:
        ced_20_atual = 0
    ced_20_meta = x[3]

    try:
        ced_50_atual = int(saldos_tratados[saldos_tratados['TECE'] == tece][
            saldos_tratados['CEDULA'] == 50]['VALOR_TOTAL'])
    except:
        ced_50_atual = 0
    ced_50_meta = x[4]

    try:
        ced_100_atual = int(saldos_tratados[saldos_tratados['TECE'] == tece][
            saldos_tratados['CEDULA'] == 100]['VALOR_TOTAL'])
    except:
        ced_100_atual = 0
    ced_100_meta = x[5]

    try:
        ced_200_atual = int(saldos_tratados[saldos_tratados['TECE'] == tece][
            saldos_tratados['CEDULA'] == 200]['VALOR_TOTAL'])
    except:
        ced_200_atual = 0
    ced_200_meta = x[6]

    base = {'indicador': ['ATUAL', 'META', 'PERC', 'FOLGA']
        ,10: [ced_10_atual, ced_10_meta
            , 0 if ced_10_meta == 0 else (ced_10_atual/ced_10_meta)
            , ced_10_atual - ced_10_meta]
        ,20: [ced_20_atual, ced_20_meta
            , 0 if ced_20_meta == 0 else (ced_20_atual/ced_20_meta)
            , ced_20_atual - ced_20_meta]
        ,50: [ced_50_atual, ced_50_meta
            , 0 if ced_50_meta == 0 else (ced_50_atual/ced_50_meta)
            , ced_50_atual - ced_50_meta]
        ,100: [ced_100_atual, ced_100_meta
            , 0 if ced_100_meta == 0 else (ced_100_atual/ced_100_meta)
            , ced_100_atual - ced_100_meta]
        ,200: [ced_200_atual, ced_200_meta
            , 0 if ced_200_meta == 0 else (ced_200_atual/ced_200_meta)
            , ced_200_atual - ced_200_meta]
        }
    
    base = pd.DataFrame(base).fillna(0).transpose()
    base = base.rename(columns=base.iloc[0])
    base = base.drop(base.index[0])
    #base['FOLGA'] = base['FOLGA'] * 0.95
    base = base.reset_index()
    base.rename(inplace=True, columns={'index':'CEDULA'})
    metas_temp = base
    metas_temp['TECE'] = tece

    if len(metas) == 0:
        metas = metas_temp
    else:
        metas = pd.concat([metas, metas_temp])

    atms_tece = alterando_atm[alterando_atm['TECE'] == tece]

    ceds_posi = list(base[base['FOLGA'] > 0]['CEDULA'])
    ceds_negativas = list(base[base['FOLGA'] < 0]['CEDULA'])
    
    for ced_neg in ceds_negativas:
        if stop == 1:
            stop = 0
            continue
        
        atms = alterando_atm[alterando_atm['TECE'] == tece][
            alterando_atm[f'CED_{ced_neg}'] > 0]
        
        for atm in atms.values:
            if stop == 1:
                break
            tml = atm[2]
            dta_prog = atm[0]
            composicao = [int(z.strip()) for z in atm[6].split('-')]
        
            config_ced_temp = {}
            for k, v in config_ced.items():
                config_ced_temp[k] = v
            config_ced_temp[ced_neg] = [config_ced_temp[ced_neg][0], config_ced_temp[ced_neg][1] * composicao.count(ced_neg)
                                        ,config_ced_temp[ced_neg][2] * composicao.count(ced_neg)]

            linha_atm = [x for x in alterando_atm[
                alterando_atm['IDT_TML'] == tml][
                alterando_atm['DTA_HOR_PRG'] == dta_prog
                ].index][0]

            ceds_posi_atm = [x for x in composicao if x in ceds_posi]

            usadas = []
            for ced_posi in ceds_posi_atm:
                if stop == 1:
                    break

                ocorrencias = ceds_posi_atm.count(ced_posi)
                if ocorrencias >= 2 and ced_posi not in usadas:
                    config_ced_temp[ced_posi] = [config_ced_temp[ced_posi][0], 
                        config_ced_temp[ced_posi][1] * ocorrencias]
                    usadas.append(ced_posi)
                    
                while True:
                    #Quantidade de alteração da ced negativa >= da ced positiva
                    parametro_alt = max([config_ced_temp[ced_neg][0], config_ced_temp[ced_posi][0]])

                    #Validando se ao alterar o k7 com cédula negativa o valor do k7 extrapola o máximo
                    valida_1 = alterando_atm.loc[linha_atm][f'CED_{ced_posi}'] + parametro_alt <= config_ced_temp[ced_posi][1]
                    
                    #Validando se o saldo da cédula na TECE permanece negativo
                    valida_2 = base[base['CEDULA'] == ced_neg].values[0][4] < 0 and base[base['CEDULA'] == ced_posi].values[0][4] > 0

                    #Valida se é possível tirar saldo da cédula positiva
                    valida_3 = base[base['CEDULA'] == ced_posi].values[0][4] - parametro_alt >= 0

                    #Valida se a alteração vai zerar o k7 da ced positiva
                    valida_4 = alterando_atm.loc[linha_atm][f'CED_{ced_neg}'] - parametro_alt > 0

                    valida_5 = alterando_atm.loc[linha_atm][f'CED_{ced_neg}'] - parametro_alt >= config_ced_temp[ced_neg][2]

                    if not valida_2:
                        stop = 1
                        break
                    
                    if  valida_1 and valida_2 and valida_3 and valida_4 and valida_5:
                        linha_saldo_posi = alterando_saldo[alterando_saldo['TECE'] == tece][
                            alterando_saldo['CEDULA'] == ced_posi].index[0]

                        linha_saldo_neg = alterando_saldo[alterando_saldo['TECE'] == tece][
                            alterando_saldo['CEDULA'] == ced_neg].index[0]

                        linha_folga_posi = base[base['CEDULA'] == ced_posi].index[0]

                        linha_folga_neg = base[base['CEDULA'] == ced_neg].index[0]

                        alterando_atm.at[linha_atm, f'CED_{ced_neg}'] -= parametro_alt
                        alterando_atm.at[linha_atm, f'CED_{ced_posi}'] += parametro_alt
                        alterando_atm.at[linha_atm, 'ALTERADO'] = 1
                        alterando_saldo.at[linha_saldo_posi, 'VALOR_TOTAL'] -= parametro_alt
                        alterando_saldo.at[linha_saldo_neg, 'VALOR_TOTAL'] += parametro_alt

                        base.at[linha_folga_posi, 'FOLGA'] -= parametro_alt
                        base.at[linha_folga_neg, 'FOLGA'] += parametro_alt
                    else:
                        break

qtd_alterados_2 = sum(alterando_atm['ALTERADO'])
with open (caminho_resumo, 'a') as resumo:
    resumo.write(f'Ao fim da etapa foram alterados {qtd_alterados_2 - qtd_alterados_1} atms\n\n')
    resumo.close()
print(f'Ao fim da etapa foram alterados {qtd_alterados_2 - qtd_alterados_1} atms.')


#%%
# ------------- TRATANDO CED GERAIS -------------
cont = 0
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\n----- Iniciando o segundo processo de adaptação de custódia (possíveis correções) -----\n')
    resumo.close()
print('\nIniciando o segundo processo de adaptação de custódia (possíveis correções)')
for x in status_card[status_card['AJUSTE'] == 1].values:
    stop = 0
    tece = x[0]
    ceds_negativas = alterando_saldo[alterando_saldo['VALOR_TOTAL'] < 0][
        alterando_saldo['TECE'] == tece].sort_values(by = 'VALOR_TOTAL')[['CEDULA']].values
    ceds_negativas = [int(x[0]) for x in ceds_negativas.tolist()]

    ceds_posi = alterando_saldo[alterando_saldo['VALOR_TOTAL'] > 0][
        alterando_saldo['TECE'] == tece].sort_values(
            by = 'VALOR_TOTAL', ascending=False)[['CEDULA']].values
    ceds_posi = [int(x[0]) for x in ceds_posi.tolist()]

    with open (caminho_resumo, 'a') as resumo:
        resumo.write(f'------ Iniciando TECE {tece} ------\n')
        resumo.close()
    print(f'------ Iniciando TECE {tece} ------')
    for ced_neg in ceds_negativas:
        if stop == 1:
            stop = 0
            continue
        
        atms = alterando_atm[alterando_atm['TECE'] == tece][
            alterando_atm[f'CED_{ced_neg}'] > 0]
        
        for atm in atms.values:
            if stop == 1:
                break
            tml = atm[2]
            dta_prog = atm[0]
            composicao = [int(z.strip()) for z in atm[6].split('-')]
           
            config_ced_temp = {}
            for k, v in config_ced.items():
                config_ced_temp[k] = v
            config_ced_temp[ced_neg] = [config_ced_temp[ced_neg][0], config_ced_temp[ced_neg][1] * composicao.count(ced_neg)
                                        ,config_ced_temp[ced_neg][2] * composicao.count(ced_neg)]

            linha_atm = [x for x in alterando_atm[
                alterando_atm['IDT_TML'] == tml][
                alterando_atm['DTA_HOR_PRG'] == dta_prog
                ].index][0]

            ceds_posi_atm = [x for x in composicao if x in ceds_posi]

            usadas = []
            for ced_posi in ceds_posi_atm:
                if stop == 1:
                    break

                ocorrencias = ceds_posi_atm.count(ced_posi)
                if ocorrencias >= 2 and ced_posi not in usadas:
                    config_ced_temp[ced_posi] = [config_ced_temp[ced_posi][0], 
                        config_ced_temp[ced_posi][1] * ocorrencias]
                    usadas.append(ced_posi)
                while True:
                    #Quantidade de alteração da ced negativa >= da ced positiva
                    parametro_alt = max([config_ced_temp[ced_neg][0], config_ced_temp[ced_posi][0]])

                    #Validando se ao alterar o k7 com cédula negativa o valor do k7 extrapola o máximo
                    valida_1 = alterando_atm.loc[linha_atm][f'CED_{ced_posi}'] + parametro_alt <= config_ced_temp[ced_posi][1]
                    
                    #Validando se o saldo da cédula na TECE permanece negativo
                    valida_2 = alterando_saldo[alterando_saldo['TECE'] == tece][
                        alterando_saldo['CEDULA'] == ced_neg].values[0][3] < 0 and alterando_saldo[alterando_saldo['TECE'] == tece][
                        alterando_saldo['CEDULA'] == ced_posi].values[0][3] > 0

                    #Valida se é possível tirar saldo da cédula positiva
                    valida_3 = alterando_saldo[alterando_saldo['TECE'] == tece][
                        alterando_saldo['CEDULA'] == ced_posi].values[0][3] - parametro_alt >= 0

                    #Valida se a alteração vai zerar o k7 da ced positiva
                    valida_4 = alterando_atm.loc[linha_atm][f'CED_{ced_neg}'] - parametro_alt > 0

                    valida_5 = alterando_atm.loc[linha_atm][f'CED_{ced_neg}'] - parametro_alt >= config_ced_temp[ced_neg][2]

                    if not valida_2:
                        stop = 1
                        break
                    
                    if  valida_1 and valida_2 and valida_3 and valida_4 and valida_5:
                        linha_saldo_posi = alterando_saldo[alterando_saldo['TECE'] == tece][
                            alterando_saldo['CEDULA'] == ced_posi].index[0]

                        linha_saldo_neg = alterando_saldo[alterando_saldo['TECE'] == tece][
                            alterando_saldo['CEDULA'] == ced_neg].index[0]

                        alterando_atm.at[linha_atm, f'CED_{ced_neg}'] -= parametro_alt
                        alterando_atm.at[linha_atm, f'CED_{ced_posi}'] += parametro_alt
                        alterando_atm.at[linha_atm, 'ALTERADO'] = 1
                        alterando_saldo.at[linha_saldo_posi, 'VALOR_TOTAL'] -= parametro_alt
                        alterando_saldo.at[linha_saldo_neg, 'VALOR_TOTAL'] += parametro_alt

                        if alterando_saldo[alterando_saldo['TECE'] == tece][alterando_saldo['CEDULA'] == ced_neg].values[0][3] > 0:
                            alterando_saldo.at[linha_saldo_neg, 'AJUSTE_REALIZADO'] = 1
                            break
                    else:
                        break

qtd_alterados_3 = sum(alterando_atm['ALTERADO'])
with open (caminho_resumo, 'a') as resumo:
    resumo.write(f'Ao fim da etapa foram alterados {qtd_alterados_3 - qtd_alterados_2} atms.\n\n')
    resumo.write(f'\n------- No total, foram alterados {qtd_alterados_3} atms -------\n\n')
    resumo.close()
print(f'Ao fim da etapa foram alterados {qtd_alterados_3 - qtd_alterados_2} atms.')
print(f'\nNo total, foram alterados {qtd_alterados_3} atms.')
alterando_saldo['VALOR_AJUSTE'] = alterando_saldo['VALOR_TOTAL'] - alterando_saldo['VALOR_ORIGINAL']


#%%
#INSERINDO ALTERAÇÕES DE CUSTÓDIA NA TABELA [MERCANTIL\B042786].[TI_GESTAO_ALT_CUSTODIA]
cursor_gnu.execute('TRUNCATE TABLE [MERCANTIL\B042786].[TI_GESTAO_ALT_CUSTODIA]')
cursor_gnu.commit()

with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nIniciando as inserções das alterações na base de custódia\n...')
    resumo.close()
print('\nIniciando as inserções das alterações da base de custódia...')

result_alterando_saldo = alterando_saldo.copy()
result_alterando_saldo['BANCO'] = '389 - BMB'
result_alterando_saldo['VLR_CEDULA_MOEDA'] = result_alterando_saldo['CEDULA']
result_alterando_saldo['DESCRICAO'] = ''
result_alterando_saldo['TRANSACAO'] = ''
result_alterando_saldo['ATUALIZACAO'] = f'{ano}-{mes}-{dia} {hora}:{minuto}:00.000'
result_alterando_saldo['OS'] = None
result_alterando_saldo['IDC_IBS'] = None
result_alterando_saldo['IDT_DTN_TEB'] = None
result_alterando_saldo['DTA_REG_OPR'] = None
result_alterando_saldo['DTA_PRG'] = None
result_alterando_saldo['PROTOCOLO'] = None

# 'Programação ATM (-)', 'Programação'
df1 = result_alterando_saldo[result_alterando_saldo['VALOR_AJUSTE'] != 0]
df1['DESCRICAO'] = 'Programação ATM (-)'
df1['TRANSACAO'] = 'Programação'

#'Saldo', 'Saldo'
df2 = result_alterando_saldo[result_alterando_saldo['VALOR_AJUSTE'] != 0]
df2['DESCRICAO'] = 'Saldo'
df2['TRANSACAO'] = 'Saldo'

#'Saldo Programação', 'Saldo Programação'
df3 = result_alterando_saldo.copy()
df3['DESCRICAO'] = 'Saldo Programação'
df3['TRANSACAO'] = 'Saldo Programação'

result_alterando_saldo = pd.concat([df1, df2, df3])

valor_total = []
for linha in result_alterando_saldo.values:
    if linha[11] == 'Saldo Programação':
        valor_total.append(linha[3])
    else:
        valor_total.append(linha[8])

result_alterando_saldo['VALOR_TOTAL'] = valor_total

result_alterando_saldo = result_alterando_saldo[['TECE'
    , 'BANCO', 'TIPO_NUMERARIO', 'VLR_CEDULA_MOEDA'
    , 'VALOR_TOTAL', 'DESCRICAO', 'TRANSACAO', 'ATUALIZACAO'
    , 'OS', 'IDC_IBS', 'IDT_DTN_TEB', 'DTA_REG_OPR', 'DTA_PRG', 'PROTOCOLO'
    ]]

#PROCESSO NOVO DE INSERÇÃO
conn_alchemy_gnu = engine.connect()
with conn_alchemy_gnu as conn:
    with conn.begin() as beg:
        result_alterando_saldo.to_sql(name="TI_GESTAO_ALT_CUSTODIA"
                , con=conn
                , if_exists='append' #replace
                , index=False
                , chunksize=500
                , schema= 'MERCANTIL\B042786')
        beg.commit()


#%%
#INSERINDO TERMINAIS ALTERADOS NA TABELA [MERCANTIL\B042786].[TC_AUT_REMESSA_ATM_ALT_COMP]
if usando_prog == 1:
    with open (caminho_resumo, 'a') as resumo:
        resumo.write('\nInserindo os atms alterados na tabela responsável por gerar as PLANS\n...')
        resumo.close()
    print('\nInserindo os atms alterados na tabela responsável por gerar as PLANS...')
    cursor_gnu.execute("TRUNCATE TABLE [MERCANTIL\B042786].[TC_AUT_REMESSA_ATM_ALT_COMP]")
    cursor_gnu.commit()
    
    temp = info_atm[['NUM_DND', 'AGENCIA']].drop_duplicates()
    alterando_atm = alterando_atm.merge(temp, how='left', on='NUM_DND')
    today = date.today()
    lista_atms = ['0' * (4 - len(str(x))) + str(x) for x in alterando_atm['IDT_TML']]
    lista_ano = [str(x.year - 2000) for x in alterando_atm['DTA_HOR_PRG']]
    lista_mes = ['0' * (2 - len(str(x.month))) + str(x.month) for x in alterando_atm['DTA_HOR_PRG']]
    lista_dia = ['0' * (2 - len(str(x.day))) + str(x.day) for x in alterando_atm['DTA_HOR_PRG']]
    lista_hora = ['0' * (2 - len(str(x.hour))) + str(x.hour) for x in alterando_atm['DTA_HOR_PRG']]
    tipos_ced = [2, 5, 10, 20, 50, 100, 200]

    lista_os = []
    for k in range(len(alterando_atm['IDT_TML'].to_list())):
        v = lista_atms[k] + lista_ano[k] + lista_mes[k] + lista_dia[k] + lista_hora[k]
        lista_os.append(v)
    alterando_atm['OS'] = lista_os

    agencia = []
    for linha in alterando_atm.values:
        local = linha[17].strip()
        if linha[0].time() == datetime.time(19, 0):
            separador = local.find('-')
            local = local[:separador+2] + 'RESERVA TECNICA'
            agencia.append(local)
        else:
            agencia.append(local)
    alterando_atm['AGENCIA'] = agencia 

    result_alterando_atm = pd.DataFrame(columns=['IDT_TML'
    ,'VALOR_DENOMINACAO'
    ,'COD_BANCO'
    ,'COD_CEN'
    ,'OS'
    ,'DTA_PROG'
    ,'DTA_ENTRG'
    ,'TIP_ENTRG'
    ,'IDT_TML'
    ,'AGENCIA'
    ,'HOR_ENTRG'
    ,'COMP'
    ,'VLR_ABT_PRG'
    ,'DENOMINACAO'
    ])

    for linha in alterando_atm.values:
        comp_temp = remove_repetidos([int(x.strip()) for x in linha[6].split('-')])
        for x in comp_temp:
            
            tam_atual = len(result_alterando_atm)
            result_alterando_atm.loc[tam_atual] = [
                linha[2] #IDT_TML
                ,linha[tipos_ced.index(x) + 7] #VALOR_DENOMINACAO
                ,'389' #COD_BANCO
                ,linha[1] #COD_CEN
                ,linha[18] #OS
                ,today #DTA_PROG
                ,linha[0].date() #DTA_ENTRG
                ,'REMESSA ATM' #TIP_ENTRG
                ,linha[2] #IDT_TML
                ,linha[17] #AGENCIA
                ,linha[0].time() #HOR_ENTRG
                ,linha[6] #COMP
                ,linha[4] #VLR_ABT_PRG
                ,x #DENOMINACAO
            ]

    conn_alchemy_gnu = engine.connect()
    with conn_alchemy_gnu as conn:
        with conn.begin() as beg:
            result_alterando_atm.to_sql(name="TC_AUT_REMESSA_ATM_ALT_COMP"
                    , con=conn
                    , if_exists='append' #replace
                    , index=False
                    , chunksize=500
                    , schema= 'MERCANTIL\B042786')
            beg.commit()


#%%
#SALVANDO BKP EM [MERCANTIL\B042786].[TC_HST_ALT_PROG]
saldo_resumo = alterando_saldo[['TECE', 'TIPO_NUMERARIO', 'CEDULA', 'VALOR_ORIGINAL', 'VALOR_AJUSTE', 'VALOR_TOTAL']
    ].merge(metas[['TECE', 'CEDULA', 'META']], how='left', on=['TECE', 'CEDULA'])
saldo_resumo['META'] = [x if x >= 0 else 0 for x in saldo_resumo['META']]

with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nGerando BKP de valores alterados no SQLGDNP.GNU\n...')
    resumo.close()
print('\nGerando BKP de valores alterados no SQLGDNP.GNU ...')

cursor_gnu.execute('DELETE FROM [MERCANTIL\B042786].[TC_HST_ALT_PROG] WHERE DATA_REF >= CAST(GETDATE() AS DATE)')
saldo_resumo['DATA_REF'] = f'{ano}-{mes}-{dia}'
saldo_resumo = saldo_resumo[['DATA_REF'
                             ,'TECE'
                             , 'TIPO_NUMERARIO'
                             , 'CEDULA'
                             , 'META'
                             , 'VALOR_ORIGINAL'
                             , 'VALOR_AJUSTE'
                             , 'VALOR_TOTAL']]

#PROCESSO NOVO DE INSERÇÃO
conn_alchemy_gnu = engine.connect()
with conn_alchemy_gnu as conn:
    with conn.begin() as beg:
        saldo_resumo.to_sql(name="TC_HST_ALT_PROG"
                , con=conn
                , if_exists='append' #replace
                , index=False
                , chunksize=500
                , schema= 'MERCANTIL\B042786')
        beg.commit()


#%%
#INSERINDO VOLUME DE ALTERAÇÕES NA TABELA [MERCANTIL\B042786].TC_VOLUME_ALT_PROG
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nGerando BKP de quantidades de ATMs alterados no SQLGDNP.GNU\n...')
    resumo.close()
print('\nGerando BKP de quantidades de ATMs alterados no SQLGDNP.GNU ...')
bkp_alterando_atm = alterando_atm[['TECE', 'DTA_HOR_PRG', 'ALTERADO']]
bkp_alterando_atm['TOTAL'] = 1
bkp_alterando_atm = bkp_alterando_atm.groupby("TECE")['ALTERADO', 'TOTAL'].sum()

bkp_alterando_atm['TECE'] = bkp_alterando_atm.index
bkp_alterando_atm['DATA'] = f'{ano}-{mes}-{dia}'
bkp_alterando_atm = bkp_alterando_atm[['DATA', 'TECE', 'TOTAL', 'ALTERADO']]

cursor_gnu.execute('''DELETE FROM [MERCANTIL\B042786].TC_VOLUME_ALT_PROG 
WHERE DATA >= CAST(GETDATE() AS DATE)''')
cursor_gnu.commit() 

# PROCESSO DE INSERÇÃO ATUAL
conn_alchemy_gnu = engine.connect()
with conn_alchemy_gnu as conn:
    with conn.begin() as beg:
        bkp_alterando_atm.to_sql(name="TC_VOLUME_ALT_PROG"
                , con=conn
                , if_exists='append' #replace
                , index=False
                , chunksize=100
                , schema= 'MERCANTIL\B042786')
        beg.commit()


#%%
#SALVANDO RESULTADO EM EXCEL PARA CONFERÊNCIA
with open (caminho_resumo, 'a') as resumo:
    resumo.write('\nSalvando processo em excel para conferência\n...')
    resumo.close()
print(('\nSalvando processo em excel para conferência...'))
arquivo = pd.ExcelWriter(path=f'K:/GSAS/09 - Coordenacao Gestao Numerario/09 Prototipos SSIS/B042786/__PYTHON__/__EXECUTAVEIS__/PROGRAMACAO AUT/BKP ALTERAÇÕES/Alterações ({dia}-{mes}-{ano}).xlsx',engine='xlsxwriter')
saldo_resumo.to_excel(arquivo, sheet_name="Saldos_finais", index=False)
alterando_atm.to_excel(arquivo, sheet_name="Atms_alterados", index=False)
abastec_prog.to_excel(arquivo, sheet_name="Atms_iniciais", index=False)
#saldos_tratados.to_excel(arquivo, sheet_name="Saldos_Iniciais", index=False)
arquivo.save()
arquivo.close()

#%%
conn_gnu.close()
conn_mbcorp.close()
