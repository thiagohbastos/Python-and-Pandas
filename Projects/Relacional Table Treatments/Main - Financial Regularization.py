#%%
# ---------- IMPORTANDO PACOTES ----------
import pandas as pd
import numpy as np
from math import isnan
import pyodbc
import locale
from datetime import date
import warnings
warnings.simplefilter("ignore")


# %%
#---------- CONECTANDO AO BANCO DE DADOS ----------
conn_gnu = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server=SQLGDNP;"
    "Database=GNU;"
    "Trusted_Connection=yes;")

conn_corp1 = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server=SQLCORP1;"
    "Database=GNUB001D;"
    "Trusted_Connection=yes;")

data_corte_trat_2 = pd.Timestamp('2022-12-20 23:59:59')


#%%
#---------- LENDO QUERIES DA PASTA ----------
base_query = ''
with open('base_GNU_e_Controle.sql', 'r', encoding='utf-8') as x:
    for linha_atm in x:
        base_query += linha_atm
x.close()

erro_saq_query = ''
with open('erro_saque.sql', 'r', encoding='utf-8') as x:
    for linha_atm in x:
        erro_saq_query += linha_atm
x.close()

em_tratamento_query = ''
with open('registros_em_tratamento.sql', 'r', encoding='utf-8') as x:
    for linha_atm in x:
        em_tratamento_query += linha_atm
x.close()

atm_info_query = ''
with open('ATM_INFO.sql', 'r', encoding='utf-8') as x:
    for linha_atm in x:
        atm_info_query += linha_atm
x.close()

dilaceradas_query = ''
with open('dilaceradas.sql', 'r', encoding='utf-8') as x:
    for linha_atm in x:
        dilaceradas_query += linha_atm
x.close()

reprovadas_query = ''
with open('reprovadas.sql', 'r', encoding='utf-8') as x:
    for linha_atm in x:
        reprovadas_query += linha_atm
x.close()

#%%
#---------- EXTRAINDO BASES NO SERVIDOR ----------
em_tratamento = pd.read_sql_query(em_tratamento_query, conn_gnu)
atm_info = pd.read_sql_query(atm_info_query, conn_gnu)
base_completa = pd.read_sql_query(base_query, conn_gnu)
erro_saq_pla = pd.read_sql_query(erro_saq_query, conn_corp1)
dilaceradas_pla = pd.read_sql_query(dilaceradas_query, conn_corp1)
reprovadas_pla = pd.read_sql_query(reprovadas_query, conn_corp1)['NUM_DMD'].to_list()


#%%
#---------- TRATANDO BASES ----------
dilaceradas_pla['NUM_DMD'] = dilaceradas_pla.NUM_DMD.astype('int64')
dilaceradas_pla = dilaceradas_pla.merge(atm_info, how='left', on='COD_TML')

erro_saq_pla['NUM_DMD'] = erro_saq_pla.NUM_DMD.astype('int64')
erro_saq_pla = erro_saq_pla.merge(atm_info, how='left', on='COD_TML')
temp = []
for x in erro_saq_pla['IDT_CTA_IDZ'].to_list():
    if isnan(x):
        temp.append(0)
    else:
        temp.append(int(x))
erro_saq_pla['IDT_CTA_IDZ'] = temp
erro_saq = erro_saq_pla
dilaceradas = dilaceradas_pla

for desconsiderar in em_tratamento.values:
    for x in range(0, desconsiderar[1]):
        base_completa = base_completa.drop(base_completa.index[base_completa['IDT_LCT'] == desconsiderar[0]].tolist()[0])

credores = base_completa[base_completa['COD_TIP_LCT'] == 1]
devedores = base_completa[base_completa['COD_TIP_LCT'] == 2]
erro_dep = base_completa[base_completa['COD_TIP_LCT'] == 4]
dilaceradas_gnu = base_completa[base_completa['COD_TIP_LCT'] == 6]

erro_saq_gnu = base_completa[base_completa['COD_TIP_LCT'] == 3][[
    'NUM_DMD'
    ,'IDT_LCT'
    ,'IDT_CTA_IDZ'
    ,'COD_TML'
    ,'NUM_DND'
    ,'VLR_LCT'
    ,'VLR_DPN'
    ,'DTA_LCT'
    ,'TRAT_PY_1'
    ,'TRAT_PY_2'
    ]]
erro_saq_gnu['ID'] = 0
erro_saq_gnu['NUM_DMD'] = [int(x) for x in erro_saq_gnu['NUM_DMD'].to_list()]


#%%
#------- ANALISANDO DIVERGENCIA DE DEMANDAS NAS BASES GNU/PLA -------
dmds_gnu = erro_saq_gnu['NUM_DMD'].to_list()
dmds_pla = erro_saq_pla['NUM_DMD'].to_list()
es_coincidentes = []
es_nao_coincidentes = []
valor_coincidentes = valor_nao_coincidentes = 0
for x in dmds_pla:
    if x in dmds_gnu and x != 0:
        es_coincidentes.append(x)
        valor_coincidentes += erro_saq_pla[erro_saq_pla['NUM_DMD'] == x].values[0][10]
        erro_saq = erro_saq.drop(erro_saq.index[erro_saq['NUM_DMD'] == x].tolist()[0])
    elif x == 0:
        continue
    else:
        es_nao_coincidentes.append(x)
        valor_nao_coincidentes += erro_saq_pla[erro_saq_pla['NUM_DMD'] == x].values[0][10]
erro_saq = erro_saq.reset_index(drop=True)
erro_saq = pd.merge(erro_saq, erro_saq_gnu, how='outer')
erro_saq = erro_saq.reset_index(drop=True)
valor_nao_coincidentes += sum(erro_saq_pla[erro_saq_pla['NUM_DMD'] == 0]['VLR_DPN'])
qtd_coincidentes = len(es_coincidentes)
qtd_nao_coincidentes = len(es_nao_coincidentes) + len(erro_saq_pla[erro_saq_pla['NUM_DMD'] == 0])
print(f'Quantidade de demandas coincidentes: {qtd_coincidentes:.0f}')
print(f'Total do valor das demandas coincidentes: R${valor_coincidentes:.2f}')
print(f'Quantidade de demandas NÃO coincidentes: {qtd_nao_coincidentes:.0f}')
print(f'Total do valor das demandas NÃO coincidentes: R${valor_nao_coincidentes:.2f}')


#%%
#------- ANALISANDO DIVERGENCIA DILACERADAS GNU/PLA -------
dilaceradas_gnu['NUM_DMD'] = [int(x) for x in dilaceradas_gnu['NUM_DMD'].to_list()]
dmd_dilaceradas_gnu = dilaceradas_gnu['NUM_DMD'].to_list()
dmd_dilaceradas_pla = dilaceradas_pla['NUM_DMD'].to_list()
dilaceradas_coincidentes = []
dilaceradas_nao_coincidentes = []
valor_dilaceradas_coincidentes = valor_dilaceradas_nao_coincidentes = 0
for x in dmd_dilaceradas_pla:
    if x in dmd_dilaceradas_gnu and x != 0:
        dilaceradas_coincidentes.append(x)
        valor_dilaceradas_coincidentes += dilaceradas_pla[dilaceradas_pla['NUM_DMD'] == x].values[0][10]
        dilaceradas = dilaceradas.drop(dilaceradas.index[dilaceradas['NUM_DMD'] == x].tolist()[0])
    elif x == 0:
        continue
    else:
        dilaceradas_nao_coincidentes.append(x)
        valor_dilaceradas_nao_coincidentes += dilaceradas_pla[dilaceradas_pla['NUM_DMD'] == x].values[0][10]
dilaceradas = dilaceradas.reset_index(drop=True)
valor_dilaceradas_nao_coincidentes += sum(dilaceradas_pla[dilaceradas_pla['NUM_DMD'] == 0]['VLR_DPN'])
qtd_coincidentes = len(dilaceradas_coincidentes)
qtd_nao_coincidentes = len(dilaceradas_nao_coincidentes)
print(f'Quantidade de demandas coincidentes: {qtd_coincidentes:.0f}')
print(f'Total do valor das demandas coincidentes: R${valor_dilaceradas_coincidentes:.2f}')
print(f'Quantidade de demandas NÃO coincidentes: {qtd_nao_coincidentes:.0f}')
print(f'Total do valor das demandas NÃO coincidentes: R${valor_dilaceradas_nao_coincidentes:.2f}')

erro_saq['FK_IDT_LCT'] = np.nan
erro_dep['FK_IDT_LCT'] = np.nan
devedores['FK_IDT_LCT'] = np.nan
qtd_lanc = [0, 0, 0, 0, 0]


#%%
#---------- TRATAMENTO GERAL COM VALOR IGUAL ----------
cont = 0
for cred_lancamento in credores.values:
    cred_atm = cred_lancamento[list(credores.columns).index('COD_TML')]
    cred_tece = cred_lancamento[list(credores.columns).index('IDT_CTA_IDZ')]
    cred_valor_lct = cred_lancamento[list(credores.columns).index('VLR_LCT')]
    cred_saldo = cred_lancamento[list(credores.columns).index('VLR_DPN')]
    cred_data = cred_lancamento[list(credores.columns).index('DTA_LCT')]
    id_cred = cred_lancamento[list(credores.columns).index('IDT_LCT')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRAT_PY_1')]
    cred_trat_py_dev = cred_lancamento[list(credores.columns).index('TRAT_PY_DEV')]
    if isnan(cred_atm):
        cred_atm = 0

    if cred_saldo > 0 and cred_valor_lct == cred_saldo:
        df_es = erro_saq.query(f'COD_TML=={cred_atm}').query(f'VLR_DPN=={cred_saldo}').values
        df_ed = erro_dep.query(f'COD_TML=={cred_atm}').query(f'VLR_DPN=={cred_saldo}').values
        df_dev = devedores.query(f'COD_TML=={cred_atm}').query(f'VLR_DPN=={cred_saldo}').values

        qtd = 0
        for es_lancamento in df_es:
            if es_lancamento[list(erro_saq.columns).index('NUM_DMD')] in reprovadas_pla:
                continue

            #variáveis
            id_es = es_lancamento[list(erro_saq.columns).index('NUM_DMD')]
            es_tece = es_lancamento[list(erro_saq.columns).index('IDT_CTA_IDZ')]
            es_atm = es_lancamento[list(erro_saq.columns).index('COD_TML')]
            es_falta = es_lancamento[list(erro_saq.columns).index('VLR_DPN')]
            es_data = es_lancamento[list(erro_saq.columns).index('DTA_LCT')]
            es_fk_id_cred = es_lancamento[list(erro_saq.columns).index('FK_IDT_LCT')]
            data_dif = cred_data - es_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if cred_tece == es_tece and es_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo == es_falta:
                cred_trat_py += es_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_1'] = cred_trat_py

                cred_saldo -= es_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred

                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'TRAT_PY_1'] += es_falta

                es_falta -= es_falta
                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'VLR_DPN'] = es_falta

                cont += 1
                qtd += 1

        qtd_lanc[0] += qtd

        qtd = 0
        for ed_lancamento in df_ed:

            #variáveis
            id_ed = ed_lancamento[list(erro_dep.columns).index('NUM_DMD')]
            ed_tece = ed_lancamento[list(erro_dep.columns).index('IDT_CTA_IDZ')]
            ed_atm = ed_lancamento[list(erro_dep.columns).index('COD_TML')]
            ed_falta = ed_lancamento[list(erro_dep.columns).index('VLR_DPN')]
            ed_data = ed_lancamento[list(erro_dep.columns).index('DTA_LCT')]
            ed_fk_id_cred = ed_lancamento[list(erro_dep.columns).index('FK_IDT_LCT')]
            data_dif = cred_data - ed_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if cred_tece == ed_tece and ed_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo == ed_falta:
                cred_trat_py += ed_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_1'] = cred_trat_py

                cred_saldo -= ed_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred

                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'TRAT_PY_1'] += ed_falta

                ed_falta -= ed_falta
                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'VLR_DPN'] = ed_falta

                qtd += 1
                cont += 1

        qtd_lanc[2] += qtd

        qtd = 0
        for dev_lancamento in df_dev:

            #variáveis
            id_dev = dev_lancamento[list(devedores.columns).index('IDT_LCT')]
            dev_tece = dev_lancamento[list(devedores.columns).index('IDT_CTA_IDZ')]
            dev_atm = dev_lancamento[list(devedores.columns).index('COD_TML')]
            dev_falta = dev_lancamento[list(devedores.columns).index('VLR_DPN')]
            dev_data = dev_lancamento[list(devedores.columns).index('DTA_LCT')]
            dev_fk_id_cred = dev_lancamento[list(devedores.columns).index('FK_IDT_LCT')]
            data_dif = cred_data - dev_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if cred_tece == dev_tece and dev_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo == dev_falta:
                cred_trat_py_dev += dev_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_DEV'] = cred_trat_py_dev

                cred_saldo -= dev_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(dev_fk_id_cred)
                except:
                    status = False
                if status:
                    dev_fk_id_cred = id_cred
                    devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                    , 'FK_IDT_LCT'] = dev_fk_id_cred
                else:
                    try:
                        dev_fk_id_cred = int(dev_fk_id_cred)
                    except:
                        pass
                    dev_fk_id_cred = f'{dev_fk_id_cred}, {id_cred}'
                    devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                    , 'FK_IDT_LCT'] = dev_fk_id_cred

                devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                , 'TRAT_PY_DEV'] += dev_falta

                dev_falta -= dev_falta
                devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                , 'VLR_DPN'] = dev_falta

                qtd += 1
                cont += 1

        qtd_lanc[4] += qtd


#%%
#---------- TRATAMENTO PRIMÁRIO ERRO DE SAQUE ----------
cont = qtd = 0
for cred_lancamento in credores.values:
    cred_atm = cred_lancamento[list(credores.columns).index('COD_TML')]
    cred_tece = cred_lancamento[list(credores.columns).index('IDT_CTA_IDZ')]
    cred_saldo = cred_lancamento[list(credores.columns).index('VLR_DPN')]
    cred_data = cred_lancamento[list(credores.columns).index('DTA_LCT')]
    id_cred = cred_lancamento[list(credores.columns).index('IDT_LCT')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRAT_PY_1')]
    if isnan(cred_atm):
        cred_atm = 0

    if cred_saldo > 0:
        df = erro_saq.query(f'COD_TML=={cred_atm}').values

        for es_lancamento in df:
            if es_lancamento[list(erro_saq.columns).index('NUM_DMD')] in reprovadas_pla:
                continue

            #variáveis
            id_es = es_lancamento[list(erro_saq.columns).index('NUM_DMD')]
            es_tece = es_lancamento[list(erro_saq.columns).index('IDT_CTA_IDZ')]
            es_atm = es_lancamento[list(erro_saq.columns).index('COD_TML')]
            es_falta = es_lancamento[list(erro_saq.columns).index('VLR_DPN')]
            es_data = es_lancamento[list(erro_saq.columns).index('DTA_LCT')]
            es_fk_id_cred = es_lancamento[list(erro_saq.columns).index('FK_IDT_LCT')]
            data_dif = cred_data - es_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if cred_tece == es_tece and es_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= es_falta:
                cred_trat_py += es_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_1'] = cred_trat_py

                cred_saldo -= es_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred

                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'TRAT_PY_1'] += es_falta

                es_falta -= es_falta
                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'VLR_DPN'] = es_falta

                cont += 1
                qtd += 1

            elif cred_tece == es_tece and es_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < es_falta:
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_1'] = cred_trat_py

                es_falta -= cred_saldo
                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'VLR_DPN'] = es_falta

                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'TRAT_PY_1'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred

                cont += 1
                qtd += 1

qtd_lanc[0] += qtd


#%%
#---------- TRATAMENTO PRIMÁRIO ERRO DE DEPÓSITO ----------
cont = qtd = 0
for cred_lancamento in credores.values:
    cred_atm = cred_lancamento[list(credores.columns).index('COD_TML')]
    cred_tece = cred_lancamento[list(credores.columns).index('IDT_CTA_IDZ')]
    cred_saldo = cred_lancamento[list(credores.columns).index('VLR_DPN')]
    cred_data = cred_lancamento[list(credores.columns).index('DTA_LCT')]
    id_cred = cred_lancamento[list(credores.columns).index('IDT_LCT')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRAT_PY_1')]
    if isnan(cred_atm):
        cred_atm = 0

    if cred_saldo > 0:
        df = erro_dep.query(f'COD_TML=={cred_atm}').values

        for ed_lancamento in df:

            #variáveis
            id_ed = ed_lancamento[list(erro_dep.columns).index('NUM_DMD')]
            ed_tece = ed_lancamento[list(erro_dep.columns).index('IDT_CTA_IDZ')]
            ed_atm = ed_lancamento[list(erro_dep.columns).index('COD_TML')]
            ed_falta = ed_lancamento[list(erro_dep.columns).index('VLR_DPN')]
            ed_data = ed_lancamento[list(erro_dep.columns).index('DTA_LCT')]
            ed_fk_id_cred = ed_lancamento[list(erro_dep.columns).index('FK_IDT_LCT')]
            data_dif = cred_data - ed_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if cred_tece == ed_tece and ed_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= ed_falta:
                cred_trat_py += ed_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_1'] = cred_trat_py

                cred_saldo -= ed_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred

                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'TRAT_PY_1'] += ed_falta

                ed_falta -= ed_falta
                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'VLR_DPN'] = ed_falta

                cont += 1
                qtd += 1

            elif cred_tece == ed_tece and ed_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < ed_falta:
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_1'] = cred_trat_py

                ed_falta -= cred_saldo
                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'VLR_DPN'] = ed_falta

                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'TRAT_PY_1'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred

                cont += 1
                qtd += 1

qtd_lanc[1] += qtd


#%%
#---------- TRATAMENTO DEVEDORES ----------
cont = qtd = 0
for cred_lancamento in credores.values:
    cred_atm = cred_lancamento[list(credores.columns).index('COD_TML')]
    cred_tece = cred_lancamento[list(credores.columns).index('IDT_CTA_IDZ')]
    cred_saldo = cred_lancamento[list(credores.columns).index('VLR_DPN')]
    cred_data = cred_lancamento[list(credores.columns).index('DTA_LCT')]
    id_cred = cred_lancamento[list(credores.columns).index('IDT_LCT')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRAT_PY_DEV')]
    if isnan(cred_atm):
        cred_atm = 0

    if cred_saldo > 0:
        df = devedores.query(f'COD_TML=={cred_atm}').values
        
        for dev_lancamento in df:

            #variáveis
            id_dev = dev_lancamento[list(devedores.columns).index('IDT_LCT')]
            dev_tece = dev_lancamento[list(devedores.columns).index('IDT_CTA_IDZ')]
            dev_atm = dev_lancamento[list(devedores.columns).index('COD_TML')]
            dev_falta = dev_lancamento[list(devedores.columns).index('VLR_DPN')]
            dev_data = dev_lancamento[list(devedores.columns).index('DTA_LCT')]
            dev_fk_id_cred = dev_lancamento[list(devedores.columns).index('FK_IDT_LCT')]
            data_dif = cred_data - dev_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if cred_tece == dev_tece and dev_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= dev_falta:
                cred_trat_py += dev_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_DEV'] = cred_trat_py

                cred_saldo -= dev_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(dev_fk_id_cred)
                except:
                    status = False
                if status:
                    dev_fk_id_cred = id_cred
                    devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                    , 'FK_IDT_LCT'] = dev_fk_id_cred
                else:
                    try:
                        dev_fk_id_cred = int(dev_fk_id_cred)
                    except:
                        pass
                    dev_fk_id_cred = f'{dev_fk_id_cred}, {id_cred}'
                    devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                    , 'FK_IDT_LCT'] = dev_fk_id_cred

                devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                , 'TRAT_PY_DEV'] += dev_falta

                dev_falta -= dev_falta
                devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                , 'VLR_DPN'] = dev_falta

                cont += 1
                qtd += 1

            elif cred_tece == dev_tece and dev_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < dev_falta:
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_DEV'] = cred_trat_py

                dev_falta -= cred_saldo
                devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                , 'VLR_DPN'] = dev_falta

                devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                , 'TRAT_PY_DEV'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(dev_fk_id_cred)
                except:
                    status = False
                if status:
                    dev_fk_id_cred = id_cred
                    devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                    , 'FK_IDT_LCT'] = dev_fk_id_cred
                else:
                    try:
                        dev_fk_id_cred = int(dev_fk_id_cred)
                    except:
                        pass
                    dev_fk_id_cred = f'{dev_fk_id_cred}, {id_cred}'
                    devedores.at[devedores.index[devedores['IDT_LCT'] == id_dev].tolist()[0]
                    , 'FK_IDT_LCT'] = dev_fk_id_cred

                cont += 1
                qtd += 1

qtd_lanc[4] += qtd


#%%
#---------- TRATAMENTO SECUNDÁRIO ERRO DE SAQUE ----------
cont = qtd = 0
for cred_lancamento in credores.values:
    #variáveis
    cred_agencia = cred_lancamento[list(credores.columns).index('NUM_DND')]
    cred_tece = cred_lancamento[list(credores.columns).index('IDT_CTA_IDZ')]
    cred_saldo = cred_lancamento[list(credores.columns).index('VLR_DPN')]
    cred_data = cred_lancamento[list(credores.columns).index('DTA_LCT')]
    id_cred = cred_lancamento[list(credores.columns).index('IDT_LCT')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRAT_PY_2')]
    if isnan(cred_agencia):
        cred_agencia = 0

    if cred_saldo > 0 and cred_data <= data_corte_trat_2 and cred_agencia != 0 and not isnan(cred_agencia):
        df = erro_saq.query(f'NUM_DND=={cred_agencia}').values

        for es_lancamento in df:
            if es_lancamento[list(erro_saq.columns).index('NUM_DMD')] in reprovadas_pla:
                continue

            #variáveis
            es_tece = es_lancamento[list(erro_saq.columns).index('IDT_CTA_IDZ')]
            es_agencia = es_lancamento[list(erro_saq.columns).index('NUM_DND')]
            es_data = es_lancamento[list(erro_saq.columns).index('DTA_LCT')]
            es_validacao = es_data <= data_corte_trat_2
            id_es = es_lancamento[list(erro_saq.columns).index('NUM_DMD')]
            es_falta = es_lancamento[list(erro_saq.columns).index('VLR_DPN')]
            es_fk_id_cred = es_lancamento[list(erro_saq.columns).index('FK_IDT_LCT')]
            data_dif = cred_data - es_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if cred_tece == es_tece and es_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= es_falta and es_validacao:
                
                cred_trat_py += es_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_2'] = cred_trat_py

                cred_saldo -= es_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred

                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'TRAT_PY_2'] += es_falta

                es_falta -= es_falta
                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'VLR_DPN'] = es_falta

                cont += 1
                qtd += 1

            elif cred_tece == es_tece and es_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < es_falta and es_validacao:             
                
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_2'] = cred_trat_py

                es_falta -= cred_saldo
                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'VLR_DPN'] = es_falta

                erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                , 'TRAT_PY_2'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saq.at[erro_saq.index[erro_saq['NUM_DMD'] == id_es].tolist()[0]
                    , 'FK_IDT_LCT'] = es_fk_id_cred

                cont += 1
                qtd += 1

qtd_lanc[2] += qtd


#%%
#---------- TRATAMENTO SECUNDÁRIO ERRO DE DEPÓSITO ----------
cont = qtd = 0
for cred_lancamento in credores.values:
    #variáveis
    cred_agencia = cred_lancamento[list(credores.columns).index('NUM_DND')]
    cred_tece = cred_lancamento[list(credores.columns).index('IDT_CTA_IDZ')]
    cred_saldo = cred_lancamento[list(credores.columns).index('VLR_DPN')]
    cred_data = cred_lancamento[list(credores.columns).index('DTA_LCT')]
    id_cred = cred_lancamento[list(credores.columns).index('IDT_LCT')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRAT_PY_2')]
    if isnan(cred_agencia):
        cred_agencia = 0
    

    if cred_saldo > 0 and cred_data <= data_corte_trat_2 and cred_agencia != 0 and not isnan(cred_agencia):
        
        df = erro_dep.query(f'NUM_DND=={cred_agencia}').values

        for ed_lancamento in df:

            #variáveis
            ed_agencia = ed_lancamento[list(erro_dep.columns).index('NUM_DND')]
            ed_tece = ed_lancamento[list(erro_dep.columns).index('IDT_CTA_IDZ')]
            ed_data = ed_lancamento[list(erro_dep.columns).index('DTA_LCT')]
            ed_validacao = ed_data <= data_corte_trat_2
            id_ed = ed_lancamento[list(erro_dep.columns).index('NUM_DMD')]
            ed_falta = ed_lancamento[list(erro_dep.columns).index('VLR_DPN')]
            ed_fk_id_cred = ed_lancamento[list(erro_dep.columns).index('FK_IDT_LCT')]
            data_dif = cred_data - ed_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if cred_tece == ed_tece and ed_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= ed_falta and ed_validacao:
                
                cred_trat_py += ed_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_2'] = cred_trat_py

                cred_saldo -= ed_falta
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred

                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'TRAT_PY_2'] += ed_falta

                ed_falta -= ed_falta
                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'VLR_DPN'] = ed_falta

                cont += 1
                qtd += 1

            elif cred_tece == dev_tece and ed_falta > 0 and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < ed_falta and ed_validacao:             
                
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'TRAT_PY_2'] = cred_trat_py

                ed_falta -= cred_saldo
                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'VLR_DPN'] = ed_falta

                erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                , 'TRAT_PY_2'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['IDT_LCT'] == id_cred].tolist()[0]
                , 'VLR_DPN'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_dep.at[erro_dep.index[erro_dep['NUM_DMD'] == id_ed].tolist()[0]
                    , 'FK_IDT_LCT'] = ed_fk_id_cred

                cont += 1
                qtd += 1

qtd_lanc[3] += qtd


#%%
#---------- CRIANDO RESUMO DAS REGULARIZAÇÕES ----------
locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')
resumo_df = pd.DataFrame(
            {'Valor': [locale.currency(erro_saq['TRAT_PY_1'].sum(), grouping=True)
                ,locale.currency(erro_saq['TRAT_PY_2'].sum(), grouping=True)
                ,locale.currency(erro_dep['TRAT_PY_1'].sum(), grouping=True)
                ,locale.currency(erro_dep['TRAT_PY_2'].sum(), grouping=True)
                ,locale.currency(devedores['TRAT_PY_DEV'].sum(), grouping=True)
                ,locale.currency(credores['TRAT_PY_1'].sum() + credores['TRAT_PY_2'].sum() + credores['TRAT_PY_DEV'].sum(), grouping=True)]
            ,'Qtd Registros': [qtd_lanc[0]
                ,qtd_lanc[2]
                ,qtd_lanc[1]
                ,qtd_lanc[3]
                ,qtd_lanc[4]
                ,sum(qtd_lanc)]
            }, index=['Trat. Primário ES', 'Trat. Secundário ES', 'Trat. Primário ED', 'Trat. Secundário ED', 'Trat. Devedores', 'TOTAL'])


#%%
#---------- ADAPTANDO FORMATO ERRO DE SAQUE PARA INSERÇÃO ----------
lista_padrao = [
    'IDT_LCT',
    'DTA_LCT',    
    'NUM_DND',    
    'COD_TML',    
    'COD_TIP_LCT',
    'VLR_LCT',    
    'IDT_CTA_IDZ',
    'COD_TSR',    
    'COD_STT_LCT',
    'VLR_DPN',    
    'NUM_DOC',    
    'NUM_DMD' 
    ]
erro_saq_pla = erro_saq[erro_saq['ID'] != 0]
erro_saq_pla['COD_TIP_LCT'] = 3
erro_saq_pla['COD_TSR'] = erro_saq_pla['IDT_CTA_IDZ']
erro_saq_pla['COD_STT_LCT'] = [1 if x > 0 else 2 for x in erro_saq_pla['VLR_DPN']]
erro_saq_pla['NUM_DOC'] = 0
erro_saq_pla['IDT_LCT'] = 0
erro_saq_pla = erro_saq_pla[lista_padrao]

dilaceradas['IDT_LCT'] = 0
dilaceradas['COD_TIP_LCT'] = 6
dilaceradas['COD_TSR'] = dilaceradas['IDT_CTA_IDZ']
dilaceradas['COD_STT_LCT'] = 5
dilaceradas['NUM_DOC'] = 0
dilaceradas = dilaceradas[lista_padrao]


#%%
#---------- ADAPTANDO DEMAIS CASOS PARA UPDATE ----------
erro_saq_gnu = erro_saq[erro_saq['ID'] == 0]
erro_saq_gnu = erro_saq_gnu.query('TRAT_PY_1!=0|TRAT_PY_2!=0')
erro_saq_gnu['COD_TIP_LCT'] = 3
erro_saq_gnu['COD_TSR'] = erro_saq_gnu['IDT_CTA_IDZ']
erro_saq_gnu['COD_STT_LCT'] = [1 if x > 0 else 2 for x in erro_saq_gnu['VLR_DPN']]
erro_saq_gnu['NUM_DOC'] = 0
erro_saq_gnu = erro_saq_gnu[lista_padrao]

credores_trat = credores.query('TRAT_PY_1!=0|TRAT_PY_2!=0|TRAT_PY_DEV!=0')
credores_trat['COD_STT_LCT'] = [1 if x > 0 else 2 for x in credores_trat['VLR_DPN']]
credores_trat = credores_trat[lista_padrao]

devedores_trat = devedores.query('TRAT_PY_1!=0|TRAT_PY_2!=0|TRAT_PY_DEV!=0')
devedores_trat['COD_STT_LCT'] = [1 if x > 0 else 2 for x in devedores_trat['VLR_DPN']]
devedores_trat = devedores_trat[lista_padrao]

erro_dep_trat = erro_dep.query('TRAT_PY_1!=0|TRAT_PY_2!=0|TRAT_PY_DEV!=0')
erro_dep_trat['COD_STT_LCT'] = [1 if x > 0 else 2 for x in erro_dep_trat['VLR_DPN']]
erro_dep_trat = erro_dep_trat[lista_padrao]

insert = pd.concat([erro_saq_pla, dilaceradas])
update = pd.concat([credores_trat, devedores_trat, erro_saq_gnu, erro_dep_trat  ])


#%%
#---------- SALVANDO ARQUIVOS RESULTADO ----------
arquivo = pd.ExcelWriter(path=f'K:/GSAS/09 - Coordenacao Gestao Numerario/09 Prototipos SSIS/B042786/__PYTHON__/CONCILIACOES GNU E CONTROLE/REPOSIÇÃO SISTEMA GNU/ARQUIVOS_PROCESSO/VALIDAÇÕES/Detalhes para Homologação ({date.today().day}-{date.today().month}-{date.today().year}).xlsx',engine='xlsxwriter')
resumo_df.to_excel(arquivo, sheet_name="RESUMO")
credores.to_excel(arquivo, sheet_name="Credores_Final", index=False)
erro_saq.to_excel(arquivo, sheet_name="Erro_Saque_Final", index=False)
erro_saq_gnu.to_excel(arquivo, sheet_name="ES_GNU", index=False)
erro_saq_pla.to_excel(arquivo, sheet_name="ES_PLA", index=False)
erro_dep.to_excel(arquivo, sheet_name="Erro_Deposito_Final", index=False)
devedores.to_excel(arquivo, sheet_name="Devedores_Final", index=False)
dilaceradas.to_excel(arquivo, sheet_name="Dilaceradas", index=False)
dilaceradas_gnu.to_excel(arquivo, sheet_name="Dilaceradas_GNU", index=False)
dilaceradas_pla.to_excel(arquivo, sheet_name="Dilaceradas_PLA", index=False)
arquivo.save()
arquivo.close()

arquivo = pd.ExcelWriter(path=f'K:/GSAS/09 - Coordenacao Gestao Numerario/09 Prototipos SSIS/B042786/__PYTHON__/CONCILIACOES GNU E CONTROLE/REPOSIÇÃO SISTEMA GNU/ARQUIVOS_PROCESSO/ARQUIVOS_FINAIS/Arquivo Final ({date.today().day}-{date.today().month}-{date.today().year}).xlsx',engine='xlsxwriter')
insert.to_excel(arquivo, sheet_name="INSERT", index=False)
update.to_excel(arquivo, sheet_name="UPDATE", index=False)
arquivo.save()
arquivo.close()


#%%
#---------- INSERINDO RESULTADO CONSOLIDADO EM BANCO ----------
'''cursor = conn_gnu.cursor()
insert['TIPO'] = 'INSERT'
update['TIPO'] = 'UPDATE'
consolidado_bd = pd.concat([insert, update])

cursor.execute(f"""DELETE FROM [MERCANTIL\B042786].[TC_BKP_CARGA_GNU] 
    WHERE DTA_REF = '{date.today().year}-{date.today().month}-{date.today().day}'""")
for  index, row in consolidado_bd.iterrows():
    cursor.execute("""INSERT INTO [MERCANTIL\B042786].[TC_BKP_CARGA_GNU] VALUES(
        CAST(GETDATE() AS DATE), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
    )"""
    ,row.DTA_LCT, row.NUM_DND, row.COD_TML, row.COD_TIP_LCT	
    ,row.VLR_LCT, row.IDT_CTA_IDZ, row.COD_TSR, row.COD_STT_LCT, row.VLR_DPN	
    ,row.NUM_DOC, row.NUM_DMD, row.TIPO)

cursor.commit()'''

#%%
#---------- FECHANDO CONEXÕES E FINALIZANDO PROCESSO ----------
conn_gnu.close()
conn_corp1.close()
print('{:-^57}'.format(' PROCESSO FINALIZADO '), end='\n\n')
