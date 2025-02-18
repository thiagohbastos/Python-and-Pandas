#%%
# ---------- IMPORTANDO PACOTES ----------
import pandas as pd
from math import isnan
import locale
import warnings
warnings.simplefilter("ignore")

# %%
#---------- EXTRAINDO DADOS EM SEUS CAMINHOS ----------
credores = pd.read_excel('//fsclt01grps.mercantil.com.br/grupos/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/DEVEDORES E CREDORES/GESTÃO CREDORES - 7914-5.xlsm', skiprows=3, sheet_name='Credores')
erro_saque = pd.read_excel('//fsclt01grps.mercantil.com.br/grupos/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/ERRO DE SAQUE E DEPOSITO/NOVO_ERRO DE SAQUE.xlsm', skiprows=4, sheet_name='ERROS DE SAQUE')
erro_deposito = pd.read_excel('//fsclt01grps.mercantil.com.br/grupos/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/ERRO DE SAQUE E DEPOSITO/NOVO_ERRO DE DEPÓSITO.xlsm', skiprows=4, sheet_name='ERROS DE DEPÓSITO')
devedores = pd.read_excel('//fsclt01grps.mercantil.com.br/grupos/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/DEVEDORES E CREDORES/GESTÃO DEVEDORES - 6864-3.xlsm', skiprows=4, sheet_name='Devedores')

# %%
#---------- INSERÇÃO DA COLUNA DE TRATAMENTO PRIMÁRIO ----------
credores['TRATAMENTO_1_PY'] = [0 for x in range(len(credores))]
credores['DEVEDORES_TRAT_PY'] = [0 for x in range(len(credores))]
erro_saque['TRATAMENTO_1_PY'] = [0 for x in range(len(erro_saque))]
erro_deposito['TRATAMENTO_1_PY'] = [0 for x in range(len(erro_deposito))]
devedores['TRATAMENTO_PY'] = [0 for x in range(len(devedores))]

#%%
#---------- INSERÇÃO DA COLUNA DE TRATAMENTO SECUNDÁRIO ----------
credores['TRATAMENTO_2_PY'] = [0 for x in range(len(credores))]
erro_saque['TRATAMENTO_2_PY'] = [0 for x in range(len(erro_saque))]
erro_deposito['TRATAMENTO_2_PY'] = [0 for x in range(len(erro_deposito))]
qtd_lanc = list()


#%%
#---------- DEFINIÇÃO DE DATA LIMITE ----------
data_maxima = "'2022-07-11'"
credores = credores.rename(columns={'Data Diferença': 'Data_Diferença'}).query(f'Data_Diferença<={data_maxima}').rename(columns={'Data_Diferença': 'Data Diferença'})
devedores = devedores.rename(columns={'Data Diferença': 'Data_Diferença'}).query(f'Data_Diferença<={data_maxima}').rename(columns={'Data_Diferença': 'Data Diferença'})
erro_saque = erro_saque.rename(columns={'INCLUSAO NA PLANILHA': 'DATA_DO_SAQUE'}).query(f'DATA_DO_SAQUE<={data_maxima}').rename(columns={'DATA_DO_SAQUE': 'INCLUSAO NA PLANILHA'})
erro_deposito = erro_deposito.rename(columns={'DATA DO DEPÓSITO': 'DATA_DO_DEPÓSITO'}).query(f'DATA_DO_DEPÓSITO<={data_maxima}').rename(columns={'DATA_DO_DEPÓSITO': 'DATA DO DEPÓSITO'})


#%%
#---------- TRATAMENTO SOBRAS ESPECÍFICAS DEVEDORES ----------
print("---------- TRATAMENTO SOBRAS ESPECÍFICAS DEVEDORES ----------")
cont = 0
dev_min_data = devedores['Data Diferença'].min()
for cred_lancamento in credores.values:
    #if cont >= 10:
    #    break

    #variáveis
    cred_atm = cred_lancamento[list(credores.columns).index('ATM')]
    cred_tece = cred_lancamento[list(credores.columns).index('Tesouraria')]
    cred_saldo = cred_lancamento[list(credores.columns).index('Diferença')]
    cred_data = cred_lancamento[list(credores.columns).index('Data Diferença')]
    id_cred = cred_lancamento[list(credores.columns).index('ID_CREDORES')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('DEVEDORES_TRAT_PY')]
    if isnan(cred_atm):
        cred_atm = 0
    df = devedores.query(f'ATM=={cred_atm}').values

    if cred_saldo >= 1000 and cred_data >= dev_min_data:
        
        for dev_lancamento in df:

            #variáveis
            id_dev = dev_lancamento[list(devedores.columns).index('ID_DEVEDORES')]
            dev_atm = dev_lancamento[list(devedores.columns).index('ATM')]
            dev_tece = dev_lancamento[list(devedores.columns).index('Tesouraria')]
            dev_falta = dev_lancamento[list(devedores.columns).index('Diferença')]
            dev_data = dev_lancamento[list(devedores.columns).index('Data Diferença')]
            dev_fk_id_cred = dev_lancamento[list(devedores.columns).index('FK_ID_CREDORES')]
            data_dif = cred_data - dev_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if dev_falta == cred_saldo and 0 <= data_dif <= 60 and cred_tece == dev_tece:
                
                cred_trat_py += dev_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'DEVEDORES_TRAT_PY'] = cred_trat_py

                cred_saldo -= dev_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(dev_fk_id_cred)
                except:
                    status = False
                if status:
                    dev_fk_id_cred = id_cred
                    devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                    , 'FK_ID_CREDORES'] = dev_fk_id_cred
                else:
                    try:
                        dev_fk_id_cred = int(dev_fk_id_cred)
                    except:
                        pass
                    dev_fk_id_cred = f'{dev_fk_id_cred}, {id_cred}'
                    devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                    , 'FK_ID_CREDORES'] = dev_fk_id_cred

                devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                , 'TRATAMENTO_PY'] += dev_falta

                dev_falta -= dev_falta
                devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                , 'Diferença'] = dev_falta

                cont += 1
                print(cont)

            #if cont >= 10:
            #    break
qtd_lanc.append(cont)
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')


# %%
#---------- TRATAMENTO PRIMÁRIO ERRO DE SAQUE ----------
print("---------- TRATAMENTO PRIMÁRIO ERRO DE SAQUE ----------")
cont = 0
for cred_lancamento in credores.values:
    #if cont >= 10:
    #    break

    #variáveis
    cred_atm = cred_lancamento[list(credores.columns).index('ATM')]
    cred_tece = cred_lancamento[list(credores.columns).index('Tesouraria')]
    cred_saldo = cred_lancamento[list(credores.columns).index('Diferença')]
    cred_data = cred_lancamento[list(credores.columns).index('Data Diferença')]
    id_cred = cred_lancamento[list(credores.columns).index('ID_CREDORES')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRATAMENTO_1_PY')]
    if isnan(cred_atm):
        cred_atm = 0
    df = erro_saque.query(f'ATM=={cred_atm}').values
    if cred_saldo > 0:
        
        for es_lancamento in df:

            #variáveis
            es_faixa = es_lancamento[list(erro_saque.columns).index('Controle')]
            es_motivo = es_lancamento[list(erro_saque.columns).index('MOTIVO')]
            es_validacao = es_faixa == 549 and es_motivo not in('02 - CEDULA MUTILADA',
            '04 - CEDULA SUSPEITA',
            '03 - CEDULA DILACERADA',
            '05 - CEDULA ENTINTADA'
            )
            id_es = es_lancamento[list(erro_saque.columns).index('ID_ES')]
            es_atm = es_lancamento[list(erro_saque.columns).index('ATM')]
            es_tece = es_lancamento[list(erro_saque.columns).index('TECE')]
            es_falta = es_lancamento[list(erro_saque.columns).index('DEVEDORES DEP.5100-5           FX.549')]
            es_data = es_lancamento[list(erro_saque.columns).index('INCLUSAO NA PLANILHA')]
            es_fk_id_cred = es_lancamento[list(erro_saque.columns).index('FK_ID_CREDORES')]
            data_dif = cred_data - es_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if es_falta > 0 and cred_tece == es_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= es_falta and es_validacao:
                
                cred_trat_py += es_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'TRATAMENTO_1_PY'] = cred_trat_py

                cred_saldo -= es_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                    , 'FK_ID_CREDORES'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                    , 'FK_ID_CREDORES'] = es_fk_id_cred

                erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                , 'TRATAMENTO_1_PY'] += es_falta

                es_falta -= es_falta
                erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                , 'DEVEDORES DEP.5100-5           FX.549'] = es_falta

                cont += 1
                print(cont)

            elif es_falta > 0 and cred_tece == es_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < es_falta and es_validacao:             
                
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'TRATAMENTO_1_PY'] = cred_trat_py

                es_falta -= cred_saldo
                erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                , 'DEVEDORES DEP.5100-5           FX.549'] = es_falta

                erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                , 'TRATAMENTO_1_PY'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                    , 'FK_ID_CREDORES'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                    , 'FK_ID_CREDORES'] = es_fk_id_cred

                cont += 1
                print(cont)

            #if cont >= 10:
            #    break
qtd_lanc.append(cont)
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')


# %%
#---------- TRATAMENTO PRIMÁRIO ERRO DE DEPÓSITO ----------
print("---------- TRATAMENTO PRIMÁRIO ERRO DE DEPÓSITO ----------")
cont = 0
for cred_lancamento in credores.values:
    #if cont >= 10:
    #    break

    #variáveis
    cred_atm = cred_lancamento[list(credores.columns).index('ATM')]
    cred_tece = cred_lancamento[list(credores.columns).index('Tesouraria')]
    cred_saldo = cred_lancamento[list(credores.columns).index('Diferença')]
    cred_data = cred_lancamento[list(credores.columns).index('Data Diferença')]
    id_cred = cred_lancamento[list(credores.columns).index('ID_CREDORES')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRATAMENTO_1_PY')]
    if isnan(cred_atm):
        cred_atm = 0
    df = erro_deposito.query(f'ATM=={cred_atm}').values

    if cred_saldo > 0:
        
        for ed_lancamento in df:

            #variáveis
            id_ed = ed_lancamento[list(erro_deposito.columns).index('ID_ED')]
            ed_atm = ed_lancamento[list(erro_deposito.columns).index('ATM')]
            ed_tece = ed_lancamento[list(erro_deposito.columns).index('TECE')]
            ed_falta = ed_lancamento[list(erro_deposito.columns).index('DEVEDORES DEP.5100-5           FX.6864-3    CI.716-4')]
            ed_data = ed_lancamento[list(erro_deposito.columns).index('DATA DO DEPÓSITO')]
            ed_fk_id_cred = ed_lancamento[list(erro_deposito.columns).index('FK_ID_CREDORES')]
            data_dif = cred_data - ed_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if ed_falta > 0 and cred_tece == ed_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= ed_falta:
                
                cred_trat_py += ed_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'TRATAMENTO_1_PY'] = cred_trat_py

                cred_saldo -= ed_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                    , 'FK_ID_CREDORES'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                    , 'FK_ID_CREDORES'] = ed_fk_id_cred

                erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                , 'TRATAMENTO_1_PY'] += ed_falta

                ed_falta -= ed_falta
                erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                , 'DEVEDORES DEP.5100-5           FX.6864-3    CI.716-4'] = ed_falta

                cont += 1
                print(cont)

            elif ed_falta > 0 and cred_tece == ed_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < ed_falta:             
                
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'TRATAMENTO_1_PY'] = cred_trat_py

                ed_falta -= cred_saldo
                erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                , 'DEVEDORES DEP.5100-5           FX.6864-3    CI.716-4'] = ed_falta

                erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                , 'TRATAMENTO_1_PY'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                    , 'FK_ID_CREDORES'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                    , 'FK_ID_CREDORES'] = ed_fk_id_cred

                cont += 1
                print(cont)

            #if cont >= 10:
            #    break
qtd_lanc.append(cont)
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')


#%%
#---------- TRATAMENTO DEVEDORES ----------
print("---------- TRATAMENTO DEVEDORES ----------")
cont = 0
for cred_lancamento in credores.values:
    #if cont >= 10:
    #    break

    #variáveis
    cred_atm = cred_lancamento[list(credores.columns).index('ATM')]
    cred_tece = cred_lancamento[list(credores.columns).index('Tesouraria')]
    cred_saldo = cred_lancamento[list(credores.columns).index('Diferença')]
    cred_data = cred_lancamento[list(credores.columns).index('Data Diferença')]
    id_cred = cred_lancamento[list(credores.columns).index('ID_CREDORES')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('DEVEDORES_TRAT_PY')]
    if isnan(cred_atm):
        cred_atm = 0
    df = devedores.query(f'ATM=={cred_atm}').values

    if cred_saldo > 0 and cred_data >= dev_min_data:
        
        for dev_lancamento in df:

            #variáveis
            id_dev = dev_lancamento[list(devedores.columns).index('ID_DEVEDORES')]
            dev_atm = dev_lancamento[list(devedores.columns).index('ATM')]
            dev_tece = dev_lancamento[list(devedores.columns).index('Tesouraria')]
            dev_falta = dev_lancamento[list(devedores.columns).index('Diferença')]
            dev_data = dev_lancamento[list(devedores.columns).index('Data Diferença')]
            dev_fk_id_cred = dev_lancamento[list(devedores.columns).index('FK_ID_CREDORES')]
            data_dif = cred_data - dev_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if dev_falta > 0 and cred_tece == dev_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= dev_falta:
                
                cred_trat_py += dev_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'DEVEDORES_TRAT_PY'] = cred_trat_py

                cred_saldo -= dev_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(dev_fk_id_cred)
                except:
                    status = False
                if status:
                    dev_fk_id_cred = id_cred
                    devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                    , 'FK_ID_CREDORES'] = dev_fk_id_cred
                else:
                    try:
                        dev_fk_id_cred = int(dev_fk_id_cred)
                    except:
                        pass
                    dev_fk_id_cred = f'{dev_fk_id_cred}, {id_cred}'
                    devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                    , 'FK_ID_CREDORES'] = dev_fk_id_cred

                devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                , 'TRATAMENTO_PY'] += dev_falta

                dev_falta -= dev_falta
                devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                , 'Diferença'] = dev_falta

                cont += 1
                print(cont)

            elif dev_falta > 0 and cred_tece == dev_tece and data_dif <= 60 and cred_saldo > 0 and cred_saldo < dev_falta:             
                
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'DEVEDORES_TRAT_PY'] = cred_trat_py

                dev_falta -= cred_saldo
                devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                , 'Diferença'] = dev_falta

                devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                , 'TRATAMENTO_PY'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(dev_fk_id_cred)
                except:
                    status = False
                if status:
                    dev_fk_id_cred = id_cred
                    devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                    , 'FK_ID_CREDORES'] = dev_fk_id_cred
                else:
                    try:
                        dev_fk_id_cred = int(dev_fk_id_cred)
                    except:
                        pass
                    dev_fk_id_cred = f'{dev_fk_id_cred}, {id_cred}'
                    devedores.at[devedores.index[devedores['ID_DEVEDORES'] == id_dev].tolist()[0]
                    , 'FK_ID_CREDORES'] = dev_fk_id_cred

                cont += 1
                print(cont)

            #if cont >= 10:
            #    break
qtd_lanc.append(cont)
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')


#%%
#---------- TRATAMENTO SECUNDÁRIO ERRO DE SAQUE ----------
print("---------- TRATAMENTO SECUNDÁRIO ERRO DE SAQUE ----------")
cont = 0
for cred_lancamento in credores.values:
    #if cont >= 10:
    #    break

    #variáveis
    cred_agencia = cred_lancamento[list(credores.columns).index('Nº dependência')]
    cred_tece = cred_lancamento[list(credores.columns).index('Tesouraria')]
    cred_saldo = cred_lancamento[list(credores.columns).index('Diferença')]
    cred_data = cred_lancamento[list(credores.columns).index('Data Diferença')]
    id_cred = cred_lancamento[list(credores.columns).index('ID_CREDORES')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRATAMENTO_2_PY')]
    if isnan(cred_agencia):
        cred_agencia = 0
    df = erro_saque.query(f'COD_DEPENDENCIA=={cred_agencia}').values

    if cred_saldo > 0 and cred_data <= pd.Timestamp('2022-07-11 23:59:59') and cred_agencia != 0 and not isnan(cred_agencia):
        
        for es_lancamento in df:

            #variáveis
            es_faixa = es_lancamento[list(erro_saque.columns).index('Controle')]
            es_motivo = es_lancamento[list(erro_saque.columns).index('MOTIVO')]
            es_agencia = es_lancamento[list(erro_saque.columns).index('COD_DEPENDENCIA')]
            es_data = es_lancamento[list(erro_saque.columns).index('INCLUSAO NA PLANILHA')]

            es_validacao = es_faixa == 549 and es_motivo not in('02 - CEDULA MUTILADA',
                '04 - CEDULA SUSPEITA',
                '03 - CEDULA DILACERADA',
                '05 - CEDULA ENTINTADA'
                ) and es_data <= pd.Timestamp('2022-06-30 23:59:59')
            id_es = es_lancamento[list(erro_saque.columns).index('ID_ES')]
            es_tece = es_lancamento[list(erro_saque.columns).index('TECE')]
            es_falta = es_lancamento[list(erro_saque.columns).index('DEVEDORES DEP.5100-5           FX.549')]
            es_fk_id_cred = es_lancamento[list(erro_saque.columns).index('DOC CREDORES dependência')]
            data_dif = cred_data - es_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if es_falta > 0 and cred_tece == es_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= es_falta and es_validacao:
                
                cred_trat_py += es_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'TRATAMENTO_2_PY'] = cred_trat_py

                cred_saldo -= es_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                    , 'DOC CREDORES dependência'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                    , 'DOC CREDORES dependência'] = es_fk_id_cred

                erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                , 'TRATAMENTO_2_PY'] += es_falta

                es_falta -= es_falta
                erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                , 'DEVEDORES DEP.5100-5           FX.549'] = es_falta

                cont += 1
                print(cont)

            elif es_falta > 0 and cred_tece == es_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < es_falta and es_validacao:             
                
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'TRATAMENTO_2_PY'] = cred_trat_py

                es_falta -= cred_saldo
                erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                , 'DEVEDORES DEP.5100-5           FX.549'] = es_falta

                erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                , 'TRATAMENTO_2_PY'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(es_fk_id_cred)
                except:
                    status = False
                if status:
                    es_fk_id_cred = id_cred
                    erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                    , 'DOC CREDORES dependência'] = es_fk_id_cred
                else:
                    try:
                        es_fk_id_cred = int(es_fk_id_cred)
                    except:
                        pass
                    es_fk_id_cred = f'{es_fk_id_cred}, {id_cred}'
                    erro_saque.at[erro_saque.index[erro_saque['ID_ES'] == id_es].tolist()[0]
                    , 'DOC CREDORES dependência'] = es_fk_id_cred

                cont += 1
                print(cont)

            #if cont >= 10:
            #    break
qtd_lanc.append(cont)
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')


#%%
#---------- TRATAMENTO SECUNDÁRIO ERRO DE DEPÓSITO ----------
print("---------- TRATAMENTO SECUNDÁRIO ERRO DE DEPÓSITO ----------")
cont = 0
for cred_lancamento in credores.values:
    #if cont >= 10:
    #    break

    #variáveis
    cred_agencia = cred_lancamento[list(credores.columns).index('Nº dependência')]
    cred_tece = cred_lancamento[list(credores.columns).index('Tesouraria')]
    cred_saldo = cred_lancamento[list(credores.columns).index('Diferença')]
    cred_data = cred_lancamento[list(credores.columns).index('Data Diferença')]
    id_cred = cred_lancamento[list(credores.columns).index('ID_CREDORES')]
    cred_trat_py = cred_lancamento[list(credores.columns).index('TRATAMENTO_2_PY')]
    if isnan(cred_agencia):
        cred_agencia = 0
    df = erro_deposito.query(f'COD_DEPENDENCIA=={cred_agencia}').values

    if cred_saldo > 0 and cred_data <= pd.Timestamp('2022-07-11 23:59:59') and cred_agencia != 0 and not isnan(cred_agencia):
        
        for ed_lancamento in df:

            #variáveis
            id_ed = ed_lancamento[list(erro_deposito.columns).index('ID_ED')]
            ed_agencia = ed_lancamento[list(erro_deposito.columns).index('COD_DEPENDENCIA')]
            ed_tece = ed_lancamento[list(erro_deposito.columns).index('TECE')]
            ed_falta = ed_lancamento[list(erro_deposito.columns).index('DEVEDORES DEP.5100-5           FX.6864-3    CI.716-4')]
            ed_data = ed_lancamento[list(erro_deposito.columns).index('DATA DO DEPÓSITO')]
            ed_fk_id_cred = ed_lancamento[list(erro_deposito.columns).index('DOC CREDORES dependência')]

            data_dif = cred_data - ed_data
            data_dif = pd.Timedelta(data_dif, 'day').days

            if ed_falta > 0 and cred_tece == ed_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo >= ed_falta and ed_data <= pd.Timestamp('2022-06-30 23:59:59'):
                
                cred_trat_py += ed_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'TRATAMENTO_2_PY'] = cred_trat_py

                cred_saldo -= ed_falta
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                    , 'DOC CREDORES dependência'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                    , 'DOC CREDORES dependência'] = ed_fk_id_cred

                erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                , 'TRATAMENTO_2_PY'] += ed_falta

                ed_falta -= ed_falta
                erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                , 'DEVEDORES DEP.5100-5           FX.6864-3    CI.716-4'] = ed_falta

                cont += 1
                print(cont)

            elif ed_falta > 0 and cred_tece == ed_tece and 0 <= data_dif <= 60 and cred_saldo > 0 and cred_saldo < ed_falta and ed_data <= pd.Timestamp('2022-06-30 23:59:59'):             
                
                cred_trat_py += cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'TRATAMENTO_2_PY'] = cred_trat_py

                ed_falta -= cred_saldo
                erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                , 'DEVEDORES DEP.5100-5           FX.6864-3    CI.716-4'] = ed_falta

                erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                , 'TRATAMENTO_2_PY'] += cred_saldo

                cred_saldo -= cred_saldo
                credores.at[credores.index[credores['ID_CREDORES'] == id_cred].tolist()[0]
                , 'Diferença'] = cred_saldo

                try:
                    status = isnan(ed_fk_id_cred)
                except:
                    status = False
                if status:
                    ed_fk_id_cred = id_cred
                    erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                    , 'DOC CREDORES dependência'] = ed_fk_id_cred
                else:
                    try:
                        ed_fk_id_cred = int(ed_fk_id_cred)
                    except:
                        pass
                    ed_fk_id_cred = f'{ed_fk_id_cred}, {id_cred}'
                    erro_deposito.at[erro_deposito.index[erro_deposito['ID_ED'] == id_ed].tolist()[0]
                    , 'DOC CREDORES dependência'] = ed_fk_id_cred

                cont += 1
                print(cont)

            #if cont >= 10:
            #    break
qtd_lanc.append(cont)
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')


# %%
#---------- REORGANIZANDO COLUNAS / CRIANDO RESUMO ----------
#credores.columns
print('{:-^57}'.format(' REORGANIZANDO COLUNAS / CRIANDO RESUMO '))
credores = credores[['ID_CREDORES', 'Transportadora', 'Tesouraria', 'Data Diferença',
       'Documento', 'Nº dependência', 'ATM', 'Dependência', 'Valor',
       'Erro de Saque', 'Valor  compensado', 'DEVEDORES_TRAT_PY', 'TRATAMENTO_1_PY', 'TRATAMENTO_2_PY', 'Diferença', 'Tipo de Diferença',
       'Data reg. erro de saque', 'Data Compensação', 'Descrição',
       'TRANSFERÊNCIA CREDORES CI 651', 'DATA TRANSFERENCIA'
       ]]
erro_saque = erro_saque[['ID_ES', 'FK_ID_CREDORES', 'Controle', 'MÊS REFERÊNCIA', 'INCLUSAO NA PLANILHA', 'DATA DO SAQUE',
       'HORA DO SAQUE', 'DEMANDA', 'TECE', 'COD_DEPENDENCIA', 'DEPENDENCIA',
       'ATM', 'VALOR DA DIFERENÇA', 'MOTIVO', 'DATA BATIMENTO ATM',
       'VALOR APURADO', 'VALOR COMPENSADO', 'TRATAMENTO_1_PY', 'TRATAMENTO_2_PY', 'SITUAÇÃO DA DIFERENÇA',
       'DATA DA CONTABILIZAÇÃO', 'DEVEDORES DEP.5100-5           FX.549',
       'Contabilizado', 'DOC  CREDORES', 'OBS.', 'REG. POR DEPENDÊNCIA',
       'DOC CREDORES dependência', 'DATA REG. DEPENDÊNCIA',
       'TRANSFERÊNCIA PARA ADM. CENTRAL- 6864-3 / C.I - 491-2',
       'DATA TRANSFERÊNCIA', 'Descrição LCD'
       ]]
erro_deposito = erro_deposito[['ID_ED', 'FK_ID_CREDORES', 'MÊS REFERÊNCIA', 'INCLUSAO PLANILHA', 'DATA DO DEPÓSITO',
       'HORA DO DEPÓSITO', 'DEMANDA', 'TECE', 'COD_DEPENDENCIA', 'DEPENDENCIA',
       'ATM', 'VALOR DA DIFERENÇA', 'DATA BATIMENTO ATM', 'VALOR APURADO',
       'VALOR COMPENSADO', 'TRATAMENTO_1_PY', 'TRATAMENTO_2_PY', 'SITUAÇÃO DA DIFERENÇA', 'DATA DA CONTABILIZAÇÃO',
       'DEVEDORES DEP.5100-5           FX.6864-3    CI.716-4',
       'CONTABILIZADO?', 'DATA CREDORA',
       'APURAÇÃO CREDORA TERMINAIS  FX 5885-9', 'DOCUMENTO CREDORES', 'OBS.',
       'REG. POR DEPENDÊNCIA', 'DATA REG. DEPENDÊNCIA', 'DOC CREDORES dependência',
       'TRANSFERÊNCIA PARA ADM. CENTRAL- 6864-3 / C.I - 491-2',
       'DATA TRANSFERÊNCIA', 'descrição lançamento'
       ]]
devedores = devedores[['ID_DEVEDORES', 'FK_ID_CREDORES', 'Transportadora', 'Tesouraria', 'Duplicidade?', 'Data Diferença',
       'Data lançamento CAB', 'Número documento CAB', 'Nº dependência',
       'ATM', 'Dependência', 'Valor', 'Valor ressarcido / compensado', 'TRATAMENTO_PY',
       'Diferença', 'Data notificação', 'Tipo de diferença', 'Tipo de Acerto',
       'Data Contabilização', 'Contestação', 'Dias Corridos',
       'Status Ressarcimento'
       ]]

locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')
resumo_df = pd.DataFrame(
            {'Valor': [locale.currency(erro_saque['TRATAMENTO_1_PY'].sum(), grouping=True)
                ,locale.currency(erro_saque['TRATAMENTO_2_PY'].sum(), grouping=True)
                ,locale.currency(erro_deposito['TRATAMENTO_1_PY'].sum(), grouping=True)
                ,locale.currency(erro_deposito['TRATAMENTO_2_PY'].sum(), grouping=True)
                ,locale.currency(devedores['TRATAMENTO_PY'].sum(), grouping=True)
                ,locale.currency(credores['DEVEDORES_TRAT_PY'].sum() + credores['TRATAMENTO_1_PY'].sum() + credores['TRATAMENTO_2_PY'].sum(), grouping=True)]
            ,'Qtd Registros': [qtd_lanc[1]
                ,qtd_lanc[4]
                ,qtd_lanc[2]
                ,qtd_lanc[5]
                ,qtd_lanc[3] + qtd_lanc[0]
                ,sum(qtd_lanc)]
            }, index=['Trat. Primário ES', 'Trat. Secundário ES', 'Trat. Primário ED', 'Trat. Secundário ED', 'Trat. Devedores', 'TOTAL'])
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')

#%%
#---------- SALVANDO ARQUIVO RESULTADO ----------
print('{:-^57}'.format(' SALVANDO ARQUIVO RESULTADO '))
arquivo = pd.ExcelWriter(path='K:/GSAS/09 - Coordenacao Gestao Numerario/03 Gestao de Diferencas/CONCILIAÇÕES PYTHON/Resultado em Homologação.xlsx',engine='xlsxwriter')
resumo_df.to_excel(arquivo, sheet_name="RESUMO")
credores.to_excel(arquivo, sheet_name="Credores_Final", index=False)
erro_saque.to_excel(arquivo, sheet_name="Erro_Saque_Final", index=False)
erro_deposito.to_excel(arquivo, sheet_name="Erro_Deposito_Final", index=False)
devedores.to_excel(arquivo, sheet_name="Devedores_Final", index=False)
arquivo.save()
arquivo.close()
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')
