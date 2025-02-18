#%%
# ---------- IMPORTANDO PACOTES ----------
import pandas as pd
from math import isnan
import locale
import warnings
warnings.simplefilter("ignore")

# %%
#---------- EXTRAINDO DADOS EM SEUS CAMINHOS ----------
credores = pd.read_excel('GESTÃO CREDORES - 7914-5.xlsm', skiprows=3, sheet_name='Credores')
erro_saque = pd.read_excel('NOVO_ERRO DE SAQUE.xlsm', skiprows=4, sheet_name='ERROS DE SAQUE')

# %%
#---------- INSERÇÃO DA COLUNA DE TRATAMENTO PRIMÁRIO ----------
credores['TRATAMENTO_1_PY'] = [0 for x in range(len(credores))]
credores['DEVEDORES_TRAT_PY'] = [0 for x in range(len(credores))]
erro_saque['TRATAMENTO_1_PY'] = [0 for x in range(len(erro_saque))]


#%%
#---------- INSERÇÃO DA COLUNA DE TRATAMENTO SECUNDÁRIO ----------
credores['TRATAMENTO_2_PY'] = [0 for x in range(len(credores))]
erro_saque['TRATAMENTO_2_PY'] = [0 for x in range(len(erro_saque))]
qtd_lanc = list()


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


# %%
#---------- REORGANIZANDO COLUNAS / CRIANDO RESUMO ----------
#credores.columns
print('{:-^57}'.format(' REORGANIZANDO COLUNAS / CRIANDO RESUMO '))
credores = credores[['ID_CREDORES', 'Transportadora', 'Tesouraria', 'Data Diferença',
       'Documento', 'Nº dependência', 'ATM', 'Dependência', 'Valor',
       'Erro de Saque', 'Valor  compensado', 'DEVEDORES_TRAT_PY', 'TRATAMENTO_1_PY', 'TRATAMENTO_2_PY', 'Diferença', 'Tipo de Diferença',
       'Data contabilização', 'Data Compensação', 'Descrição',
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


#%%
#---------- SALVANDO ARQUIVO RESULTADO ----------
print('{:-^57}'.format(' SALVANDO ARQUIVO RESULTADO '))
arquivo = pd.ExcelWriter(path='Resultado em Homologação.xlsx',engine='xlsxwriter')
credores.to_excel(arquivo, sheet_name="Credores_Final", index=False)
erro_saque.to_excel(arquivo, sheet_name="Erro_Saque_Final", index=False)
arquivo.save()
arquivo.close()
print('{:-^57}'.format(' ETAPA FINALIZADA '), end='\n\n')
