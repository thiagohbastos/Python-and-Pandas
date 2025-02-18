#%%
#---------- IMPORTANDO PACOTES ----------
import pandas as pd
from math import isnan
import warnings
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from time import sleep
warnings.simplefilter("ignore")

# %%
#---------- TRAZENDO e TRATANDO BASES A LANÇAR ----------
#ERROS DE SAQUE PRIMÁRIOS E SECUNDÁRIOS
lanc_ES = pd.read_excel('Resultado em Homologação.xlsx', sheet_name='Erro_Saque_Final')[
    ['ID_ES', 'FK_ID_CREDORES', 'DOC CREDORES dependência', 'TRATAMENTO_1_PY', 
    'TRATAMENTO_2_PY', 'TECE', 'ATM', 'DATA DO SAQUE', 'VALOR DA DIFERENÇA']]
lanc_ES = lanc_ES.query('TRATAMENTO_1_PY>0 or TRATAMENTO_2_PY>0')
lanc_ES['LANCADO'] = ['NÃO' for x in range(len(lanc_ES))]
lanc_ES_2 = lanc_ES.query('TRATAMENTO_2_PY>0')

#ERROS DE DEPÓSITO PRIMÁRIOS E SECUNDÁRIOS
lanc_ED = pd.read_excel('Resultado em Homologação.xlsx', sheet_name='Erro_Deposito_Final')[
    ['ID_ED', 'FK_ID_CREDORES', 'DOC CREDORES dependência', 'TRATAMENTO_1_PY', 
    'TRATAMENTO_2_PY', 'TECE', 'ATM', 'DATA DO DEPÓSITO', 'VALOR DA DIFERENÇA']]
lanc_ED = lanc_ED.query('TRATAMENTO_1_PY>0 or TRATAMENTO_2_PY>0')
lanc_ED['LANCADO'] = ['NÃO' for x in range(len(lanc_ED))]
lanc_ED_2 = lanc_ED.query('TRATAMENTO_2_PY>0')

#DEVEDORES
lanc_DEV = pd.read_excel('Resultado em Homologação.xlsx', sheet_name='Devedores_Final')[
    ['ID_DEVEDORES', 'FK_ID_CREDORES', 'TRATAMENTO_PY', 'Tesouraria',
    'ATM', 'Data Diferença', 'Valor']]
lanc_DEV = lanc_DEV.query('TRATAMENTO_PY>0')
lanc_DEV['LANCADO'] = ['NÃO' for x in range(len(lanc_DEV))]

# %%
#---------- ABRINDO NAVEGADOR PARA INICIAR LANÇAMENTOS ----------
navegador = Chrome()
navegador.maximize_window()
navegador.get('https://intranet2.mercantil.com.br/MB.MVC.UI.LCD.LancamentoContabilDescentralizado/Lancamento/Incluir?opcao=2,1')

# %%
#---------- INICIANDO LANÇAMENTOS ES PRIMÁRIOS ----------
cont = erros_es = 0
for lancamento in lanc_ES.values:
    #VARIÁVEIS
    valor = lancamento[list(lanc_ES.columns).index('TRATAMENTO_1_PY')]
    id_es = lancamento[list(lanc_ES.columns).index("ID_ES")]
    complemento = f'ES{id_es}/CRED{lancamento[list(lanc_ES.columns).index("FK_ID_CREDORES")]}'
    indivDebito = lancamento[list(lanc_ES.columns).index("TECE")]
    es_status = lancamento[list(lanc_ES.columns).index("LANCADO")]

    if count == 10:
        break

    if es_status in ['NÃO', 'ERRO']:
        #navegador.refresh()
        sleep(2)
        #EMPRESA
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('02', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #PROCESSAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IndicadorProcessamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('N - ', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #NATUREZA - Nº DO LANÇAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdNaturezaLancamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('017', Keys.ENTER)
                break
            except:
                sleep(0.5)
        
        while True:
            doc = navegador.find_element(By.XPATH, '//*[@id="NumeroDocumento"]').get_attribute('value')
            if doc == '':
                sleep(0.3)
            else:
                sleep(0.3)
                break
        
        #VALOR
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="Valor"]').send_keys(f'{valor}00')
                break
            except:
                sleep(0.5)

        #COMPLEMENTO HISTÓRICO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="ComplementoHistorico"]').send_keys(complemento)
                break
            except:
                sleep(0.5)

        #TECE DE DÉBITO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="IdIndivDebito"]').send_keys(indivDebito)
                break
            except:
                sleep(0.5)

        x = 0
        while True:
            sleep(1)
            try:
                navegador.find_element(By.XPATH, '//*[@id="Incluir"]').click()
            except:
                pass
            msg = navegador.find_element(By.XPATH, '//*[@id="tab-incluir-lancamento"]/div[1]').text
            validador = msg == 'INCLUSAO EFETUADA COM SUCESSO.' or msg == ''
            if validador:
                lanc_ES.at[lanc_ES.index[lanc_ES['ID_ES'] == id_es].tolist()[0], 'LANCADO'] = 'SIM'
                cont += 1
                print(f'{cont}/{lanc_ES["ATM"].count()} -- ERROS: {erros_es}')
                sleep(1)
                break
            else:
                x += 1
                if x > 3:
                    print(msg)
                    erros_es += 1
                    lanc_ES.at[lanc_ES.index[lanc_ES['ID_ES'] == id_es].tolist()[0], 'LANCADO'] = 'ERRO'
                    print(f'{cont}/{lanc_ES["ATM"].count()} -- ERROS: {erros_es}')
                    while True:
                        try:
                            navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                            navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('Selecione', Keys.ENTER)
                            break
                        except:
                            sleep(0.5)
                    break

# %%
#---------- INICIANDO LANÇAMENTOS ES SECUNDÁRIOS ----------
cont = erros_es_2 = 0
for lancamento in lanc_ES_2.values:
    #VARIÁVEIS
    valor = lancamento[list(lanc_ES_2.columns).index('TRATAMENTO_2_PY')]
    id_es_2 = lancamento[list(lanc_ES_2.columns).index("ID_ES")]
    complemento = f'ES{id_es_2}/CRED{lancamento[list(lanc_ES_2.columns).index("DOC CREDORES dependência")]}'
    indivDebito = lancamento[list(lanc_ES_2.columns).index("TECE")]
    es_status_2 = lancamento[list(lanc_ES_2.columns).index("LANCADO")]

    if count == 10:
        break

    if es_status_2 in ['NÃO', 'ERRO']:
        #navegador.refresh()
        sleep(2)
        #EMPRESA
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('02', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #PROCESSAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IndicadorProcessamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('N - ', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #NATUREZA - Nº DO LANÇAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdNaturezaLancamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('017', Keys.ENTER)
                break
            except:
                sleep(0.5)

        while True:
            doc = navegador.find_element(By.XPATH, '//*[@id="NumeroDocumento"]').get_attribute('value')
            if doc == '':
                sleep(0.3)
            else:
                sleep(0.3)
                break

        #VALOR
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="Valor"]').send_keys(f'{valor}00')
                break
            except:
                sleep(0.5)

        #COMPLEMENTO HISTÓRICO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="ComplementoHistorico"]').send_keys(complemento)
                break
            except:
                sleep(0.5)

        #TECE DE DÉBITO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="IdIndivDebito"]').send_keys(indivDebito)
                break
            except:
                sleep(0.5)

        x = 0
        while True:
            sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="Incluir"]').click()
            msg = navegador.find_element(By.XPATH, '//*[@id="tab-incluir-lancamento"]/div[1]').text
            validador = msg == 'INCLUSAO EFETUADA COM SUCESSO.'
            if validador:
                lanc_ES_2.at[lanc_ES_2.index[lanc_ES_2['ID_ES'] == id_es_2].tolist()[0], 'LANCADO'] = 'SIM'
                cont += 1
                print(f'{cont}/{lanc_ES_2["ATM"].count()} -- ERROS: {erros_es_2}')
                break
            else:
                x += 1
                if x > 3:
                    print(msg)
                    erros_es_2 += 1
                    lanc_ES_2.at[lanc_ES_2.index[lanc_ES_2['ID_ES'] == id_es_2].tolist()[0], 'LANCADO'] = 'ERRO'
                    print(f'{cont}/{lanc_ES_2["ATM"].count()} -- ERROS: {erros_es_2}')
                    while True:
                        try:
                            navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                            navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('Selecione', Keys.ENTER)
                            break
                        except:
                            sleep(0.5)
                    break

# %%
#---------- INICIANDO LANÇAMENTOS ED PRIMÁRIOS ----------
cont = erros_ed = 0
for lancamento in lanc_ED.values:
    #VARIÁVEIS
    valor = lancamento[list(lanc_ED.columns).index('TRATAMENTO_1_PY')]
    id_ed = lancamento[list(lanc_ED.columns).index("ID_ED")]
    complemento = f'ED{id_ed}/CRED{lancamento[list(lanc_ED.columns).index("FK_ID_CREDORES")]}'
    indivDebito = lancamento[list(lanc_ED.columns).index("TECE")]
    ed_status = lancamento[list(lanc_ED.columns).index("LANCADO")]

    if count == 10:
        break

    if ed_status in ['NÃO', 'ERRO']:
        sleep(2)
        #navegador.refresh()
        #EMPRESA
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('02', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #PROCESSAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IndicadorProcessamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('N - ', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #NATUREZA - Nº DO LANÇAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdNaturezaLancamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('022', Keys.ENTER)
                break
            except:
                sleep(0.5)

        while True:
            doc = navegador.find_element(By.XPATH, '//*[@id="NumeroDocumento"]').get_attribute('value')
            if doc == '':
                sleep(0.3)
            else:
                sleep(0.3)
                break

        #VALOR
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="Valor"]').send_keys(f'{valor}00')
                break
            except:
                sleep(0.5)

        #COMPLEMENTO HISTÓRICO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="ComplementoHistorico"]').send_keys(complemento)
                break
            except:
                sleep(0.5)

        #TECE DE DÉBITO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="IdIndivDebito"]').send_keys(indivDebito)
                break
            except:
                sleep(0.5)

        x = 0
        while True:
            sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="Incluir"]').click()
            msg = navegador.find_element(By.XPATH, '//*[@id="tab-incluir-lancamento"]/div[1]').text
            validador = msg == 'INCLUSAO EFETUADA COM SUCESSO.'
            if validador:
                lanc_ED.at[lanc_ED.index[lanc_ED['ID_ED'] == id_ed].tolist()[0], 'LANCADO'] = 'SIM'
                cont += 1
                print(f'{cont}/{lanc_ED["ATM"].count()} -- ERROS: {erros_ed}')
                break
            else:
                x += 1
                if x > 3:
                    print(msg)
                    erros_ed += 1
                    lanc_ED.at[lanc_ED.index[lanc_ED['ID_ED'] == id_ed].tolist()[0], 'LANCADO'] = 'ERRO'
                    print(f'{cont}/{lanc_ED["ATM"].count()} -- ERROS: {erros_ed}')
                    while True:
                        try:
                            navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                            navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('Selecione', Keys.ENTER)
                            break
                        except:
                            sleep(0.5)
                    break

# %%
#---------- INICIANDO LANÇAMENTOS ED SECUNDÁRIOS ----------
cont = erros_ed_2 = 0
for lancamento in lanc_ED_2.values:
    #VARIÁVEIS
    valor = lancamento[list(lanc_ED_2.columns).index('TRATAMENTO_2_PY')]
    id_ed_2 = lancamento[list(lanc_ED_2.columns).index("ID_ED")]
    complemento = f'ED{id_ed_2}/CRED{lancamento[list(lanc_ED_2.columns).index("DOC CREDORES dependência")]}'
    indivDebito = lancamento[list(lanc_ED_2.columns).index("TECE")]
    ed_status_2 = lancamento[list(lanc_ED_2.columns).index("LANCADO")]

    if count == 10:
        break

    if ed_status_2 in ['NÃO', 'ERRO']:
        #navegador.refresh()
        sleep(2)
        #EMPRESA
        navegador.refresh()
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('02', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #PROCESSAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IndicadorProcessamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('N - ', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #NATUREZA - Nº DO LANÇAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdNaturezaLancamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('022', Keys.ENTER)
                break
            except:
                sleep(0.5)

        while True:
            doc = navegador.find_element(By.XPATH, '//*[@id="NumeroDocumento"]').get_attribute('value')
            if doc == '':
                sleep(0.3)
            else:
                sleep(0.3)
                break

        #VALOR
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="Valor"]').send_keys(f'{valor}00')
                break
            except:
                sleep(0.5)

        #COMPLEMENTO HISTÓRICO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="ComplementoHistorico"]').send_keys(complemento)
                break
            except:
                sleep(0.5)

        #TECE DE DÉBITO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="IdIndivDebito"]').send_keys(indivDebito)
                break
            except:
                sleep(0.5)

        x = 0
        while True:
            sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="Incluir"]').click()
            msg = navegador.find_element(By.XPATH, '//*[@id="tab-incluir-lancamento"]/div[1]').text
            validador = msg == 'INCLUSAO EFETUADA COM SUCESSO.'
            if validador:
                lanc_ED_2.at[lanc_ED_2.index[lanc_ED_2['ID_ED'] == id_ed_2].tolist()[0], 'LANCADO'] = 'SIM'
                cont += 1
                print(f'{cont}/{lanc_ED_2["ATM"].count()} -- ERROS: {erros_ed_2}')
                break
            else:
                x += 1
                if x > 3:
                    print(msg)
                    erros_ed_2 += 1
                    lanc_ED_2.at[lanc_ED_2.index[lanc_ED_2['ID_ED'] == id_ed_2].tolist()[0], 'LANCADO'] = 'ERRO'
                    print(f'{cont}/{lanc_ED_2["ATM"].count()} -- ERROS: {erros_ed_2}')
                    while True:
                        try:
                            navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                            navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('Selecione', Keys.ENTER)
                            break
                        except:
                            sleep(0.5)
                    break

#%%
#---------- INICIANDO LANÇAMENTOS ED PRIMÁRIOS ----------
cont = erros_dev = 0
for lancamento in lanc_DEV.values:
    #VARIÁVEIS
    valor = lancamento[list(lanc_DEV.columns).index('TRATAMENTO_PY')]
    id_dev = lancamento[list(lanc_DEV.columns).index("ID_DEVEDORES")]
    complemento = f'DEV{id_dev}/CRED{lancamento[list(lanc_DEV.columns).index("FK_ID_CREDORES")]}'
    indivDebito = lancamento[list(lanc_DEV.columns).index("Tesouraria")]
    dev_status = lancamento[list(lanc_DEV.columns).index('LANCADO')]

    if count == 10:
        break

    if dev_status in ['NÃO', 'ERRO']:
        #navegador.refresh()
        sleep(2)
        #EMPRESA
        #navegador.refresh()
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('02', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #PROCESSAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IndicadorProcessamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('N - ', Keys.ENTER)
                break
            except:
                sleep(0.5)

        #NATUREZA - Nº DO LANÇAMENTO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="select2-IdNaturezaLancamento-container"]').click()
                navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('002', Keys.ENTER)
                break
            except:
                sleep(0.5)

        while True:
            doc = navegador.find_element(By.XPATH, '//*[@id="NumeroDocumento"]').get_attribute('value')
            if doc == '':
                sleep(0.3)
            else:
                sleep(0.3)
                break

        #VALOR
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="Valor"]').send_keys(f'{valor}00')
                break
            except:
                sleep(0.5)

        #COMPLEMENTO HISTÓRICO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="ComplementoHistorico"]').send_keys(complemento)
                break
            except:
                sleep(0.5)

        #TECE DE DÉBITO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="IdIndivDebito"]').send_keys(indivDebito)
                break
            except:
                sleep(0.5)
        sleep(0.5)
        #CRÉDITO
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="IdIndivCredito"]').send_keys(indivDebito)
                break
            except:
                sleep(0.5)

        x = 0
        while True:
            sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="Incluir"]').click()
            msg = navegador.find_element(By.XPATH, '//*[@id="tab-incluir-lancamento"]/div[1]').text
            validador = msg == 'INCLUSAO EFETUADA COM SUCESSO.'
            if validador:
                lanc_DEV.at[lanc_DEV.index[lanc_DEV['ID_DEVEDORES'] == id_dev].tolist()[0], 'LANCADO'] = 'SIM'
                cont += 1
                print(f'{cont}/{lanc_DEV["ATM"].count()} -- ERROS: {erros_dev}')
                break
            else:
                x += 1
                if x > 3:
                    print(msg)
                    erros_dev += 1
                    lanc_DEV.at[lanc_DEV.index[lanc_DEV['ID_DEVEDORES'] == id_dev].tolist()[0], 'LANCADO'] = 'ERRO'
                    print(f'{cont}/{lanc_DEV["ATM"].count()} -- ERROS: {erros_dev}')
                    while True:
                        try:
                            navegador.find_element(By.XPATH, '//*[@id="select2-IdentificadorEmpresa-container"]').click()
                            navegador.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('Selecione', Keys.ENTER)
                            break
                        except:
                            sleep(0.5)
                    break
