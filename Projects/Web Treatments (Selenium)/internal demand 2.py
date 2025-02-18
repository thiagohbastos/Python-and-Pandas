#%%
import warnings
import pyodbc
import pandas as pd
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from time import sleep
from datetime import date
warnings.simplefilter("ignore")


#%%
conn_ods = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server=ods;"
    "Database=DBZB099D;"
    "Trusted_Connection=yes;")

#%%
query_dmd = """SELECT  TRIM(D.NOM_CLI) AS NOM_CLI
       ,D.NUM_DMD
       ,TRIM(D.CTD_LNK_DMD) AS CTD_LNK_DMD
	   ,CAST(D.DTA_ABE AS DATE) AS DTA_ABE
	   ,R.DES_RTA
FROM DEMANDA_ABERTA D with
(nolock
)
LEFT JOIN RESPOSTA_DMD_ABERT R
ON D.NUM_DMD = R.NUM_DMD
WHERE DES_TIP_AST_DMD = 'AUTOATENDIMENTO : ERRO : SAQUE : RESSARCIMENTO AUTOMATICO'
AND DES_STT_DMD = 'Designada'
AND R.IDT_IFO IN (7)
--ORDER BY DTA_ABE DESC"""

tabela = pd.read_sql_query(query_dmd, conn_ods)

#%%
navegador = Chrome("K:/GSAS/09 - Coordenacao Gestao Numerario/09 Prototipos SSIS/B042786/__PYTHON__/chromedriver.exe")
navegador.maximize_window()

plataforma = r'https://plataformamb.mercantil.com.br/minha-agencia/gestao-ressarcimento'

msg_atendimento = 'Cliente ressarcido.'

#%%
#navegador.get(plataforma)
atualizar_pla = dict()
lista_nome = []
lista_demanda = []
lista_valor = []
lista_data = []
lista_tipo = []
qtd_erro_saq = 0
qtd_atendidas = 0
temp = 0
for registro in tabela.values:
    nome_cliente = registro[0]
    caminho = registro[2]
    caminho = caminho.replace('\\', '/')
    demanda = registro[1]
    data_abe = registro[3]
    valor_ress = float(registro[4].split(',')[0])

    #navegador.switch_to.window(navegador.window_handles[0])

    #navegador.switch_to.new_window('tab')
    cont = 0
    while True:
        try: 
            navegador.get(caminho)
            break
        except:
            sleep(0.5)
            cont += 0.5
            if cont >= 10:
                break
    if cont >= 10:
        continue
    
    iframe_formulario = navegador.find_element(By.XPATH, '//*[@id="form_conteudo"]')
    navegador.switch_to.frame(iframe_formulario)
    tipo_dmd = navegador.find_element(By.XPATH, '//*[@id="xSec3_1"]/table[2]/tbody/tr[10]/td/font[2]').text.upper()

    if 'ERRO' in tipo_dmd:
        qtd_erro_saq += 1
        #navegador.close()
        continue
    qtd_atendidas += 1

    #navegador.switch_to.window(navegador.window_handles[0])

    navegador.execute_script(script='''window.top.frames[1].document.getElementById('btnEditarFormBotoes').click();window.top.frames[2].document.getElementById('btnEditar').click()''')

    cont = 0
    while True:
        try: 
            navegador.switch_to.alert.accept()
            break
        except:
            sleep(0.5)
            cont += 0.5
            if cont >= 10:
                break
    if cont >= 10:
        continue

    lista_nome.append(nome_cliente)
    lista_demanda.append(demanda)
    lista_valor.append(valor_ress)
    lista_data.append(data_abe)
    lista_tipo.append(tipo_dmd)
    
    navegador.find_element(By.XPATH, '/html/body/form/table[3]/tbody/tr[8]/td/textarea').send_keys(msg_atendimento)

    x = navegador.find_element(By.XPATH,'/html/body/form/table[3]/tbody/tr[4]/td[2]/font/select')
    Select(x).select_by_value('_j8dnmq82jdtm7b1u6dsg58rrkc5m0_') #Com solução total
                              
    navegador.execute_script(script='''window.top.frames[2].document.getElementById('btnAtender').click()''')

atualizar_pla['Nomes'] = lista_nome
atualizar_pla['Demanda'] = lista_demanda
atualizar_pla['Data'] = lista_data
atualizar_pla['Valor'] = lista_valor
atualizar_pla['Tipo'] = lista_tipo

df_atualizar_pla = pd.DataFrame(atualizar_pla)


#%%
arquivo = pd.ExcelWriter(path=f'K:/GSAS/09 - Coordenacao Gestao Numerario/09 Prototipos SSIS/B042786/__PYTHON__/DMD ERRO SAQUE AUTOMÁTICO/APROVAR PLA/Aprovar PLA {date.today().day}-{date.today().month}-{date.today().year}.xlsx',engine='xlsxwriter')
df_atualizar_pla.to_excel(arquivo, sheet_name="REGISTROS", index=False)
arquivo.save()
arquivo.close()


#%%
navegador.quit()
conn_ods.close()

# %%
