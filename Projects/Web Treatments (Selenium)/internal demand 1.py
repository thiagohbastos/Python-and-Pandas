#%%
# ------------- IMPORTANDO BIBLIOTECAS -------------
#pyinstaller --name="DMD MANUTENÇÃO ATM" --onefile "Respondendo_DMD.py"
import warnings
import pyodbc
import pandas as pd
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from time import sleep
warnings.simplefilter("ignore")

#%%
# ------------- CRIANDO CONEXÃO SQLGDNP/GNU -------------
conn_gdnp = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server=sqlgdnp;"
    "Database=GNU;"
    "Trusted_Connection=yes;")



#%%
# ------------- BUSCANDO DEMANDAS PARA ATENDIMENTO -------------
query_dmd = """SELECT  A.NUM_DMD
       ,A.CTD_LNK_DMD
       /*,A.IDT_TML
	   ,B.STATUS_TML_ATUAL
	   ,B.SALDO_TERMINAL_ATUAL
	   ,B.SITUACAO_TML*/
	   ,CASE WHEN B.STATUS_TML_ATUAL = 'INDISPONIVEL' 
	   AND B.SALDO_TERMINAL_ATUAL >= 80000
	   AND (B.SITUACAO_GAA NOT LIKE '%DISPENSADOR COM O CASSETE DE REJEIÇÃO%' AND B.SITUACAO_GAA NOT LIKE '%CASSETE VALIDADOR CHEIO%') 
	   THEN 0 ELSE 1
	   END AS ATENDE
FROM [MERCANTIL\B040466].TC_DMD_MANUT A WITH (NOLOCK)
LEFT JOIN [MERCANTIL\B038660].TI_PRINCIPAL_TERMINAIS B WITH (NOLOCK)
ON A.IDT_TML = B.IDT_TML
WHERE GR_ATUAL like 'GR_Retaguarda_Numerario%'
AND DES_STT_DMD != 'Atendida'
--AND A.DTA_ABE < '2023-11-09'
ORDER BY 3 DESC, 1
"""

tabela = pd.read_sql_query(query_dmd, conn_gdnp)

#%%
# ------------- ABRINDO E CONFIGURANDO O NAVEGADOR -------------
navegador = Chrome()
navegador.maximize_window()

msg_atendimento = '''
Prezados, boa tarde.

O terminal será abastecido na próxima programação.

Atenciosamente,

Numerário
'''

msg_redir = '''
Prezados, boa tarde.

Terminal possui saldo e apresenta status de indisponível, gentileza reavaliar.

Atenciosamente,

Numerário
'''


#%%
# ------------- DEFININDO FUNÇÕES -------------
def atende_dmd():
    atd = dev = erros = ja_atd = 0
    dmd_atd = []
    dmd_dev = []
    dmd_erros = []
    dmd_ja_atd = []
    resumo = {}
    for registro in tabela.values:
        try:
            caminho = registro[1]
            demanda = registro[0]
            direcao = registro[2]
            cont = 0
            while True:
                try: 
                    navegador.get(caminho)
                    break
                except Exception as erro:
                    sleep(0.5)
                    cont += 0.5
                    if cont >= 10:
                        erros += 1
                        dmd_erros.append([demanda, erro])
                        print(f'{erros}º erro. Página da {demanda} demanda não carregou.')
                        break
            if cont >= 10:
                continue

            sleep(0.5)
            iframe_formulario = navegador.find_element(By.XPATH, '//*[@id="form_conteudo"]')
            navegador.switch_to.frame(iframe_formulario)
            status = navegador.find_element(By.XPATH, '/html/body/form/table[1]/tbody/tr[2]/td[2]/b[2]/font').text
            
            if status.upper() == 'ATENDIDA':
                ja_atd += 1
                dmd_ja_atd.append(demanda)
                print(f'{ja_atd}ª Demanda já atendida: {demanda}')
                continue

            navegador.switch_to.default_content()
            sleep(0.5)
            navegador.execute_script(script='''window.top.frames[1].document.getElementById('btnEditarFormBotoes').click();window.top.frames[2].document.getElementById('btnEditar').click()''')

            if direcao == 1:
                cont = 0
                while True:
                    try: 
                        navegador.switch_to.alert.accept()
                        break
                    except Exception as erro:
                        sleep(0.5)
                        cont += 0.5
                        if cont >= 10:
                            erros += 1
                            dmd_erros.append([demanda, erro])
                            print(f'{erros}º erro. Demanda não atendida: {demanda}')
                            break
                if cont >= 10:
                    continue
                
                sleep(0.5)
                iframe_formulario = navegador.find_element(By.XPATH, '//*[@id="form_conteudo"]')

                # Muda o foco para o iframe
                navegador.switch_to.frame(iframe_formulario)

                navegador.find_element(By.XPATH, '/html/body/form/table[3]/tbody/tr[10]/td/textarea').send_keys(msg_atendimento)
                
                x = navegador.find_element(By.XPATH,'/html/body/form/table[3]/tbody/tr[6]/td[2]/font/select')
                Select(x).select_by_value('_j8dnmq82jdtm7b1u6dsg58rrkc5m0_')
                navegador.execute_script(script='''window.top.frames[2].document.getElementById('btnAtender').click()''')

                atd += 1
                print(f'{atd}ª demanda atendida: {demanda}')
                dmd_atd.append(demanda)
                
            else:
                cont = 0
                while True:
                    try: 
                        navegador.switch_to.alert.accept()
                        break
                    except:
                        sleep(0.5)
                        cont += 0.5
                        if cont >= 5:
                            erros += 1
                            dmd_erros.append(demanda)
                            print(f'{erros}º erro. Demanda não atendida: {demanda}')
                            break
                if cont >= 5:
                    continue

                iframe_formulario = navegador.find_element(By.XPATH, '//*[@id="form_conteudo"]')
                navegador.switch_to.frame(iframe_formulario)

                navegador.find_element(By.XPATH, '//*[@id="xSec5_1"]/table[2]/tbody/tr[8]/td/font/textarea').send_keys(msg_redir)
                
                x = navegador.find_element(By.XPATH,'//*[@id="xSec5_1"]/table[3]/tbody/tr/td/font/select')
                Select(x).select_by_value('1')
                navegador.execute_script(script='''javascript:window.top.frames[2].document.getElementById('btnDevolverAoGRAnterior').click()''')

                while True:
                    try: 
                        navegador.switch_to.alert.accept()
                        break
                    except:
                        sleep(0.5)
                        cont += 0.5
                        if cont >= 5:
                            break
                if cont >= 5:
                    continue

                dev += 1
                dmd_dev.append(demanda)
                print(f'{dev}ª demanda devolvida: {demanda}')
                
        except Exception as erro:
            erros += 1
            dmd_erros.append([demanda, erro])
            print(f'{erros}º erro. Demanda não atendida: {demanda}')
    resumo['atd'] = [atd, dmd_atd]
    resumo['dev'] = [dev, dmd_dev]
    resumo['erros'] = [erros, dmd_erros]
    resumo['ja_atd'] = [ja_atd, dmd_ja_atd]
    return resumo

def pegar_resposta():
    while True:
        try:
            resp = int(input('''
De acordo com o relatório de atendimento, você deseja:
[0] - Para sair
[1] - Para atender novamente
            '''))
            if resp in (0, 1):
                break
            else:
                print('Opção inválida! Responda apenas "0" ou "1"')
        except:
            print('Opção inválida! Responda apenas "0" ou "1"')
    return resp



#%%
# ------------- EXECUTANDO TRATATIVAS -------------
relatorio = atende_dmd()


#%%
# ------------- GERANDO RELATORIO E INICIANDO LOOP -------------
tot_dmd = len(tabela)
tot_atd = tabela['ATENDE'].sum()
tot_dev = tot_dmd - tot_atd

while True:
    print(f'''-----------------------------------------
Das demandas de ATENDIMENTO: 
→ {relatorio['atd'][0]}/{tot_atd} foram atendidas agora.
→ {relatorio['ja_atd'][0]}/{tot_atd} já estavam atendidas

RESUMO → {relatorio['atd'][0] + relatorio['ja_atd'][0]}/{tot_atd} estão atendidas.
-----------------------------------------
''')
    print(f'''-----------------------------------------
Das demandas de DEVOLUÇÃO: 
→ {relatorio['dev'][0]}/{tot_dev} foram devolvidas agora.
-----------------------------------------
''')
    print(f'''-----------------------------------------
Das demandas TOTAIS: 
→ {relatorio['erros'][0]}/{tot_dmd} apresentaram possíveis erros.
-----------------------------------------
''')

    resp = pegar_resposta()
    if resp == 0:
        break
    else:
        relatorio = atende_dmd()


#%%
# ------------- FINALIZANDO O PROGRAMA -------------
navegador.quit()
conn_gdnp.close()
