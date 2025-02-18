#%%
#pyinstaller --name="BSC - DOWNLOAD RELATORIO DIARIO" --onefile "BSC - Download relatorio.py" --noconsole
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By
from time import sleep
import os
import calendar
from datetime import date, datetime, timedelta

import warnings
warnings.simplefilter("ignore")
sleep(60*10)


#%%
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {"download.default_directory":
    "K:\\GSAS\\09 - Coordenacao Gestao Numerario\\00 Estudos\\16 KPI´s GSA\\BSC"
    })
driver = webdriver.Chrome()
driver.maximize_window()


# %%
dta_ini = date.today() - timedelta(days=1)
dta_ini = datetime.strftime(dta_ini, "01/%m/%Y")

dta_fim = date.today() - timedelta(days=1)
ult_dia = calendar.monthrange(dta_fim.year, dta_fim.month)[1]
dta_fim = datetime.strftime(dta_fim, f"{ult_dia}/%m/%Y")


#%%
driver.get('http://gaaweb2/MB.Web.UI.GAA.MonitoracaoATM/RelatorioTempoDisponibilidade.aspx?opcao=3,1')
cont = 0
while True:
    if cont > 5:
        print('Erro ao carregar a página!')
        break
    try:
        driver.find_element(By.XPATH, '//*[@id="ctl00_contentPlaceHolder_txtDataInicio"]').send_keys(dta_ini)
        driver.find_element(By.XPATH, '//*[@id="ctl00_contentPlaceHolder_txtDataFim"]').send_keys(dta_fim)
        break
    except:
        driver.refresh()
        sleep(0.5)
        cont += 1

#%%
#clica em pesquisar
while True:
    try:
        driver.execute_script('document.querySelector("#ctl00_contentPlaceHolder_btnPesquisar").click()')
        sleep(60)
        break
    except:
        sleep(1)

#%%
#clica em salvar
while True:
    try:
        sleep(3)
        driver.find_element(By.XPATH, '//*[@id="ctl00_contentPlaceHolder_MBReportViewer1_ctl05_ctl04_ctl00_ButtonLink"]').click()
        sleep(3)
        break
    except:
        sleep(1)

#%%
#clica em CSV
while True:
    try:
        driver.find_element(By.XPATH, '//*[@id="ctl00_contentPlaceHolder_MBReportViewer1_ctl05_ctl04_ctl00_Menu"]/div[2]/a').click()
        sleep(90)
        break
    except:
        sleep(1)


#%%
#Deletando arquivo antigo
try:
    os.remove("K:/GSAS/09 - Coordenacao Gestao Numerario/00 Estudos/16 KPI´s GSA/BSC/MES_ATUAL.csv")
    sleep(1)
except:
    pass


#%%
caminho = "K:/GSAS/09 - Coordenacao Gestao Numerario/00 Estudos/16 KPI´s GSA/BSC/"
lista_arquivos = os.listdir(caminho)

lista_datas = []
for arquivo in lista_arquivos:
    # descobrir a data desse arquivo
    if '.csv' in arquivo[-4:]:
        data = os.path.getmtime(f'{caminho}/{arquivo}')
        lista_datas.append((data, arquivo))
lista_datas.sort(reverse=True)
ultimo_arquivo = lista_datas[0]


#%%
#Renomeando arquivo
os.rename(src=fr"K:/GSAS/09 - Coordenacao Gestao Numerario/00 Estudos/16 KPI´s GSA/BSC/{ultimo_arquivo[1]}"
          ,dst=r"K:/GSAS/09 - Coordenacao Gestao Numerario/00 Estudos/16 KPI´s GSA/BSC/MES_ATUAL.csv")


#%%
driver.close()
