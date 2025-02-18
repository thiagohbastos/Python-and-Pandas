from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {'protocol_handler.excluded_schemes.tel': False})

import os

import time
import shutil

def login_paytrack( username, password ):

    # Página de login Paytrack
    driver.get("https://login.paytrack.com.br/")

    # Digitar usuário
    wait.until(EC.visibility_of_element_located((By.ID, "normal_login_username"))).send_keys(username)
    # Botão próximo
    driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/div/form/div[4]/button").click()

    # Digitar senha
    wait.until(EC.visibility_of_element_located((By.ID, "normal_login_password"))).send_keys(password)
    # Botão Login
    driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/div/form/div[4]/div/button").click()
    #time.sleep(60)

    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    time.sleep(10)

    # Popup
    try:
        driver.find_element(By.XPATH, "/html/body/div[6]/div/div/div[2]/div[1]/span").click()
        time.sleep(5)
    except:
        pass

def baixar_relatorio( 
    relatorio,
    data_inicial, 
    data_final, 
    id_data_inicial, 
    id_data_final 
):
    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(5)
    
    # Pesquisar relatório    
    wait.until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ))).send_keys(Keys.CONTROL + "a")
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(Keys.DELETE)
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(relatorio)
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/button'
                    ).click()

    # Acessar relatório
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[6]/ul/li/div/div/p'
                    ).click()
    time.sleep(2)

    # Inserir Parâmetros
    wait.until(EC.visibility_of_element_located((By.ID, id_data_inicial
                    ))).send_keys(data_inicial)
    driver.find_element(By.ID, id_data_final
                    ).send_keys(data_final)
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[1]/select'
                    ).send_keys("EXCEL (XLSX)")

    # Download
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[2]/button[1]'
                    ).click()
    
def baixar_relatorio_criterio( 
    relatorio,
    data_inicial, 
    data_final, 
    id_data_inicial, 
    id_data_final 
):
    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(5)

    # Pesquisar relatório    
    wait.until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ))).send_keys(Keys.CONTROL + "a")
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(Keys.DELETE)
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(relatorio)
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/button'
                    ).click()

    # Acessar relatório
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[6]/ul/li/div/div/p'
                    ).click()
    time.sleep(2)

    # Inserir Parâmetros
    wait.until(EC.visibility_of_element_located((By.ID, id_data_inicial
                    ))).send_keys(data_inicial)
    driver.find_element(By.ID, id_data_final
                    ).send_keys(data_final)
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[1]/table/tbody/tr/td[1]/div/div[4]/div[2]/div[1]/input'
                    ).send_keys("S")
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[1]/select'
                    ).send_keys("EXCEL (XLSX)")

    # Download
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[2]/button[1]'
                    ).click()

def baixar_relatorio_duas_datas( 
    relatorio,
    data_inicial, 
    data_final, 
    id_data_inicial, 
    id_data_final,
    id_data_inicial_partida, 
    id_data_final_partida
):
    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(5)
    
    # Pesquisar relatório    
    wait.until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ))).send_keys(Keys.CONTROL + "a")
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(Keys.DELETE)
    driver.find_element(By.XPATH, f'//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/input'
                    ).send_keys(relatorio)
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/button'
                    ).click()

    # Acessar relatório
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[2]/div/div[2]/div[6]/ul/li/div/div/p'
                    ).click()
    time.sleep(2)

    # Inserir Parâmetros
    wait.until(EC.visibility_of_element_located((By.ID, id_data_inicial
                    ))).send_keys(data_inicial)
    driver.find_element(By.ID, id_data_final
                    ).send_keys(data_final)
    driver.find_element(By.ID, id_data_inicial_partida
                    ).send_keys(data_inicial)
    driver.find_element(By.ID, id_data_final_partida
                    ).send_keys(data_final)
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[1]/select'
                    ).send_keys("EXCEL (XLSX)")

    # Download
    driver.find_element(By.XPATH, '//*[@id="div_relatorio"]/div/div[1]/div[2]/div[1]/div[2]/div[2]/button[1]'
                    ).click()

def buscar_historico_download(seconds):
    # Carregar página de histórico
    try:
        wrapper = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="reportList"]/relatorio-async-react/div/div/div[2]/div/div/div/div/div/div/div/div/table'
                                            ))
        )
    except TimeoutException:
        print("Element did not show up")
        
    # Página de relatórios
    driver.get("https://app.paytrack.com.br/#/relatorios-gerenciais")
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(seconds)

    # Acessar relatório
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[1]/div/div/button[2]'
                        ).click()
    try:
        wrapper = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[3]/div/div[2]/div[2]/div/div/div/div/div/div/div/div/table/tbody/tr[1]/td[5]/button'
                                        ))
        )
    except TimeoutException:
        print("Element did not show up")
    
    # Download
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[3]/div/div[2]/div[2]/div/div/div/div/div/div/div/div/table/tbody/tr[1]/td[5]/button'
                        ).click()

    # Voltar para página de relatórios
    driver.find_element(By.XPATH, '//*[@id="body_menu_novo"]/div/div/visao-geral-relatorios-gerenciais/div/div/div[1]/div/div/button[1]'
                        ).click()
    
def mover_para_pasta( relatorio, download_path, destination_path ):
    # Esperar download
    while not os.path.exists(download_path):
        time.sleep(2)
    if os.path.isfile(download_path):
        print(f"{relatorio} downloaded")

    # Mover para pasta
    shutil.move(download_path, destination_path)
    print(f"{relatorio} moved")
