# -*- coding: utf-8 -*-

#%%  Bibliotecas

import cv2
import pyautogui
import numpy as np
from pathlib import Path
import time
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
import chromedriver_autoinstaller
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import ui
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import urllib.request

#%% Janela do Programa

# Cria a janela principal
import tkinter as tk

# Cria a janela
janela = tk.Tk()
janela.title("Busca Arq Cota Semanal")

# Cria os elementos da interface gráfica
texto_orientacao = tk.Label(janela, text="MOB - Cota Semanal")
texto_orientacao.grid(column=1, row=1)


# Cria a checkbox para verificar mes atual
checkbox_var = tk.IntVar()
checkbox = tk.Checkbutton(janela, text="Cota para o mês atual?", variable=checkbox_var)
checkbox.grid(column=1, row=3)

def buttonClicked():
    global mes_atual
    mes_atual = checkbox_var.get()
    if mes_atual == 0:
        texto_alerta1 = tk.Label(janela, text="Baixe os dados de orçamento colaborador diretamente do SGC")
        texto_alerta1.grid(column=1, row=5)
        
    time.sleep(5)
    janela.destroy()

botaoFinal = tk.Button(janela, text="Rodar Código", command=buttonClicked)
botaoFinal.grid(column=1, row=4)

# Executa o loop principal da janela
janela.mainloop()


#%%  Caminho para salva os arquivos

home = str(Path.home())
caminho_arquivos_salvar = home + ''

# Define o caminho da pasta de downloads
downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

#%% Inicialização do Google

chromedriver_autoinstaller.install()

# Create a webdriver instance for the Chrome browser
driver = webdriver.Chrome()
driver.maximize_window()

# Navigate to the website
driver.get('https://combustivel.thinc.com.br/web')


#%% Login na mob.adm

username_field = driver.find_element(By.ID, 'name')
password_field = driver.find_element(By.ID, 'pass')
username_field.send_keys('')
password_field.send_keys('')

submit_button = driver.find_element(By.ID, 'press')
submit_button.click()

#%% Esperar a prox pag carregar para executar o prox comando
def employeersReport():

# Entrar na Aba Funcionarios
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "app-header")))
    
    entrar_funcionarios = driver.find_element_by_xpath('//*[@id="_t_view_4"]/div[1]/div[2]/a[4]')
    entrar_funcionarios.click()

## Clica para exportar em excel e muda o nome do arquivo
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "_t_view_7")))

    baixar_excel_func = driver.find_element_by_xpath('//*[@id="excel-export"]')
    baixar_excel_func.click()

    time.sleep(20)
    
# Obtém a lista de arquivos na pasta de downloads
    files = os.listdir(downloads_folder)

# Ordena a lista de arquivos pelo horário de modificação
# para obter o último arquivo salvo
    files.sort(key=lambda x: os.path.getmtime(os.path.join(downloads_folder, x)))

# Obtém o nome do último arquivo salvo
    last_saved_file = files[-1]

# Define o caminho da pasta de destino
    destination_folder = caminho_arquivos_salvar

# Define o novo nome do arquivo
    new_file_name = 'base_employees_report.xlsx'

# Monta o caminho completo do arquivo de origem
    source_path = os.path.join(downloads_folder, last_saved_file)

# Monta o caminho completo do arquivo de destino
    destination_path = os.path.join(destination_folder, new_file_name)

# Copia o arquivo de origem para o destino e o renomeia
    shutil.move(source_path, destination_path)

    print('Download base_employees_report feito')

employeersReport()

#%% Esperar a prox pag carregar para executar o prox comando
def vehiclesReport():

# Entrar na Aba Veículos
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "dropdownMenu1")))
    
    entrar_veic1 = driver.find_element_by_xpath('//*[@id="dropdownMenu1"]')
    entrar_veic1.click()
    
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "_t_view_8")))
    
    entrar_veic2 = driver.find_element_by_xpath('//*[@id="_t_view_8"]/div[1]/div[2]/div/ul/li[2]/a')
    entrar_veic2.click()

## Clica para exportar em excel e muda o nome do arquivo
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "_t_view_11")))

    baixar_excel_veic = driver.find_element_by_xpath('//*[@id="excel-export"]')
    baixar_excel_veic.click()
    
    time.sleep(20)

# Obtém a lista de arquivos na pasta de downloads
    files = os.listdir(downloads_folder)

# Ordena a lista de arquivos pelo horário de modificação
# para obter o último arquivo salvo
    files.sort(key=lambda x: os.path.getmtime(os.path.join(downloads_folder, x)))

# Obtém o nome do último arquivo salvo
    last_saved_file = files[-1]

# Define o caminho da pasta de destino
    destination_folder = caminho_arquivos_salvar

# Define o novo nome do arquivo
    new_file_name = 'base_vehicles_report.xlsx'

# Monta o caminho completo do arquivo de origem
    source_path = os.path.join(downloads_folder, last_saved_file)

# Monta o caminho completo do arquivo de destino
    destination_path = os.path.join(destination_folder, new_file_name)

# Copia o arquivo de origem para o destino e o renomeia
    shutil.move(source_path, destination_path)

    print('Download base_vehicles_report feito')
    
    
vehiclesReport()

#%% Esperar a prox pag carregar para executar o prox comando

def baseOrcamentoSGC():

# Sai do Login e entra em outro para exportar o orçamento
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "_t_view_12")))
    
    sai_login = driver.find_element_by_xpath('//*[@id="_t_view_12"]/div[2]/div[2]')
    sai_login.click()
    
#  Login na doumit.mob

    wait = WebDriverWait(driver, 30)
    wait.until(EC.visibility_of_element_located((By.ID, "login-panel")))

    username_field = driver.find_element(By.ID, 'name')
    password_field = driver.find_element(By.ID, 'pass')
    username_field.send_keys('')
    password_field.send_keys('')

# Find the submit button and click on it
    submit_button = driver.find_element(By.ID, 'press')
    submit_button.click()
    
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "signinf")))
    
    verificacao_login = driver.find_element_by_xpath('//*[@id="signinf"]/div[1]/a')
    verificacao_login.click()

# Entrar na Aba Orçamento
    

    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "_t_view_12")))
    
    entrar_orc1 = driver.find_element_by_xpath('//*[@id="_t_view_12"]/div[1]/div[2]/a[4]')
    entrar_orc1.click()
    
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, "_t_view_13")))
    
    entrar_orc2 = driver.find_element_by_xpath('//*[@id="_t_view_13"]/ul/li[2]/a')
    entrar_orc2.click()

## Clica para exportar em excel e muda o nome do arquivo
    wait = WebDriverWait(driver, 120)
    wait.until(EC.visibility_of_element_located((By.ID, "_t_view_18")))

    baixar_excel_orc_sgc = driver.find_element_by_xpath('//*[@id="excel-export"]')
    baixar_excel_orc_sgc.click()
    
    time.sleep(20)

# Obtém a lista de arquivos na pasta de downloads
    files = os.listdir(downloads_folder)

# Ordena a lista de arquivos pelo horário de modificação
# para obter o último arquivo salvo
    files.sort(key=lambda x: os.path.getmtime(os.path.join(downloads_folder, x)))

# Obtém o nome do último arquivo salvo
    last_saved_file = files[-1]

# Define o caminho da pasta de destino
    destination_folder = caminho_arquivos_salvar

# Define o novo nome do arquivo
    new_file_name = 'base_orcamento_sgc.xlsx'

# Monta o caminho completo do arquivo de origem
    source_path = os.path.join(downloads_folder, last_saved_file)

# Monta o caminho completo do arquivo de destino
    destination_path = os.path.join(destination_folder, new_file_name)
    
# Copia o arquivo de origem para o destino e o renomeia
    shutil.move(source_path, destination_path)

    print('Download base_orcamento_sgc feito')
    
if mes_atual == 0:
    print('cota para prox mes')
else:
    baseOrcamentoSGC()

#%% TicketLog

        

#%%
# Close the browser
driver.quit()

#%%JANELA DE FINALIZAÇÃO
janela_alerta = tk.Tk()
janela_alerta.title("Resultados")

# Cria os elementos da interface gráfica
texto_orientacao_final = tk.Label(janela_alerta, text="Arquivos Salvos na Pasta Arq_Cota_Python em ")
texto_orientacao_final.grid(column=1, row=1)

if mes_atual == 0:
    texto_orientacao_final2 = tk.Label(janela_alerta, text="Baixe o orçamento dos colaboradores do proximo mes diretamente do SGC")
    texto_orientacao_final2.grid(column=1, row=2)
    
janela_alerta.mainloop()

    