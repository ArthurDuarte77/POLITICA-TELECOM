from selenium.webdriver.common.keys import Keys
from collections import defaultdict
import docx
import time
import subprocess
import threading
import subprocess
import os
import time
from tqdm import tqdm
import shutil
import json
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import re

service = Service()
options = webdriver.ChromeOptions()

titulo_arquivo = ""
# options.add_argument("--headless=new")

options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)


driver = webdriver.Chrome(service=service, options=options)
driver.get("https://web.whatsapp.com/")
counter = 0
while True:
    test = driver.find_elements(By.XPATH, '//*[@id="app"]/div/div[2]/div[3]/header/header/div/div[1]/h1')
    if test:
        break
    else:
        counter += 1
        if counter > 100:
            break;
        time.sleep(0.5)
time.sleep(10)
driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[3]/div/div[1]/div/div[2]/button').click()
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[3]/div/div[1]/div/div[2]/div[2]/div/div/p').send_keys("3791332517")
time.sleep(5)
driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[3]/div/div[3]/div[1]/div/div/div[2]/div').click()

def extrair_informacoes_por_loja(caminho_arquivo):
    # Carrega o documento Word
    doc = docx.Document(caminho_arquivo)
    
    # Dicionário para armazenar informações por loja
    lojas = defaultdict(list)
    loja_atual = None
    
    # Itera sobre os parágrafos do documento
    for paragrafo in doc.paragraphs:
        texto = paragrafo.text.strip()
        if texto:
            # Verifica se o texto é um nome de loja
            if texto.startswith("*") and texto.endswith("*"):
                loja_atual = texto.strip("*")
            elif loja_atual:
                # Adiciona o texto à lista da loja atual
                lojas[loja_atual].append(texto)
    
    return lojas

def enviar():
    caminho_arquivo = 'dados_extraidos.docx'  # Substitua pelo caminho do seu arquivo Word
    lojas = extrair_informacoes_por_loja(caminho_arquivo)
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys("--- POLITICA-JFA ---")
    time.sleep(2)
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
    for loja, detalhes in lojas.items():
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(f"*{loja}*")
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.CONTROL, Keys.RETURN)
        for detalhe in detalhes:
            time.sleep(2)
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(detalhe)
            time.sleep(1)
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.CONTROL, Keys.RETURN)
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
def executar_codigo():
    subprocess.run(['python', 'run_all.py'])


while True:
    executar_codigo()
    enviar()


    time.sleep(900)
