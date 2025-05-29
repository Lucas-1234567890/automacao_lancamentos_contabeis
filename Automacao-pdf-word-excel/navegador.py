from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from time import sleep
from config import URL_SITE

def iniciar_navegador():
    servico = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument(r'user-data-dir=C:\Selenium\ProfileSelenium')
    navegador = webdriver.Chrome(service=servico, options=options)
    navegador.get(URL_SITE)
    navegador.maximize_window()
    return navegador

def login(navegador):
    esperar_xpath(navegador, '/html/body/div/div/form/div[1]/input')
    preencher_xpath(navegador, '/html/body/div/div/form/div[1]/input', 'lucas@hotmail.com')
    preencher_xpath(navegador, '/html/body/div/div/form/div[2]/input', 'senha')
    clicar_xpath(navegador, '/html/body/div/div/form/button')

def acessar_pagina_patrimonio(navegador):
    sleep(5)
    clicar_xpath(navegador, '/html/body/div/div/div/div[1]/div/div/a')

def esperar_xpath(navegador, xpath, timeout=30):
    for _ in range(timeout):
        if len(navegador.find_elements(By.XPATH, xpath)) > 0:
            return True
        sleep(1)
    raise Exception(f"Elemento {xpath} não encontrado")

def preencher_xpath(navegador, xpath, valor):
    campo = navegador.find_element(By.XPATH, xpath)
    sleep(0.3)
    campo.click()
    campo.send_keys(valor)

def preencher_id(navegador, id_campo, valor):
    campo = navegador.find_element(By.ID, id_campo)
    sleep(0.3)
    campo.click()
    campo.send_keys(valor)

def clicar_xpath(navegador, xpath):
    botao = navegador.find_element(By.XPATH, xpath)
    sleep(0.3)
    botao.click()

def fechar_navegador(navegador):
    navegador.quit()
