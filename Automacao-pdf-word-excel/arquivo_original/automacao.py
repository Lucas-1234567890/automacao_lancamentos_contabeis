# 🚀 ===== IMPORTS =====
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from time import sleep
from docx import Document
import os
import pandas as pd
import pdfplumber
import re


# 🔧 ===== FUNÇÕES DE SUPORTE =====
def preencher_id(id_campo, valor):
    campo = navegador.find_element(By.ID, id_campo)
    sleep(0.3)
    campo.click()
    campo.send_keys(valor)


def preencher_xpath(xpath, valor):
    campo = navegador.find_element(By.XPATH, xpath)
    sleep(0.5)
    campo.click()
    campo.send_keys(valor)


def clicar_xpath(xpath):
    botao = navegador.find_element(By.XPATH, xpath)
    sleep(0.5)
    botao.click()


def processar_pasta(caminho_pasta, extensao, func_extrair_dados):
    contador = 0
    for nome_arquivo in os.listdir(caminho_pasta):
        if nome_arquivo.endswith(extensao):
            contador += 1
            print(f"📄 Processando {nome_arquivo}")
            caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)
            dados = func_extrair_dados(caminho_arquivo)

            for id_campo, valor in dados.items():
                preencher_id(id_campo, valor)

            clicar_xpath('//*[@id="balanco-form"]/div[2]/button')
            sleep(1)
    print(f"✅ Total de arquivos .{extensao} processados na pasta: {contador}")


# 🌐 ===== CONFIGURAÇÃO DO NAVEGADOR =====
servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument(r'user-data-dir=C:\Selenium\ProfileSelenium')

navegador = webdriver.Chrome(service=servico, options=options)
navegador.get("https://contabil-devaprender.netlify.app/")
navegador.maximize_window()

# 🕒 Aguardar carregamento da página inicial
while len(navegador.find_elements(By.XPATH, '/html/body/div')) < 1:
    sleep(1)
sleep(1)

# 🔑 ===== LOGIN =====
preencher_xpath('/html/body/div/div/form/div[1]/input', 'lucas@hotmail.com')
preencher_xpath('/html/body/div/div/form/div[2]/input', 'meunomeelucaseessasenhaedifcil')
clicar_xpath('/html/body/div/div/form/button')

# ⏳ Aguardar redirecionamento e acessar balanço patrimonial
sleep(5)
clicar_xpath('/html/body/div/div/div/div[1]/div/div/a')


# 📄 ===== LEITOR DE DOCUMENTO WORD =====
def inserir_valores_de_documento_word(caminho_arquivo_word):
    arquivo_word = Document(caminho_arquivo_word)

    dados = {
        'ativo_circulante': '', 'caixa_equivalentes': '', 'contas_receber': '',
        'estoques': '', 'ativo_nao_circulante': '', 'imobilizado': '',
        'intangivel': '', 'total_ativo': ''
    }

    for tabela in arquivo_word.tables:
        for linha in tabela.rows:
            texto = linha.cells[0].text.strip()
            valor = linha.cells[1].text.strip()

            if 'Ativo Circulante' in texto:
                dados['ativo_circulante'] = valor
            elif '- Caixa e Equivalentes' in texto:
                dados['caixa_equivalentes'] = valor
            elif '- Contas a Receber' in texto:
                dados['contas_receber'] = valor
            elif '- Estoques' in texto:
                dados['estoques'] = valor
            elif 'Ativo Não Circulante' in texto:
                dados['ativo_nao_circulante'] = valor
            elif '- Imobilizado' in texto:
                dados['imobilizado'] = valor
            elif '- Intangível' in texto:
                dados['intangivel'] = valor
            elif 'Total do Ativo' in texto:
                dados['total_ativo'] = valor

    return dados


# 📊 ===== LEITOR DE DOCUMENTO EXCEL =====
def inserir_valores_de_documento_excel(caminho_excel):
    df = pd.read_excel(caminho_excel)

    dados = {
        'ativo_circulante': '', 'caixa_equivalentes': '', 'contas_receber': '',
        'estoques': '', 'ativo_nao_circulante': '', 'imobilizado': '',
        'intangivel': '', 'total_ativo': ''
    }

    for _, row in df.iterrows():
        texto = str(row['Ativo']).strip().lower()
        valor_raw = str(row['Valor (R$)']).strip()

        valor_tratado = valor_raw.replace('.', '').replace(',', '.')

        try:
            valor_float = float(valor_tratado)
            valor = str(valor_float)
        except:
            valor = valor_raw

        if texto == 'ativo circulante':
            dados['ativo_circulante'] = valor
        elif 'caixa' in texto:
            dados['caixa_equivalentes'] = valor
        elif 'contas a receber' in texto:
            dados['contas_receber'] = valor
        elif 'estoques' in texto:
            dados['estoques'] = valor
        elif texto == 'ativo não circulante':
            dados['ativo_nao_circulante'] = valor
        elif 'imobilizado' in texto:
            dados['imobilizado'] = valor
        elif 'intangível' in texto:
            dados['intangivel'] = valor
        elif 'total do ativo' in texto:
            dados['total_ativo'] = valor

    return dados


# 📑 ===== LEITOR DE DOCUMENTO PDF =====
def inserir_valores_de_documento_pdf(caminho_pdf):
    dados = {
        'ativo_circulante': '', 'caixa_equivalentes': '', 'contas_receber': '',
        'estoques': '', 'ativo_nao_circulante': '', 'imobilizado': '',
        'intangivel': '', 'total_ativo': ''
    }

    texto_completo = ""

    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + '\n'

    linhas = texto_completo.split('\n')

    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue

        resultado = re.match(r'^(.*?)([\d\.,]+)$', linha)
        if not resultado:
            continue

        texto = resultado.group(1).strip()
        valor = resultado.group(2).strip()

        if 'Ativo Circulante' in texto:
            dados['ativo_circulante'] = valor
        elif 'Caixa' in texto:
            dados['caixa_equivalentes'] = valor
        elif 'Contas a Receber' in texto:
            dados['contas_receber'] = valor
        elif 'Estoques' in texto:
            dados['estoques'] = valor
        elif 'Ativo Não Circulante' in texto:
            dados['ativo_nao_circulante'] = valor
        elif 'Imobilizado' in texto:
            dados['imobilizado'] = valor
        elif 'Intangível' in texto:
            dados['intangivel'] = valor
        elif 'Total do Ativo' in texto:
            dados['total_ativo'] = valor

    return dados


# 🗂️ ===== PROCESSAMENTO DOS DOCUMENTOS =====
pasta_excel = r'C:\Users\Lucas\OneDrive\Arquivos diversos\Desktop\Automacao-pdf-word-excel\pasta_excel'
pasta_pdf = r'C:\Users\Lucas\OneDrive\Arquivos diversos\Desktop\Automacao-pdf-word-excel\pasta_pdf'
pasta_word = r'C:\Users\Lucas\OneDrive\Arquivos diversos\Desktop\Automacao-pdf-word-excel\pasta_word'

processar_pasta(pasta_pdf, '.pdf', inserir_valores_de_documento_pdf)
processar_pasta(pasta_excel, '.xlsx', inserir_valores_de_documento_excel)
processar_pasta(pasta_word, '.docx', inserir_valores_de_documento_word)

# 🏁 ===== FINALIZAR =====
navegador.quit()
