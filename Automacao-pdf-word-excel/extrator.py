import pandas as pd
import pdfplumber
import re
from docx import Document

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
