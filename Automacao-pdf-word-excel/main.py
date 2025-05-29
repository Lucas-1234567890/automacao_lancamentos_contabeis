from navegador import iniciar_navegador, login, acessar_pagina_patrimonio, fechar_navegador
from processador import processar_pasta
from extrator import inserir_valores_de_documento_pdf, inserir_valores_de_documento_excel, inserir_valores_de_documento_word
from config import *

navegador = iniciar_navegador()
login(navegador)
acessar_pagina_patrimonio(navegador)

processar_pasta(navegador, PASTA_PDF, '.pdf', inserir_valores_de_documento_pdf)
processar_pasta(navegador, PASTA_EXCEL, '.xlsx', inserir_valores_de_documento_excel)
processar_pasta(navegador, PASTA_WORD, '.docx', inserir_valores_de_documento_word)

fechar_navegador(navegador)
