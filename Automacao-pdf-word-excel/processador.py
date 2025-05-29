from navegador import preencher_id, clicar_xpath
from time import sleep
import os

def processar_pasta(navegador, caminho_pasta, extensao, func_extrair_dados):
    contador = 0
    for nome_arquivo in os.listdir(caminho_pasta):
        if nome_arquivo.endswith(extensao):
            contador += 1
            print(f"[processar_pasta] Arquivo {contador}: {nome_arquivo}")
            caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)
            dados = func_extrair_dados(caminho_arquivo)

            for id_campo, valor in dados.items():
                preencher_id(navegador, id_campo, valor)
            
            clicar_xpath(navegador, '//*[@id="balanco-form"]/div[2]/button')
            sleep(1)

    print(f"Total de arquivos {extensao} processados na pasta {caminho_pasta}: {contador}")
