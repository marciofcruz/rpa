# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
# 17/07/2024

# Carregamento de Bibliotecas
from jproperties import Properties
import os
import pathlib 
import pandas as pd
from pathlib import Path
from PyPDF2 import PdfReader
from PyPDF2 import PdfWriter
import logging

logger = logging.getLogger("PyPDF2")
logger.setLevel(logging.ERROR)

# Configurações Gerais
configs = Properties()

with open('app-config.properties', 'rb') as config_file:
  configs.load(config_file)   
  
diretorio_origem = str(configs.get("diretorio_origem").data)
diretorio_destino = str(configs.get("diretorio_destino").data)   

def lista_arquivos_pdf():
  desktop = pathlib.Path(diretorio_origem)
  return list(desktop.rglob("*.PDF"))

def get_posicao_corte_inicial(textooriginal, token, usa_sufixo):
  posicao = textooriginal.find(token)
  corte_inicial = posicao
  
  if posicao>0 and usa_sufixo:
    corte_inicial = posicao+len(str(token))
    
  return corte_inicial

def get_campos_arquivo(texto):
  
  cnpj = ''
  competencia = ''
  data_emissao = ''
  
  posicao_atento = get_posicao_corte_inicial(texto, 'ATENTO BRASIL SA', True)
  
  if posicao_atento>0:
    auxiliar = texto[posicao_atento:posicao_atento+57].strip()
    auxiliar = auxiliar[0:18]
    cnpj = auxiliar.strip()
    cnpj = cnpj.replace('.','').replace('/','').replace('-','').replace('\n','')
  else:
    posicao_atento = get_posicao_corte_inicial(texto, 'ATENTO DO BRASIL SA', True)
    
    if posicao_atento>0:
      auxiliar = texto[posicao_atento:posicao_atento+57].strip()
      auxiliar = auxiliar[0:18]
      cnpj = auxiliar.strip()
      cnpj = cnpj.replace('.','').replace('/','').replace('-','').replace('\n','')
       

  posicao_competencia = get_posicao_corte_inicial(texto, 'COMP:', True)
  if posicao_competencia>0:
    competencia = texto[posicao_competencia:posicao_competencia+9]
    competencia = competencia.strip()
    competencia = competencia.replace('/','')
    
  posicao_pagina = get_posicao_corte_inicial(texto, 'PÁG :', True)    
  if posicao_pagina>0:
    data_emissao = texto[posicao_pagina:posicao_pagina+10]
    data_emissao = data_emissao.strip()
    data_emissao = data_emissao.replace('/','-')
    
  return cnpj, competencia, data_emissao

def gravar_cartao(ultimo_nome_cartao, reader, pagina_inicio_documento, cont_pagina):
  writer = PdfWriter()
  
  cont = pagina_inicio_documento
  try:
    while cont<(cont_pagina):
      writer.add_page(reader.pages[cont])
      cont += 1 
      
    writer.write(ultimo_nome_cartao)
  finally:
    writer = None
    
def main():
  lista = lista_arquivos_pdf()

  if (lista is None) or (len(lista)==0):
    raise Exception("nenhum arquivo PDF de espelho de cartão de ponto encontrado na pasta origem")

  posicao = 0
  
  for item in lista:
    reader = PdfReader(item, strict=False)

    posicao+=1
    
    nome_arquivo = Path(item).stem
    
    linha = f'-{nome_arquivo}'
    print(f'{posicao}/{len(lista)} - {linha}')
    
    if len(reader.pages)>0:
      try:
        pagina = reader.pages[0]
        
        cnpj, competencia, data_emissao = get_campos_arquivo(pagina.extract_text())
        
        nome_arquivo = f'/RE_{cnpj} - {competencia} - {data_emissao}'# - {hora_emissao}'
        ultimo_nome_cartao = diretorio_destino+'/'+str(nome_arquivo)+'.pdf'  
        
        gravar_cartao(ultimo_nome_cartao, reader, len(reader.pages)-3, len(reader.pages))
        
      except Exception as e:
        print('Erro:', linha, e)
        
if __name__ == "__main__":
  main()
