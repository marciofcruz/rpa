# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
# 25/07/2024

# Carregamento de Bibliotecas
from jproperties import Properties
import os
import pathlib
from PyPDF2 import PdfReader
from PyPDF2 import PdfWriter

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

def get_nome_arquivo_tomadores(numero_pagina, texto):
  
  linhas = texto.split('\n')

  vencimento = ''
  cnpj_primeiro = ''
  cnpj_ultimo = ''
  cnpj = ''
  
  cont_linha = 0
  for linha in linhas:
    elementos = linha.split(' ')
    
    if cont_linha==1:
      vencimento = elementos[0]
    elif len(elementos)>6:
      auxiliar = elementos[6]
      auxiliar = auxiliar.upper()
      
      if auxiliar.find('TRABALHADORES')>=0:
        cnpj = auxiliar[-7:].strip()
        
        cnpj = cnpj.replace('-','')
        
        if cnpj.isdigit():
          if cnpj_primeiro == '':
            cnpj_primeiro = cnpj
            cnpj_ultimo = cnpj
          else:
            cnpj_ultimo = cnpj      
    
    cont_linha+=1
    
  nome_arquivo = 'Relação de Tomadores de Serviço '+vencimento.replace('/','-')+' '+cnpj_primeiro+' a '+cnpj_ultimo+ '.pdf'
  return nome_arquivo

def get_nome_arquivo_tipo_de_valor(numero_pagina, texto):
  
  linhas = texto.split('\n')

  vencimento = ''
  cnpj_primeiro = ''
  cnpj_ultimo = ''
  cnpj = ''
  
  cont_linha = 0
  for linha in linhas:
    elementos = linha.split(' ')
    
    if cont_linha==1:
      vencimento = elementos[0]
    else:
      auxiliar = elementos[0]
      auxiliar = auxiliar.upper()
      
      if auxiliar.find('TRABALHADORES')>=0:
        cnpj = auxiliar[-7:].strip()
        
        cnpj = cnpj.replace('-','')
        
        if cnpj.isdigit():
          if cnpj_primeiro == '':
            cnpj_primeiro = cnpj
            cnpj_ultimo = cnpj
          else:
            cnpj_ultimo = cnpj      
    
    cont_linha+=1
    
  nome_arquivo = 'Relação de Tipos de Valor '+vencimento.replace('/','-')
  
  if cnpj_primeiro == '':
    nome_arquivo = nome_arquivo+' página '+str(numero_pagina)+'.pdf'
  else:
    nome_arquivo = nome_arquivo+' '+cnpj_primeiro+' a '+cnpj_ultimo+ '.pdf'
    
  return nome_arquivo

def get_nome_arquivo_categoria(numero_pagina, texto):
  
  linhas = texto.split('\n')

  vencimento = ''
  cnpj_primeiro = ''
  cnpj_ultimo = ''
  cnpj = ''
  
  cont_linha = 0
  for linha in linhas:
    elementos = linha.split(' ')
    
    if cont_linha==1:
      vencimento = elementos[0]
    else:
      auxiliar = elementos[1]
      auxiliar = auxiliar.upper()
      
      if auxiliar.find('TRABALHADORES')>=0:
        cnpj = auxiliar[-7:].strip()
        
        cnpj = cnpj.replace('-','')
        
        if cnpj.isdigit():
          if cnpj_primeiro == '':
            cnpj_primeiro = cnpj
            cnpj_ultimo = cnpj
          else:
            cnpj_ultimo = cnpj
      
    cont_linha+=1
    
  nome_arquivo = 'Relação de Categorias '+vencimento.replace('/','-')+' '+cnpj_primeiro+' a '+cnpj_ultimo+ '.pdf'

  return nome_arquivo

def get_nome_arquivo_estabelecimento(numero_pagina, texto):

  vencimento = ''  
  guia = ''
  linhas = texto.split('\n')
  
  cont_linha = 0
  for linha in linhas:
    elementos = linha.split(' ')
    
    if cont_linha==1:
      vencimento = elementos[0]
      guia = elementos[1]
    else:
      auxiliar = elementos[1]
      auxiliar = auxiliar.upper()
      
    cont_linha+=1
      
  nome_arquivo = 'Relação de Estabelecimentos '+vencimento.replace('/','-')+' GUIA '+guia[0:10]+'.pdf'
  
  return nome_arquivo

def get_tipo_pagina(texto):
  if  texto.upper().find('RELAÇÃO DE CATEGORIAS')>0:
    return 'CATEGORIA'
  elif texto.upper().find('RELAÇÃO DE ESTABELECIMENTOS')>0:
    return 'ESTABELECIMENTO'
  elif texto.upper().find('RELAÇÃO DE TIPOS DE VALOR')>0:
    return 'TIPOSDEVALOR'
  elif texto.upper().find('TOMADORES DE SERVIÇO')>0:
    return 'TOMADORES'
  else:
    return None

def gravar_arquivo(nome_arquivo, reader, numero_pagina):
  writer = PdfWriter()

  try:
    writer.add_page(reader.pages[numero_pagina])
    writer.write(nome_arquivo)
  finally:
    writer = None

def main():
  lista = lista_arquivos_pdf()
  
  if (lista is None) or (len(lista)==0):
    raise Exception("nenhum arquivo PDF de espelho de cartão de ponto encontrado na pasta origem")
  
  for item in lista:
    reader = PdfReader(item)
    
    writer_estabelecimento = PdfWriter()
    
    nome_arquivo_estabelecimento = None
    
    print(f'{item}')
    
    cont_pagina = 0
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      tipo_pagina = get_tipo_pagina(pagina.extract_text())
      
      if tipo_pagina == 'CATEGORIA':
        nome_arquivo = get_nome_arquivo_categoria(cont_pagina, pagina.extract_text())
        gravar_arquivo(os.path.join(diretorio_destino, nome_arquivo), reader, cont_pagina)
      elif tipo_pagina == 'ESTABELECIMENTO':

        if nome_arquivo_estabelecimento is None:
          nome_arquivo_estabelecimento = get_nome_arquivo_estabelecimento(cont_pagina, pagina.extract_text())
          
        writer_estabelecimento.add_page(reader.pages[cont_pagina])
          
      elif tipo_pagina == 'TIPOSDEVALOR':
        nome_arquivo = get_nome_arquivo_tipo_de_valor(cont_pagina, pagina.extract_text())
        gravar_arquivo(os.path.join(diretorio_destino, nome_arquivo), reader, cont_pagina)

      elif tipo_pagina == 'TOMADORES':
        nome_arquivo = get_nome_arquivo_tomadores(cont_pagina, pagina.extract_text())
        gravar_arquivo(os.path.join(diretorio_destino, nome_arquivo), reader, cont_pagina)
        
      cont_pagina += 1
    
    if nome_arquivo_estabelecimento is not None:
      try:
        writer_estabelecimento.write(os.path.join(diretorio_destino, nome_arquivo_estabelecimento))
      except Exception as e:
        print('Erro ao gerar Relação de Estabelecimentos ', e)
        

    reader = None
    writer_estabelecimento = None
    
if __name__ == "__main__":
  main()

