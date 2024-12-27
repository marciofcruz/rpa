# Script desenvolvido 
# 27/06/2024

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

def get_campos_ordem(texto):
    texto = texto.replace('Aguarde. Processando...','')
    
    eh_ordem_valida = texto.find('Instrução')>0
    re = ''
    nome = ''
    
    posicao_apos_re = get_posicao_corte_inicial(texto, 'Re:', True)
    posicao_antes_nome = get_posicao_corte_inicial(texto, 'Nome Compledo do(a) empregrado(a):', True)
    
    if posicao_apos_re>0:
      re = str(texto[posicao_apos_re:posicao_apos_re+8]).strip()
      
    if posicao_antes_nome>0:
      auxiliar = str(texto[posicao_antes_nome:posicao_antes_nome+50]).strip()
      posicao_corte = get_posicao_corte_inicial(auxiliar, "\n", False )
      nome = str(auxiliar[0:posicao_corte]).strip()
    
    return eh_ordem_valida, re, nome


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
  
  for item in lista:
    reader = PdfReader(item)
    
    pagina_inicio_documento = 0
    
    ultimo_nome = ''
    
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      eh_ordem_valida, re, nome = get_campos_ordem(pagina.extract_text())
      
      if eh_ordem_valida==True:
        
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)} RE {re} - {nome}')
        
        nome_arquivo = nome+' - RE '+re
        
        if ultimo_nome!='' and ultimo_nome!=nome_arquivo:
          gravar_cartao(diretorio_destino+'/'+str(ultimo_nome)+'.pdf', reader, pagina_inicio_documento, cont_pagina)
          print('Gravando: ', diretorio_destino+'/'+str(ultimo_nome)+'.pdf')
          pagina_inicio_documento = cont_pagina

        ultimo_nome = nome_arquivo
      else:
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)}')
        
      cont_pagina += 1
    
    if nome_arquivo!='':
      gravar_cartao(diretorio_destino+'/'+str(ultimo_nome)+'.pdf', reader, pagina_inicio_documento, len(reader.pages))
      print('Gravando: ', diretorio_destino+'/'+str(ultimo_nome)+'.pdf')
      
    reader = None
    
if __name__ == "__main__":
  main()

