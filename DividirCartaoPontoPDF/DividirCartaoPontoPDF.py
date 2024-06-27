# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
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

def get_campos_cartao(numero_pagina, texto):
    texto = texto.replace('Aguarde. Processando...','')
    
    eh_inicio_cartao = texto.find('ESPELHO DE PONTO')>0
    re = ''
    pis = ''
    nome = ''
    cnpj = ''
    periodo = ''
    
    posicao_apos_re = get_posicao_corte_inicial(texto, 'RE:', True)
    posicao_antes_pis = get_posicao_corte_inicial(texto, 'PIS:', False)
    
    if posicao_apos_re>0 and posicao_antes_pis>0:
      re = str(texto[posicao_apos_re:posicao_antes_pis]).strip()
      
    if posicao_antes_pis>0:
      posicao_apos_pis = get_posicao_corte_inicial(texto, 'PIS:', True)
      pis = str(texto[posicao_apos_pis:posicao_apos_pis+12]).strip()
      
    posicao_depois_nome = get_posicao_corte_inicial(texto, 'Nome:', True)
    posicao_antes_admissao = get_posicao_corte_inicial(texto, 'Admissão:', False)
    if posicao_depois_nome>0 and posicao_antes_admissao>0:
      nome = texto[posicao_depois_nome:posicao_antes_admissao].strip()
      
    posicao_depois_cnpj = get_posicao_corte_inicial(texto, 'C.N.P.J.:', True)
    posicao_antes_site = get_posicao_corte_inicial(texto, 'Site:', False)
    if posicao_depois_cnpj>0 and posicao_antes_site>0:
      cnpj = texto[posicao_depois_cnpj:posicao_antes_site].strip()
      cnpj = cnpj.replace('/','')
      cnpj = cnpj.replace('.','')
      cnpj = cnpj.replace('-','')

    posicao_depois_referencia = get_posicao_corte_inicial(texto, 'Referencia:', True)
    posicao_antes_periodo = get_posicao_corte_inicial(texto, 'Período:', False)
    if posicao_depois_referencia>0 and posicao_antes_periodo>0:
      periodo = texto[posicao_depois_referencia:posicao_antes_periodo].strip()
      periodo = periodo.replace(' de ',' ')

    return eh_inicio_cartao, cnpj, periodo, re, pis, nome


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
    
    ultimo_nome_cartao = ''
    
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      eh_inicio_cartao, cnpj, periodo, re, pis, nome = get_campos_cartao(cont_pagina, pagina.extract_text())
      
      if eh_inicio_cartao==True:
        
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)} RE {re} - {nome}')
        
        if ultimo_nome_cartao!='':
          gravar_cartao(ultimo_nome_cartao, reader, pagina_inicio_documento, cont_pagina)

        nome_arquivo = cnpj+' - '+periodo+' - RE '+re+' - PIS '+pis+' - '+nome
        
        ultimo_nome_cartao = diretorio_destino+'/'+str(nome_arquivo)+'.pdf'  
        
        pagina_inicio_documento = cont_pagina
      else:
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)}')
        
      cont_pagina += 1
    
    if ultimo_nome_cartao!='':
      gravar_cartao(ultimo_nome_cartao, reader, pagina_inicio_documento, len(reader.pages))
      
    reader = None
    
if __name__ == "__main__":
  main()

