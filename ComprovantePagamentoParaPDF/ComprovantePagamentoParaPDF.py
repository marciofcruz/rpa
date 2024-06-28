# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
# 28/06/2024

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

def get_campos_super_digital(numero_pagina, texto):
  eh_inicio_comprovante = texto.find('COMPROVANTE DE PAGAMENTO')>=0
  documento = ''
  data_pagamento = ''
  valor = ''
  banco = ''
  nome = ''

  posicao_apos_autenticacao_legis = get_posicao_corte_inicial(texto, 'Autentic. Legis', True)
    
  if posicao_apos_autenticacao_legis>0:
    linhas = texto[posicao_apos_autenticacao_legis:].split('\n')
      
    documento = linhas[6]
    banco = linhas[1]
    nome = linhas[5]
    data_pagamento = str(linhas[12] ).replace('/','-')
    valor= linhas[15]
    
  return eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento

def get_campos_itau(numero_pagina, texto):
  eh_inicio_comprovante = texto.find('Transferência efetuada em')>=0
  
  documento = ''
  data_pagamento = ''
  valor = ''
  banco = ''
  nome = ''
  
  if eh_inicio_comprovante:
    linhas = texto.split('\n')

    valor = linhas[1]

    banco = 'Itau'
    documento = linhas[0].replace('  ','').replace(' - ','-')
    data_pagamento = linhas[12][38:48].replace('/','-')
    
    posicao_apos_ponto = get_posicao_corte_inicial(linhas[12], '.', True)
    posicao_antes_sispag = get_posicao_corte_inicial(linhas[12], 'SISPAG', False)
    nome = linhas[12][posicao_apos_ponto:posicao_antes_sispag]
  
  return eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento

def get_campos_bradesco(numero_pagina, texto):
  eh_inicio_comprovante = texto.find('COMPROVANTE DE CRÉDITO EM CONTA')>=0
  
  documento = ''
  data_pagamento = ''
  valor = ''
  banco = ''
  nome = ''
  
  if eh_inicio_comprovante:
    linhas = texto.split('\n')

    # cont = 0
    # for linha in linhas:
    #   print(cont, ':', linha)
    #   cont+=1
    
    valor = linhas[17]
    banco = 'BRADESCO'
    nome = linhas[14]

    posicao_apos_espaco = get_posicao_corte_inicial(linhas[15], ' ', True)
    if posicao_apos_espaco>=0:
      documento = linhas[15][0:posicao_apos_espaco].strip()
      data_pagamento = linhas[15][posicao_apos_espaco:].strip().replace('/', '-')

  return eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento

  

def get_campos_comprovante(numero_pagina, texto):
  eh_inicio_comprovante = False
  documento = ''
  data_pagamento = ''
  valor = ''
  banco = '.'
  nome = '.'
    
  if texto.upper().find('997 - SUPER DIGITAL')>0:
    eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_super_digital(numero_pagina, texto)
  elif texto.upper().find('033 - SANTANDER')>0:
    eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_super_digital(numero_pagina, texto)
  elif texto.upper().find('BRADESCO')>0:
    eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_bradesco(numero_pagina, texto)
  else:
    eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_itau(numero_pagina, texto)
    
  return eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento

def gravar_comprovante(ultimo_nome_cartao, reader, pagina_inicio_documento, cont_pagina):
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
    
    ultimo_nome_comprovante = ''
    
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_comprovante(cont_pagina, pagina.extract_text())
      
      if eh_inicio_comprovante==True:
        
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)} documento {documento} - {nome}')
        
        if ultimo_nome_comprovante!='':
          gravar_comprovante(ultimo_nome_comprovante, reader, pagina_inicio_documento, cont_pagina)

        nome_arquivo = documento+' - '+nome+' - '+valor+' - '+data_pagamento+' - '+banco
        
        ultimo_nome_comprovante = diretorio_destino+'/'+str(nome_arquivo)+'.pdf'  
        
        pagina_inicio_documento = cont_pagina
      else:
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)}')
        
      cont_pagina += 1
      
    if ultimo_nome_comprovante!='':
      gravar_comprovante(ultimo_nome_comprovante, reader, pagina_inicio_documento, cont_pagina)
      
    reader = None
    
if __name__ == "__main__":
  main()

