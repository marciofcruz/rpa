# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
# 28/06/2024

# Carregamento de Bibliotecas
from jproperties import Properties
import os
import os.path
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
  eh_inicio_comprovante = texto.find('COMPROVANTE DE PAGAMENTO')>=0 or texto.find('COMPROVANTE DE TRANSFERENCIA')>=0 

  documento = ''
  data_pagamento = ''
  valor = ''
  banco = ''
  nome = ''

  posicao_apos_autenticacao_legis = get_posicao_corte_inicial(texto, 'Autentic. Legis', True)
    
  if posicao_apos_autenticacao_legis>0:
    linhas = texto[posicao_apos_autenticacao_legis:].split('\n')
    
    # cont = 0
    # for linha in linhas:
    #   print(cont,'-', linha)
    #   cont+=1
    
    documento = linhas[6]
    banco = linhas[1]
    nome = linhas[5]
    data_pagamento = str(linhas[12] ).replace('/','-')
    valor= linhas[15]
    
  return eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento

def get_campos_mentore(numero_pagina, texto):
  eh_inicio_comprovante = texto.find('COMPROVANTE DE PAGAMENTO')>=0 or texto.find('COMPROVANTE DE TRANSFERENCIA')>=0 

  documento = ''
  data_pagamento = ''
  valor = ''
  banco = ''
  nome = ''

  posicao_apos_autenticacao_legis = get_posicao_corte_inicial(texto, 'Autentic. Legis', True)
    
  if posicao_apos_autenticacao_legis>0:
    linhas = texto[posicao_apos_autenticacao_legis:].split('\n')
    
    # cont = 0
    # for linha in linhas:
    #   print(cont,'-', linha)
    #   cont+=1
    
    documento = linhas[6]
    banco = linhas[1]
    nome = linhas[5]
    data_pagamento = str(linhas[12] ).replace('/','-')
    valor= linhas[15]
    
  return eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento


def get_campos_super(numero_pagina, texto):
  eh_inicio_comprovante = texto.find('CNPJ')>=0 

  documento = ''
  data_pagamento = ''
  valor = ''
  banco = ''
  nome = ''

  linhas = texto[:].split('\n')
  
  linha_cpf = 0
  linha_valor = 0

  count = 0
  for linha in linhas:
    if linha.find('CPF')>=0:
      linha_cpf = count
    elif linha.find('R$')>=0:
      linha_valor = count
      
    print(count, linha)

    count+=1

  if eh_inicio_comprovante:
    linhas = texto[:].split('\n')
      
    documento = linhas[linha_cpf]
    documento = documento[-15:]
    banco = 'CONTASUPER'
    nome = linhas[linha_cpf]
    nome = nome[0:nome.find('-')] 
    data_pagamento = str(linhas[linha_valor+1][0:10] ).replace('/','-')
    valor= linhas[linha_valor][0:15]

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
    
    cont = 0
    for linha in linhas:
      if linha.find('R$')>=0:
        valor = linha
        
      if linha.upper().find('SISPAG')>=0:
        posicao_apos_ponto = get_posicao_corte_inicial(linha, '.', True)
        posicao_antes_sispag = get_posicao_corte_inicial(linha, 'SISPAG', False)
        nome = linha[posicao_apos_ponto:posicao_antes_sispag]
        data_pagamento = linha[38:48].replace('/','-')
      
      print(cont, linha)
      cont+=1

    banco = 'Itau'
    documento = linhas[1].replace('  ','').replace(' - ','-')

    print(eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento)
    
    #input()
    
  
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
    
    linha_atento = 0
    cont = 0
    for linha in linhas:
      if linha.find('ATENTO')>=0 and linha.find('BRASIL')>=0:
        linha_atento = cont
      #print(cont, ':', linha)
      cont+=1
      
    valor = linhas[linha_atento+5]
    banco = 'BRADESCO'
    nome = linhas[linha_atento+2]
    
    posicao_apos_espaco = get_posicao_corte_inicial(linhas[linha_atento+3], ' ', True)
    if posicao_apos_espaco>=0:
      documento = linhas[linha_atento+3][0:posicao_apos_espaco].strip()
      data_pagamento = linhas[linha_atento+3][posicao_apos_espaco:].strip().replace('/', '-')


  # print('-'*10)
  # print('linha da atento', linha_atento)
  # print('data_pagamento', data_pagamento)
  # print('valor', valor)
  # print('nome', nome)
  # print('documento', documento)
  # input()
      
  return eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento
  

def get_campos_comprovante(numero_pagina, texto):
  eh_inicio_comprovante = False
  documento = ''
  data_pagamento = ''
  valor = ''
  banco = '.'
  nome = '.'
    
  
  try:
    if (texto.upper().find('SUPER DIGITAL')>0):
      eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_super_digital(numero_pagina, texto)
    elif  (texto.upper().find('CONTASUPER')>0):
      eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_super(numero_pagina, texto)

    elif texto.upper().find('033 - SANTANDER')>0:
      eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_super_digital(numero_pagina, texto)
    elif texto.upper().find('BRADESCO')>0:
      eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_bradesco(numero_pagina, texto)
    elif texto.upper().find('MÊNTORE')>0:
      eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_mentore(numero_pagina, texto)
    else:
      eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_itau(numero_pagina, texto)
  except Exception as e:
    print(texto.upper())
      
  return eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento

def gravar_comprovante(ultimo_nome_cartao, reader, pagina_inicio_documento):
  writer = PdfWriter()

  try:
    writer.add_page(reader.pages[pagina_inicio_documento])
    writer.write(ultimo_nome_cartao)
  finally:
    writer = None

def main():
  lista = lista_arquivos_pdf()
  
  if (lista is None) or (len(lista)==0):
    raise Exception("nenhum arquivo PDF de espelho de cartão de ponto encontrado na pasta origem")
  
  for item in lista:
    reader = PdfReader(item)
    
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      eh_inicio_comprovante, data_pagamento, valor, banco, nome, documento = get_campos_comprovante(cont_pagina, pagina.extract_text())
      
      if eh_inicio_comprovante==True:
        
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)} documento {documento} - {nome}')
        
        nome_arquivo = documento+' - '+nome+' - '+valor+' - '+data_pagamento+' - '+banco
        nome_arquivo = nome_arquivo.replace('/','').replace('\t','').strip()
        
        arquivo_salvar = diretorio_destino+'/'+str(nome_arquivo)+'.pdf'  

        cont_versao = 0        
        while os.path.isfile(arquivo_salvar) :
          cont_versao += 1
          
          arquivo_salvar = diretorio_destino+'/'+str(nome_arquivo)+' (Cópia '+str(cont_versao)+')'+'.pdf'  
        
        gravar_comprovante(arquivo_salvar, reader, cont_pagina)
        
        # print(arquivo_salvar)
        # print(pagina.extract_text())
      else:
        if len(pagina.extract_text().strip())>3:
          print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)}')
          
          arquivo_salvar =  diretorio_destino+'/'+"{:05d}".format(cont_pagina+1)+' nao_gerou.pdf'
          gravar_comprovante(arquivo_salvar, reader, cont_pagina)
        
      cont_pagina += 1
      
    reader = None
    
if __name__ == "__main__":
  main()

