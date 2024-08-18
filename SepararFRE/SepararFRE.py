# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
# 05/08/2024

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

def get_funcionario(numero_pagina, texto):
  
  matricula = ''
  nome = ''
  
  linhas = texto[:].split('\n')
    
  for linha in linhas:
    linha = linha.strip()
    elementos = linha.split()
    
    auxiliar = elementos[0].strip().upper()
    
    if auxiliar=='MATRÍCULA:':
      matricula = elementos[1]
      
      lista = elementos[3:]
      nome = ' '.join(lista)
      
  return matricula, nome
      
  

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
    
    ultima_matricula = ''
    
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]

      matricula, nome = get_funcionario(cont_pagina, pagina.extract_text())
      
      if matricula.strip()=='':
        matricula = ultima_matricula
      
      if (ultima_matricula!='') and (matricula!=ultima_matricula) and (ultimo_nome.strip()!=''):
        gravar_cartao(os.path.join(diretorio_destino, ultimo_nome), reader, pagina_inicio_documento, cont_pagina)
        
        pagina_inicio_documento = cont_pagina
      
      if matricula.strip()!='':
        ultimo_nome = f'{nome} - {matricula}'+'.pdf'
        ultima_matricula = matricula
      
      print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)}')
        
      cont_pagina += 1
    
    if ultimo_nome!='' and  (ultimo_nome.strip()!=''):
      gravar_cartao(os.path.join(diretorio_destino, ultimo_nome), reader, pagina_inicio_documento, len(reader.pages))
      
    reader = None
    
if __name__ == "__main__":
  main()

