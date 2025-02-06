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

def get_campos_informe(texto):
    texto = texto.replace('Aguarde. Processando...','')
    
    eh_inicio_rendimento = texto.find('Rendimentos Tributáveis')>0
    
    exercicio = ''
    nome = ''
    
    posicao_exercicio_depois =  get_posicao_corte_inicial(texto, 'Exercício de', True)
    if posicao_exercicio_depois>0:
      exercicio = texto[posicao_exercicio_depois:posicao_exercicio_depois+5].strip()
    
    posicao_apos_nome_completo = get_posicao_corte_inicial(texto, ' Nome Completo ', True)
    posicao_antes_lotacao  = get_posicao_corte_inicial(texto, 'Lotação', False)
    
    if posicao_apos_nome_completo>0 and posicao_antes_lotacao>0:
      nome = str(texto[posicao_apos_nome_completo:posicao_antes_lotacao]).strip()
    
    return eh_inicio_rendimento, exercicio, nome


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
    
    ultimo_nome_informe = ''
    
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      eh_inicio_informe, exercicio, nome = get_campos_informe(pagina.extract_text())
      
      if eh_inicio_informe==True:
        
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)} - {nome}')
        
        if ultimo_nome_informe!='':
          gravar_cartao(ultimo_nome_informe, reader, pagina_inicio_documento, cont_pagina)

        nome_arquivo = 'Informe IRPF '+exercicio+' - '+nome
        
        ultimo_nome_informe = diretorio_destino+'/'+str(nome_arquivo)+'.pdf'  
        
        pagina_inicio_documento = cont_pagina
      else:
        print(f'{item} - Página {cont_pagina+1}/{len(reader.pages)}')
        
      cont_pagina += 1
    
    if ultimo_nome_informe!='':
      gravar_cartao(ultimo_nome_informe, reader, pagina_inicio_documento, len(reader.pages))
      
    reader = None
    
if __name__ == "__main__":
  main()

