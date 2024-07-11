# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
# 07/07/2024

# Carregamento de Bibliotecas
from jproperties import Properties
import os
import pathlib 
import pandas as pd
from pathlib import Path
from PyPDF2 import PdfReader
import logging

logger = logging.getLogger("PyPDF2")
logger.setLevel(logging.ERROR)

# Configurações Gerais
configs = Properties()

lista_eventos = (2005, 2006,2007,2015,2016,2017,2045,2046,2055,2056,2090,2092,2100,2102,2130,2140,4230,4231,4232,4235,4236,4237,4280,4285,4287, 22455)

with open('app-config.properties', 'rb') as config_file:
  configs.load(config_file)   
  
diretorio_origem = str(configs.get("diretorio_origem").data)
diretorio_destino = str(configs.get("diretorio_destino").data)   

def lista_arquivos_pdf():
  desktop = pathlib.Path(diretorio_origem)
  return list(desktop.rglob("*.PDF"))

def eh_quantidade(texto):
  if (len(texto)>=4) and (len(texto)<=4):
    try:
      int(texto)
      return True
    except:
      return False

def get_evento_folha(nome_arquivo, texto, ano, mes, empresa):
  global lista_eventos
  
  retorno = None
    
  linhas = texto[:].split('\n')
  
  for linha in linhas:
    linha = linha.strip()
    elementos = linha.split()
    
    try:
      codigo = int(elementos[0])
    except:
      codigo = 0
      
    quantidade_normal = 0
    
    razao_valor_quantidade = 0
      
    if codigo>0:
      for codigo_evento in lista_eventos:
        if codigo_evento==codigo:
          try:
            total = float(elementos[len(elementos)-1].strip().replace('.','').replace(',','.'))
          except:
            total = 0
            
          posicao_quantidade = 0
            
          for elemento in elementos:
            if (posicao_quantidade>0) and eh_quantidade(elemento):
              break
              
            posicao_quantidade+=1

          if posicao_quantidade>0:
            try:
              valor_normal = float(elementos[posicao_quantidade+1].strip().replace('.','').replace(',','.'))
            except:
              valor_normal = 0
              
            if eh_quantidade(elementos[posicao_quantidade]):
              quantidade_normal = int(elementos[posicao_quantidade])
              
            if quantidade_normal>0:
              razao_valor_quantidade = valor_normal / quantidade_normal
            
          dicionario = {'ano':ano,
                        'mes': mes,
                        'empresa': empresa,
                        'arquivo' : nome_arquivo,
                        'codigo_evento' : codigo,
                        'quantidade_normal': quantidade_normal,
                        'valor_normal' : valor_normal,
                        'razao_valor_quantidade_normal': razao_valor_quantidade,
                        'total': total}
        
          if retorno is None:
            retorno = pd.DataFrame([dicionario])
          else:
            retorno = pd.concat([retorno, pd.DataFrame([dicionario])])
        
  return retorno

def main():
  lista = lista_arquivos_pdf()

  if (lista is None) or (len(lista)==0):
    raise Exception("nenhum arquivo PDF de espelho de cartão de ponto encontrado na pasta origem")

  df_retorno = None# pd.DataFrame(columns=['ano', 'mes', 'empresa', 'arquivo','codigo_evento', 'valor'])
  
  posicao = 0
  
  for item in lista:
    reader = PdfReader(item, strict=False)

    posicao+=1
    
    nome_arquivo = Path(item).stem
    
    ano_mes = nome_arquivo[-6:]
    mes = ano_mes[0:2]
    ano = ano_mes[2:6]
    
    elementos_nome = nome_arquivo.replace('__', '_').replace('Resumo_Folha', 'Resumo').split('_')

    empresa = elementos_nome[len(elementos_nome)-2]
    
    linha = f'-{nome_arquivo}'
    print(f'{posicao}/{len(lista)} - {linha}')
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      try:
        df_auxiliar = get_evento_folha(nome_arquivo, pagina.extract_text(), ano, mes, empresa)
        
        if df_retorno is None:
          df_retorno = df_auxiliar
        else:
          df_retorno = pd.concat([df_retorno, df_auxiliar])
      except Exception as e:
          print('Erro:', linha, e)
      
      cont_pagina += 1
      
  
  if df_retorno is not None:
    print(df_retorno.head(20))
    
    df_retorno.to_excel(os.path.join(diretorio_destino, 'eventos_empresa_ano_mes.xlsx'), index=False)

if __name__ == "__main__":
  main()
