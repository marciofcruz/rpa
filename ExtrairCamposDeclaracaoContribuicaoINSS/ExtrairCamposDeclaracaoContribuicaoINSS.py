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

def get_numero(entrada):
  entrada =entrada.replace('\n', '').replace(',','.')
  entrada = entrada.strip()
  try:
    entrada = float(entrada)
  except:
    entrada = 0
    
  return entrada


def get_evento(nome_arquivo, texto):
  
  inscricao = ''
  aliq_rat = 0
  fap = 0 
  rat_ajustado = ''
  competencia = ''
 
  retorno = None
  
  #print(texto)
  
  posicao_atento_brasil = get_posicao_corte_inicial(texto, 'ATENTO BRASIL SA', True)
  
  if posicao_atento_brasil>0:
    inscricao = texto[posicao_atento_brasil:posicao_atento_brasil+20] 
    inscricao = inscricao.strip()
    
  posicao_aliq_rat_depois = get_posicao_corte_inicial(texto, 'ALIQ RAT:', True)
  
  if posicao_aliq_rat_depois>=0:
    posicao_fap_antes =  get_posicao_corte_inicial(texto, 'FAP:', False)
    aliq_rat = texto[posicao_aliq_rat_depois:posicao_fap_antes]
    aliq_rat = get_numero(aliq_rat)
    
  posicao_fap_depois = get_posicao_corte_inicial(texto, 'FAP:', True)
  if posicao_fap_depois>=0:
    fap = texto[posicao_fap_depois:posicao_fap_depois+5]
    fap = get_numero(fap)
    
  posicao_rat_ajustado = get_posicao_corte_inicial(texto, 'RAT AJUSTADO:', True)
  if posicao_rat_ajustado>=0:
    rat_ajustado = texto[posicao_rat_ajustado:posicao_rat_ajustado+6]
    rat_ajustado = get_numero(rat_ajustado)
    
  linhas = texto[:].split('\n')
    
  for linha in linhas:
    posicao_valor_recolher = get_posicao_corte_inicial(linha, 'APURAÇÃO DO VALOR A RECOLHER', True)
    
    if posicao_valor_recolher>=0:
      competencia = linha[-7:]
      break
    
  if competencia != '':
    dicionario = {'nome_arquivo': nome_arquivo,
                  'competencia': competencia,
                  'inscricao': inscricao,
                  'aliq_rat': aliq_rat,
                  'fap': fap,
                  'rat_ajustado': rat_ajustado}
    
  if retorno is None:
    retorno = pd.DataFrame([dicionario])
  else:
    retorno = pd.concat([retorno, pd.DataFrame([dicionario])])
        
  return retorno
  
def main():
  lista = lista_arquivos_pdf()

  if (lista is None) or (len(lista)==0):
    raise Exception("nenhum arquivo PDF de espelho de cartão de ponto encontrado na pasta origem")

  df_retorno = None
  
  posicao = 0
  
  for item in lista:
    reader = PdfReader(item, strict=False)

    posicao+=1
    
    nome_arquivo = Path(item).stem
    
    linha = f'-{nome_arquivo}'
    print(f'{posicao}/{len(lista)} - {linha}')
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      try:
        df_auxiliar = get_evento(nome_arquivo, pagina.extract_text())
        
        if df_retorno is None:
          df_retorno = df_auxiliar
        else:
          df_retorno = pd.concat([df_retorno, df_auxiliar])
      except Exception as e:
          print('Erro:', linha, e)
      
      cont_pagina += 1
      
  
  if df_retorno is not None:
    print(df_retorno.head(20))
    
    df_retorno.to_excel(os.path.join(diretorio_destino, 'campos_comprovante_declaracao_inss.xlsx'), index=False)

if __name__ == "__main__":
  main()
