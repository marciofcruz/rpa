# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
# 27/04/2025

# Carregamento de Bibliotecas
from jproperties import Properties
import os
import pathlib
from PyPDF2 import PdfReader
from PyPDF2 import PdfWriter
from datetime import datetime
import pandas as pd

# Configurações Gerais
configs = Properties()

with open('app-config.properties', 'rb') as config_file:
  configs.load(config_file)   
  
diretorio_origem = str(configs.get("diretorio_origem").data)
diretorio_destino = str(configs.get("diretorio_destino").data)   

def lista_arquivos_pdf():
  desktop = pathlib.Path(diretorio_origem)
  return list(desktop.rglob("*.PDF"))

def get_campos_cartao(arquivo, numero_pagina, texto):
    texto = texto.replace('Aguarde. Processando...','')
    
    if 'Relação de Trabalhadores' in texto:
      ok = True
    
      linhas = texto[:].split('\n')

      empresa = linhas[3].strip()[5:]
      guia = linhas[3].strip()[1:18]
      qtde_trabalhadores = linhas[3].strip()[18:20]
      elementos = linhas[3].split()    
      
      empresa = ' '.join(elementos[5:])
      
      elementos = linhas[4].split()        
      emitida_por = elementos[7]
      
      elementos = linhas[9].split()        
      estabelecimento = elementos[1]
      
      elementos = linhas[13].split()        
      tomador = ' '.join(elementos[2:])
      
      lista_dicionarios = []
      
      cont = 0
      for linha in linhas:
          linha = linha.strip()
          elementos = linha.split()    
          
          posicao_inicial = 0
          
          base_remuneracao_total = 0
          valor_fgts_guia = 0
          juros = 0
          multa = 0
          total = 0
          matricula = ''
          cpf = ''
          comp_apuracao = ''
          nome = ''
          categoria = ''
          tipo_deposito = ''
          
          if cont>=17:
            if elementos[0]=='Rescisório' or elementos[0]=='Mensal':
              tipo_deposito = elementos[0]
              posicao_inicial = 1
            elif elementos[0]=='Verba' or elementos[0]=='Multa':
              tipo_deposito = ' '.join(elementos[0:2])
              posicao_inicial = 2
              
            if posicao_inicial>0:
              
              try:
                vencimento = datetime.strptime(elementos[posicao_inicial+5].strip(), "%d/%m/%Y").date()
              except ValueError:
                vencimento = None
              
              base_remuneracao_total = float(elementos[posicao_inicial].replace('.', '').replace(',', '.'))
              valor_fgts_guia = float(elementos[posicao_inicial+1].replace('.', '').replace(',', '.'))
              juros = float(elementos[posicao_inicial+2].replace('.', '').replace(',', '.'))
              multa = float(elementos[posicao_inicial+3].replace('.', '').replace(',', '.'))
              total = float(elementos[posicao_inicial+4].replace('.', '').replace(',', '.'))
              
              matricula = elementos[posicao_inicial+6]
              cpf = elementos[posicao_inicial+7]
              comp_apuracao = elementos[posicao_inicial+9]
              nome = ' '.join(elementos[posicao_inicial+10:-1])
              categoria = elementos[-1]
              
              lista_dicionarios.append({
                    'arquivo': arquivo,
                    'pagina:': numero_pagina+1,
                    'Empresa': empresa,
                    'Qtd. Trabalhadores FGTS': qtde_trabalhadores,              
                    'Estabelecimento': estabelecimento,
                    'Tomador': tomador,
                    'Guia': guia,
                    'Comp. Apuração': comp_apuracao,
                    'Matrícula': matricula,
                    'CPF': cpf,
                    'Nome': nome,
                    'Categoria': categoria,
                    'Vencimento': vencimento,
                    'Tipo Depósito': tipo_deposito,
                    'Base Remuneração Total': base_remuneracao_total,
                    'Valor FGTS Categoria na Guia': valor_fgts_guia,
                    'Juros': juros,
                    'Multa': multa,
                    'Total': total})
          
          cont += 1
              
      return True, lista_dicionarios
    else:
      return False, None
    
def main():
  lista = lista_arquivos_pdf()
  
  if (lista is None) or (len(lista)==0):
    raise Exception("nenhum arquivo PDF de espelho de cartão de ponto encontrado na pasta origem")
  
  df = None
  
  for item in lista:
    print('Arquivo: ', item)
    
    reader = PdfReader(item)
    
    cont_pagina = 0 
    while cont_pagina<=(len(reader.pages)-1):
      pagina = reader.pages[cont_pagina]
      
      continuar, dicionario = get_campos_cartao(os.path.basename(item), cont_pagina, pagina.extract_text())
      
      if continuar:
        if df is None:
          df = pd.DataFrame(dicionario)
        else:
          df = pd.concat([df, pd.DataFrame(dicionario)], ignore_index=True)
          
      cont_pagina += 1
        
  timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
  output_filename = f"Guia_GFD_Detalhada_{timestamp}.xlsx"
  df.to_excel(os.path.join(diretorio_destino, output_filename), index=False)
      
    
if __name__ == "__main__":
  main()

