# Script desenvolvido por Thiago Fernandes Cruz - thiago@fernandescruz.com
# 18/03/2024

# Carregamento de Bibliotecas
from jproperties import Properties
import os
import pathlib
import pandas as pd
import shutil
from zipfile import ZipFile
from fpdf import FPDF
from PyPDF2 import PdfFileMerger

# Configurações Gerais
configs = Properties()

with open('app-config.properties', 'rb') as config_file:
  configs.load(config_file)   
  
diretorio_origem = str(configs.get("diretorio_origem").data)
diretorio_temp = str(configs.get("diretorio_temp").data)
diretorio_destino = str(configs.get("diretorio_destino").data)   

separador = '-' * 73

# Funçòes Auxiliares
def remover_temp():
  desktop = pathlib.Path(diretorio_temp)
  lista_arquivo = list(desktop.rglob("*.txt"))

  for item in lista_arquivo:
    os.remove(item)

  lista_arquivo = list(desktop.rglob("*"))

  for item in lista_arquivo:
    os.rmdir(item)
    
def lista_arquivos_txt():
  desktop = pathlib.Path(diretorio_temp)
  return list(desktop.rglob("*.txt"))

def create_pdf(conteudo, output_file):
  # Create a new FPDF object
  pdf = FPDF()

  # Add a new page to the PDF
  pdf.add_page()

  # Set the font and font size
  pdf.set_font('Courier', size=9)

  # Write the text to the PDF
  pdf.write(2, conteudo)

  # Save the PDF
  pdf.output(output_file)


def get_lista_arquivos_zip():
  # Pegar Arquivos do tipo PDF e alimentar num "list"
  return  [f for f in os.listdir(diretorio_origem) if pathlib.Path(f).suffix.lower()=='.zip']
    
def main():
  # Retornar lista de Arquivos

  lista_arquivos_zip = get_lista_arquivos_zip()

  if (lista_arquivos_zip is None) or (len(lista_arquivos_zip)==0):
    raise Exception("nenhum arquivo encontrado na pasta origem")

  for nome_arquivo in lista_arquivos_zip:
    caminho_arquivo = diretorio_origem+'/'+nome_arquivo
    
    diretorio_zip = os.path.join(diretorio_temp, nome_arquivo)
    
    print(diretorio_zip)

    with ZipFile(caminho_arquivo, 'r') as zip_file:
      zip_file.extractall(diretorio_zip)
      
  # Separar o Conteudo dos arquivos
  lista = lista_arquivos_txt()

  for item in lista:
    with open(item, 'r') as arquivo:
      conteudo = arquivo.read()
      
      extratos = conteudo.split(separador)
      
      for extrato in extratos:
        if extrato.find('FGTS')>0:
          cnpj = extrato[2386:2390].strip()
          nome = extrato[1124:1167].strip()
          pis = extrato[1523:1536].strip()
          maior_competencia = extrato[2723:2745].strip().replace('/','')

          nome_arquivo = cnpj+' - '+nome + ' - '+pis + ' - '+maior_competencia
          caminho_pdf = diretorio_destino+'/'+str(nome_arquivo)+'.pdf'  
          
          print(caminho_pdf)
          
          create_pdf(extrato, caminho_pdf)
      

if __name__ == "__main__":
  main()

