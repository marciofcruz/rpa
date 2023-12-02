from jproperties import Properties
import os
import pathlib
import pandas as pd
import shutil

cont_pasta = 0

configs = Properties()
  
with open('app-config.properties', 'rb') as config_file:
  configs.load(config_file)
   
def get_lista_arquivos():
  diretorio_origem = str(configs.get("diretorio_origem").data)
      
  # Pegar Arquivos do tipo PDF e alimentar num "list"
  return  [f for f in os.listdir(diretorio_origem) if pathlib.Path(f).suffix.lower()=='.pdf']
  
def get_df_arquivos(lista_arquivos):
  df_arquivos = pd.DataFrame(columns=['arquivo_origem'])

  for arquivo in lista_arquivos:
    dicionario = dict(
      arquivo_origem = os.path.join(str(configs.get("diretorio_origem").data), arquivo)
      
    )
    
    df_arquivos = pd.concat([df_arquivos, pd.DataFrame([dicionario])])
    
  return df_arquivos

def mover_df_arquivos(df):
  diretorio_destino= str(configs.get("diretorio_destino").data)
  
  def get_pasta_viavel():
    global cont_pasta

    pasta_candidata = os.path.join(diretorio_destino, str(cont_pasta))
    
    if not os.path.exists(pasta_candidata):
      os.mkdir(pasta_candidata)
      return pasta_candidata
    else:
      lista_arquivos_existentes = [f for f in os.listdir(pasta_candidata) if pathlib.Path(f).suffix.lower()=='.pdf']
      
      if len(lista_arquivos_existentes) < int(configs.get("maximo_arquivos_pasta").data):
        return pasta_candidata
      else:
        cont_pasta = cont_pasta + 1
        return get_pasta_viavel()
  
  def get_destino(tupla):
    destino = os.path.join(get_pasta_viavel(), os.path.basename(tupla["arquivo_origem"]))  
    
    shutil.move(tupla["arquivo_origem"], destino)
    
    return destino 
  
  df["arquivo_destino"] = df.apply(get_destino, axis=1)
  
  return df
  
   
def main():
  # Retornar lista de Arquivos
  lista_arquivos = get_lista_arquivos()

  if lista_arquivos is None or len(lista_arquivos)==0:
    raise Exception("nenhum arquivo encontrado na pasta ", str(configs.get("diretorio_origem").data))

  # Preparar Data Frame com a coluna origem
  df_arquivos = get_df_arquivos(lista_arquivos)
  
  # Ajustar Coluna Destino
  df_arquivos = mover_df_arquivos(df_arquivos)
    
  print(df_arquivos.head(30))
      

if __name__ == "__main__":
  main()