# Carregando Bibliotecas
from jproperties import Properties

# Inicializando Variáveis Globais
configs = Properties()

with open('app-config.properties', 'rb') as config_file:
   configs.load(config_file)

# Funções do script
def main():
   # print(str(configs.get("diretorio_origem").data))

   for i in range(int(configs.get("quantidade_arquivos_origem").data)):
      nome_arquivo_fake = str(configs.get("diretorio_origem").data)+'\\arquivo '+str(i)+'.pdf'
      
      with open(nome_arquivo_fake, "w") as arquivo_fake:
         arquivo_fake.write('Escrevi algo para ter alguns bytes')
      
# Sequencia principal de execução
if __name__ == "__main__":
  main()