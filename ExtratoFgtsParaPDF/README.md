# ExtratoFgtsParaPDF
Projeto desenvolvido por Thiago Fernandes Cruz para fazer o processamnento em lote dos arquivos oriundos do sistema da Caixa, gerando PDF diretamente
O script gera na pasta destino vários PDFs com os nomes dos funcionários PIS e Competência

### Como configurar:
ediar o arquivo app-config.properties e, editar os par6metros

O diretório origem deve conter os zips baixados diretamente do site da Caixa

### Como executar:
pip install -r requeriments.txt
configuar app-config.properties
python ExtratoFgtsParaPDF.py -> Executar processo