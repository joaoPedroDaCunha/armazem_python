import xml.etree.ElementTree as ET
import os
import pandas as pd

# Definir o namespace
ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

# Caminho para a pasta contendo os arquivos XML
caminho_pasta = 'C:\\Rodopar xml'

# Lista para armazenar os dados
dados = []

# Iterar sobre todos os arquivos na pasta
for nome_arquivo in os.listdir(caminho_pasta):
    if nome_arquivo.endswith('.xml'):
        caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)
        
        # Carregar o arquivo XML
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()
        
        # Encontrar todos os campos 'xProd' dentro de 'prod'
        produtos = root.findall('.//nfe:prod/nfe:xProd', ns)
        infAdProds = root.findall('.//nfe:det/nfe:infAdProd', ns)
        
        # Adicionar os dados à lista
        for prod, infAdProd in zip(produtos, infAdProds):
            if prod is not None:
                dados.append([nome_arquivo, prod.text, infAdProd.text if infAdProd is not None else ""])
            else:
                dados.append([nome_arquivo, "Campo 'xProd' não encontrado", ""])

# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(dados, columns=['XML', 'Produto', 'infAdProd'])

# Verificar se existem duplicatas no campo 'Produto'
duplicados = df[df.duplicated(subset=['XML', 'Produto'], keep=False)]

# Aplicar estilo para pintar de amarelo as linhas com duplicatas
def highlight_duplicates(row):
    if row.name in duplicados.index:
        return ['background-color: yellow'] * len(row)
    return [''] * len(row)

# Aplicar a formatação ao DataFrame
styled_df = df.style.apply(highlight_duplicates, axis=1)

# Salvar o DataFrame estilizado em um arquivo Excel
styled_df.to_excel('dados_xprod.xlsx', index=False, engine='openpyxl')

print("Dados salvos em 'dados_xprod.xlsx' com as duplicatas destacadas em amarelo.")
