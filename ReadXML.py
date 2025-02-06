import xml.etree.ElementTree as ET
import os
import pandas as pd

# Definir o namespace
ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

# Caminho para a pasta contendo os arquivos XML
caminho_pasta = 'C:\Rodopar xml'

# Lista para armazenar os dados
dados = []

# Iterar sobre todos os arquivos na pasta
for nome_arquivo in os.listdir(caminho_pasta):
    if nome_arquivo.endswith('.xml'):
        caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)
        
        # Carregar o arquivo XML
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()
        
        # Encontrar o campo 'xProd' dentro de 'prod'
        xprod = root.find('.//nfe:prod/nfe:xProd', ns)
        infAdProd = root.find('.//nfe:det/nfe:infAdProd',ns)
        
        # Adicionar os dados à lista
        if xprod is not None:
            dados.append([nome_arquivo, xprod.text,infAdProd.text])
        else:
            dados.append([nome_arquivo, "Campo 'xProd' não encontrado"])

# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(dados, columns=['XML', 'Produto','infAdProd'])

# Salvar o DataFrame em um arquivo Excel
df.to_excel('dados_xprod.xlsx', index=False, engine='openpyxl')

print("Dados salvos em 'dados_xprod.xlsx'")
