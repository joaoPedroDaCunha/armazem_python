import os

# Abre (ou cria) o arquivo txt em modo de escrita (write mode)
with open('meuarquivo.txt', 'w') as arquivo:
    # Escreve uma linha no arquivo
    arquivo.write('kkkkkkkkkkkkkkkkkkkkkkkkkkk\n')

print('Texto escrito com sucesso!')

# Abre o arquivo txt automaticamente no editor de texto padrão do sistema
os.startfile('meuarquivo.txt')
