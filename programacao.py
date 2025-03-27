import time
import os

def txtprogramação(date, horario, nome, telefone, placa, tipo, trans, forn, prod, carga, nf1, peso1, nf2=None, peso2=None, nf3=None, peso3=None, checkbox2=None ,checkbox3=None):
    try:
        with open('meuarquivo.txt', 'w') as arquivo:
            # Escreve uma linha no arquivo
            if checkbox2 == 0 and checkbox3 == 0 :
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1} {forn} {peso1} {placa} {tipo} {prod} {carga} {trans}\n')
            if checkbox2 == 1 and checkbox3 == 0 :
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1}/{nf2} {forn} {peso1+peso2} {placa} {tipo} {prod} {carga} {trans}\n')
            if checkbox2 == 1 and checkbox3 == 1 :
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1}/{nf2}/{nf3} {forn} {peso1+peso2+peso3} {placa} {tipo} {prod} {carga} {trans}\n')

        print('Texto escrito com sucesso!')

        time.sleep(3)
        # Abre o arquivo txt automaticamente no editor de texto padrão do sistema
        os.startfile('meuarquivo.txt')
    except Exception as e:
        print(e)