import time

def txtprogramação(date, horario, nome, telefone, placa, tipo, trans, forn, prod, carga, nf1, peso1, nf2=None, peso2=None, nf3=None, peso3=None):
    try:
        with open('meuarquivo.txt', 'w') as arquivo:
            if nf2 is None and nf3 is None:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1} {forn} {int(peso1)} {placa} {tipo} {prod} {carga} {trans}\n')
            elif nf3 is None:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1}/{nf2} {forn} {int(peso1 + peso2)} {placa} {tipo} {prod} {carga} {trans}\n')
            else:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1}/{nf2}/{nf3} {forn} {int(peso1 + peso2 + peso3)} {placa} {tipo} {prod} {carga} {trans}\n')

        print('Texto escrito com sucesso!')
        import os  # Importe os dentro da função, se não for usar em outro lugar
        time.sleep(3)
        os.startfile('meuarquivo.txt')
    except Exception as e:
        print(f"Ocorreu um erro: {e}") # Imprime o erro para ajudar na depuração.
        # Aqui você pode retornar o erro para ser tratado em main.py, se necessário.
        return f"Erro: {e}" # Exemplo: retorna uma string com a mensagem de erro.