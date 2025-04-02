import time
import os
from tkinter import messagebox
import pythoncom

def txtprogramação(date, horario, nome, telefone, placa, tipo, trans, forn, prod, carga, nf1, peso1, nf2, peso2, nf3, peso3, checkbox2 ,checkbox3):
    pythoncom.CoInitialize()
    try:
        with open('meuarquivo.txt', 'w') as arquivo:

            if nf2 == None and nf3 == None :
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1} {forn} {peso1} {placa} {tipo} {prod} {carga} {trans}\n')
            if nf2 != None and nf3 == None :
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1}/{nf2} {forn} {peso1+peso2} {placa} {tipo} {prod} {carga} {trans}\n')
            if nf1 != None and nf3 != None :
                arquivo.write(f'{date} {horario} {nome} {telefone} {nf1}/{nf2}/{nf3} {forn} {peso1+peso2+peso3} {placa} {tipo} {prod} {carga} {trans}\n')

        print('Texto escrito com sucesso!')

        time.sleep(1)
        os.startfile('meuarquivo.txt')
    except Exception as e:
        messagebox.showerror(e)
    finally:
        pythoncom.CoUninitialize()

def txtprogramaçãoEmb(date,horario,nome,telefone,placa,tipo,trans,forn,qtdtotalEmb,nfembalagem1,nfpaleteEmb1,codprod1,qtdpaleteEmb1,valEmb1,nomeprod1,contUnid1,lotef1,pesoEmb1,nfembalagem2,nfpaleteEmb2,codprod2,qtdpaleteEmb2,valEmb2,nomeprod2,contUnid2,lotef2,pesoEmb2
              ,nfembalagem3,nfpaleteEmb3,codprod3,qtdpaleteEmb3,valEmb3,nomeprod3,contUnid3,lotef3,pesoEmb3,nfembalagem4,nfpaleteEmb4,codprod4,qtdpaleteEmb4,valEmb4,nomeprod4,contUnid4,lotef4,pesoEmb4
              ,nfembalagem5,nfpaleteEmb5,codprod5,qtdpaleteEmb5,valEmb5,nomeprod5,contUnid5,lotef5,pesoEmb5,nfembalagem6,nfpaleteEmb6,codprod6,qtdpaleteEmb6,valEmb6,nomeprod6,contUnid6,lotef6,pesoEmb6):
    pythoncom.CoInitialize()
    try:
        with open('meuarquivo.txt', 'w') as arquivo:
            if nfembalagem2 == None and nfembalagem3 == None and nfembalagem4 == None and nfembalagem5 == None and nfpaleteEmb6 == None:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nfembalagem1} Embalagem {forn} {placa} {tipo} {trans}')
            elif nfembalagem3 == None and nfembalagem4 == None and nfembalagem5 == None and nfembalagem6:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nfembalagem1}/{nfembalagem2} Embalagem {forn} {placa} {tipo} {trans}')
            elif nfembalagem4 == None and nfembalagem5 == None and nfembalagem6 ==None:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nfembalagem1}/{nfembalagem2}/{nfembalagem3} Embalagem {forn} {placa} {tipo} {trans}')
            elif nfembalagem5 == None and nfembalagem6 == None:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nfembalagem1}/{nfembalagem2}/{nfembalagem3}/{nfembalagem4} Embalagem {forn} {placa} {tipo} {trans}')
            elif nfembalagem6 == None:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nfembalagem1}/{nfembalagem2}/{nfembalagem3}/{nfembalagem4}/{nfembalagem5} Embalagem {forn} {placa} {tipo} {trans}')
            else:
                arquivo.write(f'{date} {horario} {nome} {telefone} {nfembalagem1}/{nfembalagem2}/{nfembalagem3}/{nfembalagem4}/{nfembalagem5}/{nfembalagem6} Embalagem {forn} {placa} {tipo} {trans}')

        print('Texto escrito com sucesso!')

        time.sleep(1)
        os.startfile('meuarquivo.txt')
    except Exception as e:
        messagebox.showerror(e)
    finally:
        pythoncom.CoUninitialize()