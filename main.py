import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import os
from threading import Thread
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from print import imprimirespelho
from programacao import txtprogramação as txt
from salvar_dados import salvar_dados as salvar,salvarEmb

def save():
    try:
        salvar(
        entry_date.get(),  # Correto: o método .get() é chamado no widget antes de usar o valor
        entry_horario.get(),
        entry_nome.get(),
        entry_telefone.get(),
        entry_placa.get(),
        entry_tipo.get(),
        entry_trans.get(),
        combobox_forn.get(),
        combobox_prod.get(),
        combobox_carga.get(),
        entry_val.get(),
        entry_nfsal1.get(),
        entry_nfpalete1.get(),
        int(entry_qtdpalete1.get() or "0"),  # Correto: .get() chamado antes de int()
        entry_lotesal1.get(),
        int(entry_peso1.get() or "0"),       # Correto: .get() chamado antes de int()
        checkbox_lote2_var,           # Correto: chamada direta no IntVar
        entry_nfsal2.get(),
        entry_nfpalete2.get(),
        int(entry_qtdpalete2.get() or "0"),  # Correto: .get() chamado antes de int()
        entry_lotesal2.get(),
        int(entry_peso2.get() or "0"),       # Correto: .get() chamado antes de int()
        checkbox_lote3_var,           # Correto: chamada direta no IntVar
        entry_nfsal3.get(),
        entry_nfpalete3.get(),
        int(entry_qtdpalete3.get() or "0"),  # Correto: .get() chamado antes de int()
        entry_lotesal3.get(),
        int(entry_peso3.get() or "0")        # Correto: .get() chamado antes de int()
        )
    except ValueError as e:
        messagebox.showerror("Erro de valor", "Nos campos de QTD e Peso devese colocar exclusivamente numeros")
    except PermissionError as e:
        messagebox.showerror("Erro de permissão", f"Permissão negada: {e}. Verifique se o arquivo está aberto em outro programa.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def saveEmb():
    salvarEmb(entry_dateEmb.get(),entry_horarioEmb.get(),entry_nomeEmb.get(),entry_telefoneEmb.get(),entry_placaEmb.get(),entry_tipoEmb.get(),entry_transEmb.get(),combobox_fornEmb.get(),entry_qtdtotalEmb.get()
              ,entry_nfembalagem1.get(),entry_nfpaleteEmb1.get(),entry_codprod1.get(),entry_qtdpaleteEmb1.get(),entry_valEmb1.get(),combobox_nomeprod1.get(),entry_contUnid1.get(),entry_lotef1.get(),entry_pesoEmb1.get()
              ,entry_nfembalagem2.get() or None,entry_nfpaleteEmb2.get() or None,entry_codprod2.get() or None,entry_qtdpaleteEmb2.get() or None,entry_valEmb2.get() or None,combobox_nomeprod2.get() or None,entry_contUnid2.get() or None,entry_lotef2.get() or None,entry_pesoEmb2.get() or None
              ,entry_nfembalagem3.get() or None,entry_nfpaleteEmb3.get() or None,entry_codprod3.get() or None,entry_qtdpaleteEmb3.get() or None,entry_valEmb3.get() or None,combobox_nomeprod3.get() or None,entry_contUnid3.get() or None,entry_lotef3.get() or None,entry_pesoEmb3.get() or None
              ,entry_nfembalagem4.get() or None,entry_nfpaleteEmb4.get() or None,entry_codprod4.get() or None,entry_qtdpaleteEmb4.get() or None,entry_valEmb4.get() or None,combobox_nomeprod4.get() or None,entry_contUnid4.get() or None,entry_lotef4.get() or None,entry_pesoEmb4.get() or None
              ,entry_nfembalagem5.get() or None,entry_nfpaleteEmb5.get() or None,entry_codprod5.get() or None,entry_qtdpaleteEmb5.get() or None,entry_valEmb5.get() or None,combobox_nomeprod5.get() or None,entry_contUnid5.get() or None,entry_lotef5.get() or None,entry_pesoEmb5.get() or None
              ,entry_nfembalagem6.get() or None,entry_nfpaleteEmb6.get() or None,entry_codprod6.get() or None,entry_qtdpaleteEmb6.get() or None,entry_valEmb6.get() or None,combobox_nomeprod6.get() or None,entry_contUnid6.get() or None,entry_lotef6.get() or None,entry_pesoEmb6.get() or None
              )

def prog():
    try:
        threadtxt = Thread(target=txt, args=(entry_date.get(),entry_horario.get(),entry_nome.get(),entry_telefone.get(),entry_placa.get(),entry_tipo.get(),entry_trans.get(),combobox_forn.get(),combobox_prod.get(),combobox_carga.get(),entry_nfsal1.get(),int(entry_peso1.get() or "0"),entry_nfsal2.get(),int(entry_peso2.get() or "0"),entry_nfsal3.get(),int(entry_peso3.get() or "0"),checkbox_lote2_var.get(),checkbox_lote3_var.get()))
        threadtxt.daemon
        threadtxt.start()
        #txt(entry_date.get(),entry_horario.get(),entry_nome.get(),entry_telefone.get(),entry_placa.get(),entry_tipo.get(),entry_trans.get(),combobox_forn.get(),combobox_prod.get(),combobox_carga.get(),entry_nfsal1.get(),int(entry_peso1.get() or "0"),entry_nfsal2.get(),int(entry_peso2.get() or "0"),entry_nfsal3.get(),int(entry_peso3.get() or "0"),checkbox_lote2_var.get(),checkbox_lote3_var.get())
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def limpar():
    try:
        entry_date.delete(0,tk.END)
        entry_horario.delete(0,tk.END)
        entry_nome.delete(0,tk.END)
        entry_telefone.delete(0,tk.END)
        entry_placa.delete(0,tk.END)
        entry_tipo.delete(0,tk.END)
        entry_trans.delete(0,tk.END)
        entry_val.delete(0,tk.END)
        entry_nfsal1.delete(0,tk.END)
        entry_nfpalete1.delete(0,tk.END)
        entry_qtdpalete1.delete(0,tk.END)
        entry_lotesal1.delete(0,tk.END)
        entry_peso1.delete(0,tk.END)
        entry_nfsal2.delete(0,tk.END)
        entry_nfpalete2.delete(0,tk.END)
        entry_qtdpalete2.delete(0,tk.END)
        entry_lotesal2.delete(0,tk.END)
        entry_peso2.delete(0,tk.END)
        entry_nfsal3.delete(0,tk.END)
        entry_nfpalete3.delete(0,tk.END)
        entry_qtdpalete3.delete(0,tk.END)
        entry_lotesal3.delete(0,tk.END)
        entry_peso3.delete(0,tk.END)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")


# Cria a janela principal
janela = tk.Tk()
janela.title("Cadastro de Informações")

# Cria o Notebook (controle de abas)
notebook = ttk.Notebook(janela)
notebook.pack(fill='both', expand=True)

# Cria o frame para a primeira aba
aba1 = ttk.Frame(notebook)
notebook.add(aba1, text="Cadastro Sal")

# Cria o frame para a segunda aba
aba2 = ttk.Frame(notebook)
notebook.add(aba2, text="Cadastro Embalagem")

# Aba 1: Cadastro Principal (Seu layout original)
label_date = tk.Label(aba1, text="Data de entrada:")
label_date.grid(row=0, column=0)
entry_date = tk.Entry(aba1)
entry_date.grid(row=0, column=1)

label_horario = tk.Label(aba1, text="Horário de chegada:")
label_horario.grid(row=0, column=3)
entry_horario = tk.Entry(aba1)
entry_horario.grid(row=0, column=4)

# Adicione aqui o restante do layout da primeira aba (igual ao seu código original)
label_espaco = tk.Label(aba1)
label_espaco.grid(row=1)

label_nome = tk.Label(aba1, text="Nome do motorista:")
label_nome.grid(row=2, column=0)
entry_nome = tk.Entry(aba1)
entry_nome.grid(row=2, column=1)

label_telefone = tk.Label(aba1, text="Telefone:")
label_telefone.grid(row=2, column=3)
entry_telefone = tk.Entry(aba1)
entry_telefone.grid(row=2, column=4)

label_espaco = tk.Label(aba1)
label_espaco.grid(row=3)

label_placa = tk.Label(aba1, text="Placa do veículo:")
label_placa.grid(row=4, column=0)
entry_placa = tk.Entry(aba1)
entry_placa.grid(row=4, column=1)

label_tipo = tk.Label(aba1, text="Tipo de Veículo:")
label_tipo.grid(row=4, column=3)
entry_tipo = tk.Entry(aba1)
entry_tipo.grid(row=4, column=4)

label_espaco = tk.Label(aba1)
label_espaco.grid(row=5)

label_trans = tk.Label(aba1, text="Transportadora:")
label_trans.grid(row=6, column=0)
entry_trans = tk.Entry(aba1)
entry_trans.grid(row=6, column=1)

label_forn = tk.Label(aba1, text="Fornecedor:")
label_forn.grid(row=7, column=0)

combobox_forn = ttk.Combobox(aba1, values=["NORSAL", "CIMSAL", "CISNE"])
combobox_forn.grid(row=7, column=1)
combobox_forn.set("Selecione uma opção")

label_prod = tk.Label(aba1, text="Produto:")
label_prod.grid(row=8, column=0)

combobox_prod = ttk.Combobox(aba1, values=["SAL REFINADO", "SAL GRANULADO GROSSO", "SAL EXTRA REFINADO","SAL ENTREFINO"])
combobox_prod.grid(row=8, column=1)
combobox_prod.set("Selecione uma opção")

label_carga = tk.Label(aba1, text="Tipo de Carga:")
label_carga.grid(row=8, column=3)

combobox_carga = ttk.Combobox(aba1, values=["BIG BAG", "25 KG"])
combobox_carga.grid(row=8, column=4)
combobox_carga.set("Selecione uma opção")

label_val = tk.Label(aba1, text="Validade:")
label_val.grid(row=10, column=0)
entry_val = tk.Entry(aba1)
entry_val.grid(row=10, column=1)

label_nfsal1 = tk.Label(aba1, text="NF do produto:")
label_nfsal1.grid(row=11, column=0)

entry_nfsal1 = tk.Entry(aba1)
entry_nfsal1.grid(row=11, column=1)

label_lotesal1 = tk.Label(aba1, text="Lote:")
label_lotesal1.grid(row=11, column=2)

entry_lotesal1 = tk.Entry(aba1)
entry_lotesal1.grid(row=11, column=3)

label_nfpalete1 = tk.Label(aba1, text="NF DO Palete:")
label_nfpalete1.grid(row=11, column=4)

entry_nfpalete1 = tk.Entry(aba1)
entry_nfpalete1.grid(row=11, column=5)

label_qtdpalete1 = tk.Label(aba1, text="QTD Palete:")
label_qtdpalete1.grid(row=11, column=6)

entry_qtdpalete1 = tk.Entry(aba1)
entry_qtdpalete1.grid(row=11, column=7)

label_peso1 = tk.Label(aba1, text="peso:")
label_peso1.grid(row=11, column=8)

entry_peso1 = tk.Entry(aba1)
entry_peso1.grid(row=11, column=9)

checkbox_lote2_var = tk.IntVar()
checkbox_lote2 = tk.Checkbutton(aba1, text="Incluir Lote 2", variable=checkbox_lote2_var)
checkbox_lote2.grid(row=12, column=0)

label_nfsal2 = tk.Label(aba1, text="NF do produto:")
label_nfsal2.grid(row=13, column=0)

entry_nfsal2 = tk.Entry(aba1)
entry_nfsal2.grid(row=13, column=1)

label_lotesal2 = tk.Label(aba1, text="Lote:")
label_lotesal2.grid(row=13, column=2)

entry_lotesal2 = tk.Entry(aba1)
entry_lotesal2.grid(row=13, column=3)

label_nfpalete2 = tk.Label(aba1, text="NF DO Palete:")
label_nfpalete2.grid(row=13, column=4)

entry_nfpalete2 = tk.Entry(aba1)
entry_nfpalete2.grid(row=13, column=5)

label_qtdpalete2 = tk.Label(aba1, text="QTD Palete:")
label_qtdpalete2.grid(row=13, column=6)

entry_qtdpalete2 = tk.Entry(aba1)
entry_qtdpalete2.grid(row=13, column=7)

label_peso2 = tk.Label(aba1, text="peso:")
label_peso2.grid(row=13, column=8)

entry_peso2 = tk.Entry(aba1)
entry_peso2.grid(row=13, column=9)

checkbox_lote3_var = tk.IntVar()
checkbox_lote3 = tk.Checkbutton(aba1, text="Incluir Lote 3", variable=checkbox_lote3_var)
checkbox_lote3.grid(row=14, column=0)

label_nfsal3 = tk.Label(aba1, text="NF do produto:")
label_nfsal3.grid(row=15, column=0)

entry_nfsal3 = tk.Entry(aba1)
entry_nfsal3.grid(row=15, column=1)

label_lotesal3 = tk.Label(aba1, text="Lote:")
label_lotesal3.grid(row=15, column=2)

entry_lotesal3 = tk.Entry(aba1)
entry_lotesal3.grid(row=15, column=3)

label_nfpalete3 = tk.Label(aba1, text="NF DO Palete:")
label_nfpalete3.grid(row=15, column=4)

entry_nfpalete3 = tk.Entry(aba1)
entry_nfpalete3.grid(row=15, column=5)

label_qtdpalete3 = tk.Label(aba1, text="QTD Palete:")
label_qtdpalete3.grid(row=15, column=6)

entry_qtdpalete3 = tk.Entry(aba1)
entry_qtdpalete3.grid(row=15, column=7)

label_peso3 = tk.Label(aba1, text="peso:")
label_peso3.grid(row=15, column=8)

entry_peso3 = tk.Entry(aba1)
entry_peso3.grid(row=15, column=9)


# Cria o botão para salvar
botao_salvar = tk.Button(aba1, text="Salvar", command=save)
botao_salvar.grid(row=16, columnspan=5)

botao_imprimir = tk.Button(aba1,text="Imprimir Espelho",command=imprimirespelho)
botao_imprimir.grid(row=16, columnspan=9)

botao_limpar = tk.Button(aba1,text="Limpar informações",command=limpar)
botao_limpar.grid(row=1, column=8)

botao_programaçao = tk.Button(aba1,text="Programação txt",command=prog)
botao_programaçao.grid(row=16, column=8)


# Aba 2: Outra aba com layout diferente
label_date = tk.Label(aba2, text="Data de entrada:")
label_date.grid(row=0, column=0)
entry_dateEmb = tk.Entry(aba2)
entry_dateEmb.grid(row=0, column=1)

label_horario = tk.Label(aba2, text="Horário de chegada:")
label_horario.grid(row=0, column=3)
entry_horarioEmb = tk.Entry(aba2)
entry_horarioEmb.grid(row=0, column=4)

# Adicione aqui o restante do layout da primeira aba (igual ao seu código original)
label_espaco = tk.Label(aba2)
label_espaco.grid(row=1)

label_nome = tk.Label(aba2, text="Nome do motorista:")
label_nome.grid(row=2, column=0)
entry_nomeEmb = tk.Entry(aba2)
entry_nomeEmb.grid(row=2, column=1)

label_telefone = tk.Label(aba2, text="Telefone:")
label_telefone.grid(row=2, column=3)
entry_telefoneEmb = tk.Entry(aba2)
entry_telefoneEmb.grid(row=2, column=4)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=3)

label_placa = tk.Label(aba2, text="Placa do veículo:")
label_placa.grid(row=4, column=0)
entry_placaEmb = tk.Entry(aba2)
entry_placaEmb.grid(row=4, column=1)

label_tipo = tk.Label(aba2, text="Tipo de Veículo:")
label_tipo.grid(row=4, column=3)
entry_tipoEmb = tk.Entry(aba2)
entry_tipoEmb.grid(row=4, column=4)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=5)

label_trans = tk.Label(aba2, text="Transportadora:")
label_trans.grid(row=6, column=0)
entry_transEmb = tk.Entry(aba2)
entry_transEmb.grid(row=6, column=1)

label_forn = tk.Label(aba2, text="Fornecedor:")
label_forn.grid(row=7, column=0)

combobox_fornEmb = ttk.Combobox(aba2, values=[])
combobox_fornEmb.grid(row=7, column=1)

label_qtdtotal = tk.Label(aba2, text="Quantidade total de Paletes")
label_qtdtotal.grid(row=7, column=3)

entry_qtdtotalEmb = tk.Entry(aba2)
entry_qtdtotalEmb.grid(row=7,column=4)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=8)

label_info1 = tk.Label(aba2, text="Produto 1")
label_info1.grid(row=9,column=1)

label_info2 = tk.Label(aba2, text="Produto 2")
label_info2.grid(row=9,column=2)

label_info3 = tk.Label(aba2, text="Produto 3")
label_info3.grid(row=9,column=3)

label_info4 = tk.Label(aba2, text="Produto 4")
label_info4.grid(row=9,column=4)

label_info5 = tk.Label(aba2, text="Produto 5")
label_info5.grid(row=9,column=5)

label_info6 = tk.Label(aba2, text="Produto 6")
label_info6.grid(row=9,column=6)

label_nfembalagem = tk.Label(aba2, text="NF PRODUTO")
label_nfembalagem.grid(row=10,column=0)

entry_nfembalagem1 = tk.Entry(aba2)
entry_nfembalagem1.grid(row=10,column=1)

entry_nfembalagem2 = tk.Entry(aba2)
entry_nfembalagem2.grid(row=10,column=2)

entry_nfembalagem3 = tk.Entry(aba2)
entry_nfembalagem3.grid(row=10,column=3)

entry_nfembalagem4 = tk.Entry(aba2)
entry_nfembalagem4.grid(row=10,column=4)

entry_nfembalagem5 = tk.Entry(aba2)
entry_nfembalagem5.grid(row=10,column=5)

entry_nfembalagem6 = tk.Entry(aba2)
entry_nfembalagem6.grid(row=10,column=6)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=11)

label_nfpaleteEmb = tk.Label(aba2, text="NF PALLET")
label_nfpaleteEmb.grid(row=12,column=0)

entry_nfpaleteEmb1 = tk.Entry(aba2)
entry_nfpaleteEmb1.grid(row=12,column=1)

entry_nfpaleteEmb2 = tk.Entry(aba2)
entry_nfpaleteEmb2.grid(row=12,column=2)

entry_nfpaleteEmb3 = tk.Entry(aba2)
entry_nfpaleteEmb3.grid(row=12,column=3)

entry_nfpaleteEmb4 = tk.Entry(aba2)
entry_nfpaleteEmb4.grid(row=12,column=4)

entry_nfpaleteEmb5 = tk.Entry(aba2)
entry_nfpaleteEmb5.grid(row=12,column=5)

entry_nfpaleteEmb6 = tk.Entry(aba2)
entry_nfpaleteEmb6.grid(row=12,column=6)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=13)

label_codprod = tk.Label(aba2, text="COD. PROD")
label_codprod.grid(row=14,column=0)

entry_codprod1 = tk.Entry(aba2)
entry_codprod1.grid(row=14,column=1)

entry_codprod2 = tk.Entry(aba2)
entry_codprod2.grid(row=14,column=2)

entry_codprod3 = tk.Entry(aba2)
entry_codprod3.grid(row=14,column=3)

entry_codprod4 = tk.Entry(aba2)
entry_codprod4.grid(row=14,column=4)

entry_codprod5 = tk.Entry(aba2)
entry_codprod5.grid(row=14,column=5)

entry_codprod6 = tk.Entry(aba2)
entry_codprod6.grid(row=14,column=6)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=15)

label_qtdpalete = tk.Label(aba2, text="QTA. PALETE")
label_qtdpalete.grid(row=16,column=0)

entry_qtdpaleteEmb1 = tk.Entry(aba2)
entry_qtdpaleteEmb1.grid(row=16,column=1)

entry_qtdpaleteEmb2 = tk.Entry(aba2)
entry_qtdpaleteEmb2.grid(row=16,column=2)

entry_qtdpaleteEmb3 = tk.Entry(aba2)
entry_qtdpaleteEmb3.grid(row=16,column=3)

entry_qtdpaleteEmb4 = tk.Entry(aba2)
entry_qtdpaleteEmb4.grid(row=16,column=4)

entry_qtdpaleteEmb5 = tk.Entry(aba2)
entry_qtdpaleteEmb5.grid(row=16,column=5)

entry_qtdpaleteEmb6 = tk.Entry(aba2)
entry_qtdpaleteEmb6.grid(row=16,column=6)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=18)

label_val = tk.Label(aba2,text="DATA VALIDADE")
label_val.grid(row=19,column=0)

entry_valEmb1 = tk.Entry(aba2)
entry_valEmb1.grid(row=19,column=1)

entry_valEmb2 = tk.Entry(aba2)
entry_valEmb2.grid(row=19,column=2)

entry_valEmb3 = tk.Entry(aba2)
entry_valEmb3.grid(row=19,column=3)

entry_valEmb4 = tk.Entry(aba2)
entry_valEmb4.grid(row=19,column=4)

entry_valEmb5 = tk.Entry(aba2)
entry_valEmb5.grid(row=19,column=5)

entry_valEmb6 = tk.Entry(aba2)
entry_valEmb6.grid(row=19,column=6)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=20)

label_prod = tk.Label(aba2,text="NOME DO PROD")
label_prod.grid(row=21,column=0)

combobox_nomeprod1 = ttk.Combobox(aba2,values=[])
combobox_nomeprod1.grid(row=21,column=1)

combobox_nomeprod2 = ttk.Combobox(aba2,values=[])
combobox_nomeprod2.grid(row=21,column=2)

combobox_nomeprod3 = ttk.Combobox(aba2,values=[])
combobox_nomeprod3.grid(row=21,column=3)

combobox_nomeprod4 = ttk.Combobox(aba2,values=[])
combobox_nomeprod4.grid(row=21,column=4)

combobox_nomeprod5 = ttk.Combobox(aba2,values=[])
combobox_nomeprod5.grid(row=21,column=5)

combobox_nomeprod6 = ttk.Combobox(aba2,values=[])
combobox_nomeprod6.grid(row=21,column=6)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=22)

label_contUnid = tk.Label(aba2, text="QTD. UNIDADE")
label_contUnid.grid(row=23,column=0)

entry_contUnid1 = tk.Entry(aba2)
entry_contUnid1.grid(row=23,column=1)

entry_contUnid2 = tk.Entry(aba2)
entry_contUnid2.grid(row=23,column=2)

entry_contUnid3 = tk.Entry(aba2)
entry_contUnid3.grid(row=23,column=3)

entry_contUnid4 = tk.Entry(aba2)
entry_contUnid4.grid(row=23,column=4)

entry_contUnid5 = tk.Entry(aba2)
entry_contUnid5.grid(row=23,column=5)

entry_contUnid6 = tk.Entry(aba2)
entry_contUnid6.grid(row=23,column=6)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=24)

label_lotef = tk.Label(aba2,text="LOTE FORNECEDOR")
label_lotef.grid(row=25,column=0)

entry_lotef1 = tk.Entry(aba2)
entry_lotef1.grid(row=25,column=1)

entry_lotef2 = tk.Entry(aba2)
entry_lotef2.grid(row=25,column=2)

entry_lotef3 = tk.Entry(aba2)
entry_lotef3.grid(row=25,column=3)

entry_lotef4 = tk.Entry(aba2)
entry_lotef4.grid(row=25,column=4)

entry_lotef5 = tk.Entry(aba2)
entry_lotef5.grid(row=25,column=5)

entry_lotef6 = tk.Entry(aba2)
entry_lotef6.grid(row=25,column=6)

label_espaco = tk.Label(aba2)
label_espaco.grid(row=26)

label_peso = tk.Label(aba2,text="PESO")
label_peso.grid(row=27,column=0)

entry_pesoEmb1 = tk.Entry(aba2)
entry_pesoEmb1.grid(row=27,column=1)

entry_pesoEmb2 = tk.Entry(aba2)
entry_pesoEmb2.grid(row=27,column=2)

entry_pesoEmb3 = tk.Entry(aba2)
entry_pesoEmb3.grid(row=27,column=3)

entry_pesoEmb4 = tk.Entry(aba2)
entry_pesoEmb4.grid(row=27,column=4)

entry_pesoEmb5 = tk.Entry(aba2)
entry_pesoEmb5.grid(row=27,column=5)

entry_pesoEmb6 = tk.Entry(aba2)
entry_pesoEmb6.grid(row=27,column=6)

botao_salvar = tk.Button(aba2, text="Salvar", command=saveEmb)
botao_salvar.grid(row=28, columnspan=5)



# Inicia o loop principal da aplicação
janela.mainloop()