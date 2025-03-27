import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from print import imprimirespelho
from programacao import txtprogramação as txt
from salvar_dados import salvar_dados as salvar

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
        messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
    except ValueError as e:
        messagebox.showerror("Erro de valor", "Nos campos de QTD e Peso devese colocar exclusivamente numeros")
    except PermissionError as e:
        messagebox.showerror("Erro de permissão", f"Permissão negada: {e}. Verifique se o arquivo está aberto em outro programa.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def prog():
    try:
        txt(entry_date.get(),entry_horario.get(),entry_nome.get(),entry_telefone.get(),entry_placa.get(),entry_tipo.get(),entry_trans.get(),combobox_forn.get(),combobox_prod.get(),combobox_carga.get(),entry_nfsal1.get(),int(entry_peso1.get() or "0"),entry_nfsal2.get(),int(entry_peso2.get() or "0"),entry_nfsal3.get(),int(entry_peso3.get() or "0"),checkbox_lote2_var.get(),checkbox_lote3_var.get())
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
aba1 = tk.Frame(notebook)
notebook.add(aba1, text="Cadastro Principal")

# Cria o frame para a segunda aba
aba2 = tk.Frame(notebook)
notebook.add(aba2, text="Outra Aba")

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
label_outra = tk.Label(aba2, text="Novo Campo na Segunda Aba:")
label_outra.grid(row=0, column=0)
entry_outra = tk.Entry(aba2)
entry_outra.grid(row=0, column=1)

botao_exemplo = tk.Button(aba2, text="Botão da Segunda Aba")
botao_exemplo.grid(row=1, column=0)

# Inicia o loop principal da aplicação
janela.mainloop()