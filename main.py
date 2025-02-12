import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import win32com.client as win32
import win32print
import win32api
import os
import time
import openpyxl as ox
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Função para salvar dados na planilha
def salvar_dados():
    date = entry_date.get()
    horario = entry_horario.get()
    nome = entry_nome.get()
    telefone = entry_telefone.get()
    placa = entry_placa.get()
    tipo = entry_tipo.get()
    trans = entry_trans.get()
    forn = combobox_forn.get()
    prod = combobox_prod.get()
    carga = combobox_carga.get()
    val = entry_val.get()

    # Dados do lote 1
    nf1 = entry_nfsal1.get()
    nfpalete1 = entry_nfpalete1.get()
    qtd1 = int(entry_qtdpalete1.get())
    lote1 = entry_lotesal1.get()
    peso1 = int(entry_peso1.get())

    # Dados do lote 2
    if checkbox_lote2_var.get() == 1:
        nf2 = entry_nfsal2.get()
        nfpalete2 = entry_nfpalete2.get()
        qtd2 = int(entry_qtdpalete2.get())
        lote2 = entry_lotesal2.get()
        peso2 = int(entry_peso2.get())

    # Dados do lote 3
    if checkbox_lote3_var.get() == 1:
        nf3 = entry_nfsal3.get()
        nfpalete3 = entry_nfpalete3.get()
        qtd3 = int(entry_qtdpalete3.get())
        lote3 = entry_lotesal3.get()
        peso3 = int(entry_peso3.get())

    # Verifica se os campos estão preenchidos
    if date and horario and nome and telefone and placa and tipo and trans and forn and carga and prod and nf1 and lote1 and peso1:
        # Cria dicionário com os dados principais
        dados_programacao = {
            'Data de Entrada': [date], 
            'Horario de entrada': [horario], 
            'Nome do motorista': [nome],
            'Telefone': [telefone],
            'NF': [nf1],
            'Fornecedor': [forn],
            'Peso total': [peso1],
            'Placa': [placa],
            'Tipo de veiculo': [tipo],
            'Tipo do produto': [prod + " " + carga],
            'Transportadora': [trans]
        }

        # Converte o dicionário em DataFrame
        df_programacao = pd.DataFrame(dados_programacao)

        # Cria lista com dados do primeiro lote
        dados_planilha = [
            {'Movimento': "ENTRADA", 'EMISSÃO NF': date, 'Placa': placa, 'Tipo de veiculo': tipo, 'Transportadora': trans, 'Material': prod, 'Tipo de Carga': carga, 'Fornecedor': forn, 'NF fornecedor': nf1, 'NF Palete': nfpalete1, 'QT NF palete': qtd1, 'Lote do fornecedor': lote1, 'Validade': val, 'Peso': peso1}
        ]

        # Converte a lista em DataFrame
        df_planilha = pd.DataFrame(dados_planilha)

        # Verifica se o lote 2 deve ser incluído
        if checkbox_lote2_var.get() == 1:
            if nf2 and lote2 and peso2:
                df_planilha = pd.concat([df_planilha, pd.DataFrame([{'Movimento': "ENTRADA", 'EMISSÃO NF': date, 'Placa': placa, 'Tipo de veiculo': tipo, 'Transportadora': trans, 'Material': prod, 'Tipo de Carga': carga, 'Fornecedor': forn, 'NF fornecedor': nf2, 'NF Palete': nfpalete2, 'QT NF palete': qtd2, 'Lote do fornecedor': lote2, 'Validade': val, 'Peso': peso2}])], ignore_index=True)

        # Verifica se o lote 3 deve ser incluído
        if checkbox_lote3_var.get() == 1:
            if nf3 and lote3 and peso3:
                df_planilha = pd.concat([df_planilha, pd.DataFrame([{'Movimento': "ENTRADA", 'EMISSÃO NF': date, 'Placa': placa, 'Tipo de veiculo': tipo, 'Transportadora': trans, 'Material': prod, 'Tipo de Carga': carga, 'Fornecedor': forn, 'NF fornecedor': nf3, 'NF Palete': nfpalete3, 'QT NF palete': qtd3, 'Lote do fornecedor': lote3, 'Validade': val, 'Peso': peso3}])], ignore_index=True)

        # Verifica se o arquivo já existe
        if os.path.exists('dados.xlsx'):
            # Abre o arquivo existente e adiciona os dados
            wb = load_workbook('dados.xlsx')

            # Adiciona os dados à planilha de programação
            if 'Programacao' not in wb.sheetnames:
                ws_programacao = wb.create_sheet("Programacao")
            else:
                ws_programacao = wb['Programacao']

            for row in dataframe_to_rows(df_programacao, index=False, header=False):
                ws_programacao.append(row)

            # Adiciona os dados à planilha de lote
            if 'Planilha' not in wb.sheetnames:
                ws_planilha = wb.create_sheet("Planilha")
            else:
                ws_planilha = wb['Planilha']

            for row in dataframe_to_rows(df_planilha, index=False, header=False):
                ws_planilha.append(row)

            # Adiciona os dados à planilha de descarga de sal (exemplo)
            if 'Descarga do Sal' not in wb.sheetnames:
                ws_descarga_sal = wb.create_sheet("Descarga do Sal")
            else:
                ws_descarga_sal = wb['Descarga do Sal']
                ws_descarga_sal['D8'] = date
                ws_descarga_sal['K8'] = horario
                ws_descarga_sal['D10'] = nome
                ws_descarga_sal['D12'] = telefone
                ws_descarga_sal['D14'] = tipo
                ws_descarga_sal['K14'] = placa
                ws_descarga_sal['D16'] = trans
                ws_descarga_sal['D18'] = forn
                ws_descarga_sal['M18'] = carga
                ws_descarga_sal['D20'] = nf1
                ws_descarga_sal['K20'] = nfpalete1
                ws_descarga_sal['P20'] = peso1
                if checkbox_lote2_var.get() == 1:
                    if nf1 == nf2 :
                        ws_descarga_sal['D22'] = " "
                        ws_descarga_sal['K22'] = " "
                        ws_descarga_sal['P22'] = " "
                        ws_descarga_sal['P20'] = peso1+peso2
                    else :
                        ws_descarga_sal['D22'] = nf2
                        ws_descarga_sal['K22'] = nfpalete2
                        ws_descarga_sal['P22'] = peso2
                else :
                    ws_descarga_sal['D22'] = " "
                    ws_descarga_sal['K22'] = " "
                    ws_descarga_sal['P22'] = " "
                if checkbox_lote3_var.get() == 1:
                    if nf1 == nf3 and nf1 == nf2:
                        ws_descarga_sal['D24'] = " "
                        ws_descarga_sal['K24'] = " "
                        ws_descarga_sal['P24'] = " "
                        ws_descarga_sal['P20'] = peso1+peso2+peso3
                    else:
                        ws_descarga_sal['D24'] = nf3
                        ws_descarga_sal['K24'] = nfpalete3
                        ws_descarga_sal['P24'] = peso3
                else :
                    ws_descarga_sal['D24'] = " "
                    ws_descarga_sal['K24'] = " "
                    ws_descarga_sal['P24'] = " "
                ws_descarga_sal['D26'] = prod
                ws_descarga_sal['L26'] = val
                ws_descarga_sal['D28'] = lote1
                ws_descarga_sal['H28'] = nf1
                ws_descarga_sal['K28'] = peso1
                ws_descarga_sal['O28'] = qtd1
                if checkbox_lote2_var.get() == 1:
                    ws_descarga_sal['D30'] = lote2
                    ws_descarga_sal['H30'] = nf2
                    ws_descarga_sal['K30'] = peso2
                    ws_descarga_sal['O30'] = qtd2
                else :
                    ws_descarga_sal['D30'] = " "
                    ws_descarga_sal['H30'] = " "
                    ws_descarga_sal['K30'] = " "
                    ws_descarga_sal['O30'] = " "
                if checkbox_lote3_var.get() == 1:
                    ws_descarga_sal['D32'] = lote3
                    ws_descarga_sal['H32'] = nf3
                    ws_descarga_sal['K32'] = peso3
                    ws_descarga_sal['O32'] = qtd3
                else :
                    ws_descarga_sal['D32'] = " "
                    ws_descarga_sal['H32'] = " "
                    ws_descarga_sal['K32'] = " "
                    ws_descarga_sal['O32'] = " "


        else:
            # Se o arquivo não existir, cria um novo
            wb = Workbook()

            # Cria a planilha de Programacao
            ws_programacao = wb.active
            ws_programacao.title = "Programacao"
            for row in dataframe_to_rows(df_programacao, index=False, header=True):
                ws_programacao.append(row)

            # Cria a planilha de Planilha
            ws_planilha = wb.create_sheet("Planilha")
            for row in dataframe_to_rows(df_planilha, index=False, header=True):
                ws_planilha.append(row)

            # Cria a planilha de Descarga do Sal
            ws_descarga_sal = wb.create_sheet("Descarga do Sal")
            ws_descarga_sal = wb['Descarga do Sal']
            ws_descarga_sal['D8'] = date
            ws_descarga_sal['K8'] = horario
            ws_descarga_sal['D10'] = nome
            ws_descarga_sal['D12'] = telefone
            ws_descarga_sal['D14'] = tipo
            ws_descarga_sal['K14'] = placa
            ws_descarga_sal['D16'] = trans
            ws_descarga_sal['D18'] = forn
            ws_descarga_sal['M18'] = carga
            ws_descarga_sal['D20'] = nf1
            ws_descarga_sal['K20'] = nfpalete1
            ws_descarga_sal['P20'] = peso1
            ws_descarga_sal['D20'] = nf1
            ws_descarga_sal['K20'] = nfpalete1
            ws_descarga_sal['P20'] = peso1
            if checkbox_lote2_var.get() == 1:
                if nf1 == nf2 :
                    ws_descarga_sal['D22'] = " "
                    ws_descarga_sal['K22'] = " "
                    ws_descarga_sal['P22'] = " "
                    ws_descarga_sal['P20'] = peso1+peso2
                else :
                    ws_descarga_sal['D22'] = nf2
                    ws_descarga_sal['K22'] = nfpalete2
                    ws_descarga_sal['P22'] = peso2
            else :
                ws_descarga_sal['D22'] = " "
                ws_descarga_sal['K22'] = " "
                ws_descarga_sal['P22'] = " "
            if checkbox_lote3_var.get() == 1:
                if nf1 == nf3 and nf1 == nf2:
                    ws_descarga_sal['D24'] = " "
                    ws_descarga_sal['K24'] = " "
                    ws_descarga_sal['P24'] = " "
                    ws_descarga_sal['P20'] = peso1+peso2+peso3
                else:
                    ws_descarga_sal['D24'] = nf3
                    ws_descarga_sal['K24'] = nfpalete3
                    ws_descarga_sal['P24'] = peso3
            else :
                ws_descarga_sal['D24'] = " "
                ws_descarga_sal['K24'] = " "
                ws_descarga_sal['P24'] = " "
            ws_descarga_sal['D26'] = prod
            ws_descarga_sal['L26'] = val
            ws_descarga_sal['D28'] = lote1
            ws_descarga_sal['H28'] = nf1
            ws_descarga_sal['K28'] = peso1
            ws_descarga_sal['O28'] = qtd1
            if checkbox_lote2_var.get() == 1:
                ws_descarga_sal['D30'] = lote2
                ws_descarga_sal['H30'] = nf2
                ws_descarga_sal['K30'] = peso2
                ws_descarga_sal['O30'] = qtd2
            else :
                ws_descarga_sal['D30'] = " "
                ws_descarga_sal['H30'] = " "
                ws_descarga_sal['K30'] = " "
                ws_descarga_sal['O30'] = " "
            if checkbox_lote3_var.get() == 1:
                ws_descarga_sal['D32'] = lote3
                ws_descarga_sal['H32'] = nf3
                ws_descarga_sal['K32'] = peso3
                ws_descarga_sal['O32'] = qtd3
            else :
                ws_descarga_sal['D32'] = " "
                ws_descarga_sal['H32'] = " "
                ws_descarga_sal['K32'] = " "
                ws_descarga_sal['O32'] = " "
        

        # Salva a planilha
        wb.save('dados.xlsx')
        txtprogramação()
        messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")

    else:
        messagebox.showerror("Erro", "Preencha todos os campos obrigatórios.")

def imprimirespelho():

    # Obtém o caminho da pasta local onde o script está sendo executado
    caminho_local = os.path.dirname(os.path.abspath(__file__))

    # Substitua 'dados.xlsx' pelo nome do seu arquivo Excel
    arquivo_excel = os.path.join(caminho_local, 'dados.xlsx')

    # Substitua 'Nome_da_Tabela' pelo nome da tabela que você deseja ler
    nome_tabela = 'Descarga do Sal'

    # Carrega a planilha
    workbook = ox.load_workbook(arquivo_excel)
    worksheet = workbook[nome_tabela]

    # Verifica se a tabela foi carregada corretamente
    print(f"Tabela carregada: {worksheet.title}")

    # Converte o arquivo Excel formatado em PDF usando win32com
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False

    wb = excel.Workbooks.Open(arquivo_excel)
    ws = wb.Worksheets(nome_tabela)

    # Salva o arquivo Excel para garantir que as configurações de impressão sejam aplicadas
    wb.Save()

    # Exporta a planilha como PDF
    pdf_path = os.path.join(caminho_local, 'tabela_formatada.pdf')
    ws.ExportAsFixedFormat(0, pdf_path)

    wb.Close()
    excel.Quit()

    print("PDF gerado com sucesso com configurações de impressão herdadas!")

    # Envia o PDF para a impressora
    printer_name = win32print.GetDefaultPrinter()
    win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)

    print(f"Enviando {pdf_path} para a impressora {printer_name}")

    # Aguarda 10 segundos antes de excluir o PDF
    time.sleep(10)

    # Exclui o PDF gerado
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
        print(f"PDF {pdf_path} excluído com sucesso!")

def txtprogramação():

    date = entry_date.get()
    horario = entry_horario.get()
    nome = entry_nome.get()
    telefone = entry_telefone.get()
    placa = entry_placa.get()
    tipo = entry_tipo.get()
    trans = entry_trans.get()
    forn = combobox_forn.get()
    prod = combobox_prod.get()
    carga = combobox_carga.get()

    # Dados do lote 1
    nf1 = entry_nfsal1.get()
    peso1 = int(entry_peso1.get())

    if checkbox_lote2_var.get() == 1:
        nf2 = entry_nfsal2.get()
        peso2 = int(entry_peso2.get())

    # Dados do lote 3
    if checkbox_lote3_var.get() == 1:
        nf3 = entry_nfsal3.get()
        peso3 = int(entry_peso3.get())
        # Abre (ou cria) o arquivo txt em modo de escrita (write mode)
    with open('meuarquivo.txt', 'w') as arquivo:
        # Escreve uma linha no arquivo
        if checkbox_lote2_var.get() == 0 and checkbox_lote3_var.get() == 0 :
            arquivo.write(f'{date} {horario} {nome} {telefone} {nf1} {forn} {peso1} {placa} {tipo} {prod}{carga} {trans}\n')
        if checkbox_lote2_var.get() == 1 and checkbox_lote3_var.get() == 0 :
            arquivo.write(f'{date} {horario} {nome} {telefone} {nf1}/{nf2} {forn} {peso1+peso2} {placa} {tipo} {prod}{carga} {trans}\n')
        if checkbox_lote2_var.get() == 1 and checkbox_lote3_var.get() == 1 :
            arquivo.write(f'{date} {horario} {nome} {telefone} {nf1}/{nf2}/{nf3} {forn} {peso1} {placa+peso2+peso3} {tipo} {prod}{carga} {trans}\n')

    print('Texto escrito com sucesso!')

    # Abre o arquivo txt automaticamente no editor de texto padrão do sistema
    os.startfile('meuarquivo.txt')



# Cria a janela principal
janela = tk.Tk()
janela.title("Cadastro de Informações")

# Cria os campos e labels
label_date = tk.Label(janela, text="Data de entrada:")
label_date.grid(row=0, column=0)
entry_date = tk.Entry(janela)
entry_date.grid(row=0, column=1)

label_horario = tk.Label(janela, text="Horário de chegada:")
label_horario.grid(row=1, column=0)
entry_horario = tk.Entry(janela)
entry_horario.grid(row=1, column=1)

label_nome = tk.Label(janela, text="Nome do motorista:")
label_nome.grid(row=2, column=0)
entry_nome = tk.Entry(janela)
entry_nome.grid(row=2, column=1)

label_telefone = tk.Label(janela, text="Telefone:")
label_telefone.grid(row=3, column=0)
entry_telefone = tk.Entry(janela)
entry_telefone.grid(row=3, column=1)

label_placa = tk.Label(janela, text="Placa do veículo:")
label_placa.grid(row=4, column=0)
entry_placa = tk.Entry(janela)
entry_placa.grid(row=4, column=1)

label_tipo = tk.Label(janela, text="Tipo de Veículo:")
label_tipo.grid(row=5, column=0)
entry_tipo = tk.Entry(janela)
entry_tipo.grid(row=5, column=1)

label_trans = tk.Label(janela, text="Transportadora:")
label_trans.grid(row=6, column=0)

entry_trans = tk.Entry(janela)
entry_trans.grid(row=6, column=1)

label_forn = tk.Label(janela, text="Fornecedor:")
label_forn.grid(row=7, column=0)

combobox_forn = ttk.Combobox(janela, values=["NORSAL", "CIMSAL", "CISNE"])
combobox_forn.grid(row=7, column=1)
combobox_forn.set("Selecione uma opção")

label_prod = tk.Label(janela, text="Produto:")
label_prod.grid(row=8, column=0)

combobox_prod = ttk.Combobox(janela, values=["SAL REFINADO", "SAL GRANULADO GROSSO", "SAL EXTRA REFINADO","SAL ENTREFINO"])
combobox_prod.grid(row=8, column=1)
combobox_prod.set("Selecione uma opção")

label_carga = tk.Label(janela, text="Tipo de Carga:")
label_carga.grid(row=9, column=0)

combobox_carga = ttk.Combobox(janela, values=["BIG BAG", "25 KG"])
combobox_carga.grid(row=9, column=1)
combobox_carga.set("Selecione uma opção")

label_val = tk.Label(janela, text="Validade:")
label_val.grid(row=10, column=0)

entry_val = tk.Entry(janela)
entry_val.grid(row=10, column=1)

label_nfsal1 = tk.Label(janela, text="NF do SAL:")
label_nfsal1.grid(row=11, column=0)

entry_nfsal1 = tk.Entry(janela)
entry_nfsal1.grid(row=11, column=1)

label_lotesal1 = tk.Label(janela, text="Lote:")
label_lotesal1.grid(row=11, column=2)

entry_lotesal1 = tk.Entry(janela)
entry_lotesal1.grid(row=11, column=3)

label_nfpalete1 = tk.Label(janela, text="NF DO Palete:")
label_nfpalete1.grid(row=11, column=4)

entry_nfpalete1 = tk.Entry(janela)
entry_nfpalete1.grid(row=11, column=5)

label_qtdpalete1 = tk.Label(janela, text="QTD Palete:")
label_qtdpalete1.grid(row=11, column=6)

entry_qtdpalete1 = tk.Entry(janela)
entry_qtdpalete1.grid(row=11, column=7)

label_peso1 = tk.Label(janela, text="peso:")
label_peso1.grid(row=11, column=8)

entry_peso1 = tk.Entry(janela)
entry_peso1.grid(row=11, column=9)

checkbox_lote2_var = tk.IntVar()
checkbox_lote2 = tk.Checkbutton(janela, text="Incluir Lote 2", variable=checkbox_lote2_var)
checkbox_lote2.grid(row=12, column=0)

label_nfsal2 = tk.Label(janela, text="NF do SAL:")
label_nfsal2.grid(row=13, column=0)

entry_nfsal2 = tk.Entry(janela)
entry_nfsal2.grid(row=13, column=1)

label_lotesal2 = tk.Label(janela, text="Lote:")
label_lotesal2.grid(row=13, column=2)

entry_lotesal2 = tk.Entry(janela)
entry_lotesal2.grid(row=13, column=3)

label_nfpalete2 = tk.Label(janela, text="NF DO Palete:")
label_nfpalete2.grid(row=13, column=4)

entry_nfpalete2 = tk.Entry(janela)
entry_nfpalete2.grid(row=13, column=5)

label_qtdpalete2 = tk.Label(janela, text="QTD Palete:")
label_qtdpalete2.grid(row=13, column=6)

entry_qtdpalete2 = tk.Entry(janela)
entry_qtdpalete2.grid(row=13, column=7)

label_peso2 = tk.Label(janela, text="peso:")
label_peso2.grid(row=13, column=8)

entry_peso2 = tk.Entry(janela)
entry_peso2.grid(row=13, column=9)

checkbox_lote3_var = tk.IntVar()
checkbox_lote3 = tk.Checkbutton(janela, text="Incluir Lote 3", variable=checkbox_lote3_var)
checkbox_lote3.grid(row=14, column=0)

label_nfsal3 = tk.Label(janela, text="NF do SAL:")
label_nfsal3.grid(row=15, column=0)

entry_nfsal3 = tk.Entry(janela)
entry_nfsal3.grid(row=15, column=1)

label_lotesal3 = tk.Label(janela, text="Lote:")
label_lotesal3.grid(row=15, column=2)

entry_lotesal3 = tk.Entry(janela)
entry_lotesal3.grid(row=15, column=3)

label_nfpalete3 = tk.Label(janela, text="NF DO Palete:")
label_nfpalete3.grid(row=15, column=4)

entry_nfpalete3 = tk.Entry(janela)
entry_nfpalete3.grid(row=15, column=5)

label_qtdpalete3 = tk.Label(janela, text="QTD Palete:")
label_qtdpalete3.grid(row=15, column=6)

entry_qtdpalete3 = tk.Entry(janela)
entry_qtdpalete3.grid(row=15, column=7)

label_peso3 = tk.Label(janela, text="peso:")
label_peso3.grid(row=15, column=8)

entry_peso3 = tk.Entry(janela)
entry_peso3.grid(row=15, column=9)


# Cria o botão para salvar
botao_salvar = tk.Button(janela, text="Salvar", command=salvar_dados)
botao_salvar.grid(row=16, columnspan=5)

botao_imprimir = tk.Button(janela,text="Imprimir Espelho",command=imprimirespelho)
botao_imprimir.grid(row=16, columnspan=9)

# Inicia a aplicação
janela.mainloop()
