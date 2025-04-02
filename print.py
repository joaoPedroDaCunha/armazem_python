import os
from threading import Thread
import time
import win32com.client as win32
import openpyxl as ox
import win32print
import win32api
from tkinter import messagebox
import sys
import pythoncom

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def imprimirespelho():
    try:
        Thread1 = Thread(target=imprimir)
        Thread1.daemon
        Thread1.start()
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def imprimirespelhoEmb():
    try:
        Thread1 = Thread(target=imprirmirEmb)
        Thread1.daemon
        Thread1.start()
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def imprimir():
    pythoncom.CoInitialize()
    try:
        caminho_local = resource_path('dados.xlsx')
        arquivo_excel = os.path.join(caminho_local)

        # O arquivo PDF será salvo no mesmo diretório do arquivo Excel
        arquivo_pdf = os.path.join(os.path.dirname(os.path.abspath(arquivo_excel)), 'tabela_formatada.pdf')

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

        # Remove as margens
        ws.PageSetup.LeftMargin = 0
        ws.PageSetup.RightMargin = 0
        ws.PageSetup.TopMargin = 0
        ws.PageSetup.BottomMargin = 0

        # Ajusta a escala para 95%
        ws.PageSetup.Zoom = 95  # Define a escala para 95%

        xlPortrait = 1
        ws.PageSetup.Orientation = xlPortrait


        # Salva o arquivo Excel para garantir que as configurações de impressão sejam aplicadas
        wb.Save()

        # Exporta a planilha como PDF
        pdf_path = os.path.join(os.path.dirname(os.path.abspath(arquivo_excel)), 'tabela_formatada.pdf')
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
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    finally:
        pythoncom.CoUninitialize()
    

def imprirmirEmb():
    pythoncom.CoInitialize()
    try:
        caminho_local = resource_path('dados.xlsx')
        arquivo_excel = os.path.join(caminho_local)
        arquivo_pdf = os.path.join(os.path.dirname(os.path.abspath(arquivo_excel)), 'tabela_formatada_emb.pdf')
        nome_tabela = 'Descarga Embalagem'
        workbook = ox.load_workbook(arquivo_excel)
        worksheet = workbook[nome_tabela]
        print(f"Tabela carregada: {worksheet.title}")
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        wb = excel.Workbooks.Open(arquivo_excel)
        ws = wb.Worksheets(nome_tabela)
        ws.PageSetup.LeftMargin = 1.5
        ws.PageSetup.RightMargin = 0
        ws.PageSetup.TopMargin = 0
        ws.PageSetup.BottomMargin = 0
        ws.PageSetup.Zoom = 90
        xlLandscape = 2
        ws.PageSetup.Orientation = xlLandscape
        wb.Save()
        pdf_path = os.path.join(os.path.dirname(os.path.abspath(arquivo_excel)), 'tabela_formatada_emb.pdf')
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
    
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    finally:
        pythoncom.CoUninitialize()