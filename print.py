import openpyxl as ox
import win32com.client as win32
import win32print
import win32api
import os
import time

# Substitua 'nome_do_arquivo.xlsx' pelo caminho do seu arquivo Excel
arquivo_excel = 'c:\\Leitor XML\\dados.xlsx'

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

wb = excel.Workbooks.Open(r'{}'.format(arquivo_excel))
ws = wb.Worksheets(nome_tabela)

# Salva o arquivo Excel para garantir que as configurações de impressão sejam aplicadas
wb.Save()

# Exporta a planilha como PDF
pdf_path = r'c:\\Leitor XML\\tabela_formatada.pdf'
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
