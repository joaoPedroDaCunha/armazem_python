import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def salvar_dados(date,horario,nome,telefone,placa,tipo,trans,forn,prod,carga,val,nf1,nfpalete1,qtd1,lote1,peso1,checkbox_lote2_var,nf2,nfpalete2,qtd2,lote2,peso2,checkbox_lote3_var,nf3,nfpalete3,qtd3,lote3,peso3):
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
            print("Sucesso", "Dados salvos com sucesso!")

        else:
            print("Erro", "Preencha todos os campos obrigatórios.")
