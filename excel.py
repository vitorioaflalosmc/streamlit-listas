import json
import xlwings as xw
import streamlit as st
import time

def preencher_excel_com_json(json_filename, novo_excel_filename):
    template1 = 'LISTAS/GEP/AREA_MEIO/lista-online.xlsx'
    template2 = 'LISTAS/GEP/AREA_MEIO/lista-presencial.xlsx'
    template3 = 'LISTAS/GEP/EMESP/lista-presencial.xlsx'
    template4 = 'LISTAS/GEP/EMESP/lista-online.xlsx'
    template5 = 'LISTAS/GEP/GURI/lista-online.xlsx'
    template6 = 'LISTAS/GEP/GURI/lista-presencial.xlsx'
    template7 = 'LISTAS/PED-GURI/lista-grupo-polos.xlsx'
    template8 = 'LISTAS/PED-GURI/lista-pedagogico.xlsx'
    template9 = 'LISTAS/SOCIAL/lista-guri.xlsx'
    template_10 = 'LISTAS/SOCIAL/lista-emesp.xlsx'
    sucesso = False


    try:
        # Carregar o JSON salvo
        with open(json_filename, 'r') as f:
            dados_json = json.load(f)

        if json_filename.startswith("lista1"):
            with xw.App(visible=True) as app:
                wb = app.books.open(template1)  # Abre o template uma única vez
                ws = wb.sheets.active

                    # Preencher as células com os dados do JSON
                ws.range('B4').value = dados_json.get("Tema", "")
                ws.range('B5').value = dados_json.get("Palestrante", "")
                ws.range('B6').value = dados_json.get("Publico Alvo", "")
                ws.range('B7').value = dados_json.get("Data", "")
                ws.range('B8').value = dados_json.get("Horario de Inicio", "")
                ws.range('D8').value = dados_json.get("Horario de Fim", "")
                ws.range('B9').value = dados_json.get("Carga Horaria", "")
                ws.range('B10').value = dados_json.get("Plataforma Online", "")
                ws.range('B11').value = dados_json.get("Contrato de Gestao", "")

                # Definir formatação: tamanho 20, fonte Calibri, cor preta, sem negrito
                cell_ranges = ['B4', 'B5', 'B6', 'B7', 'B8', 'D8', 'B9', 'B10', 'B11']
                for cell in cell_ranges:
                    cell_range = ws.range(cell)
                    cell_range.api.Font.Size = 20
                    cell_range.api.Font.Name = "Calibri"
                    cell_range.api.Font.Bold = True
                    cell_range.api.Font.Color = 0x000000  # Cor preta
                    # Alinhar à esquerda
                    cell_range.api.HorizontalAlignment = -4131  # Constante para xlLeft

                ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address  # Ajuste o range conforme necessário

                # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
                wb.save(novo_excel_filename)

                wb.close()  # Fechar o workbook após salvar
                app.quit()  # Fechar o aplicativo Excel completamente


        elif json_filename.startswith("lista2"):
                        # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template2)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('B4').value = dados_json.get("Tema", "")
            ws.range('B5').value = dados_json.get("Palestrante", "")
            ws.range('B6').value = dados_json.get("Publico Alvo", "")
            ws.range('B7').value = dados_json.get("Data", "")
            ws.range('B8').value = dados_json.get("Horario de Inicio", "")
            ws.range('D8').value = dados_json.get("Horario de Fim", "")
            ws.range('B9').value = dados_json.get("Carga Horaria", "")
            ws.range('B10').value = dados_json.get("Local", "")
            ws.range('B11').value = dados_json.get("Contrato de Gestao", "")
            
            # Definir formatação: tamanho 20, fonte Calibri, cor preta, sem negrito
            cell_ranges = ['B4', 'B5', 'B6', 'B7', 'B8', 'D8', 'B9', 'B10', 'B11']
            for cell in cell_ranges:
                cell_range = ws.range(cell)
                cell_range.api.Font.Size = 20
                cell_range.api.Font.Name = "Calibri"
                cell_range.api.Font.Bold = True
                cell_range.api.Font.Color = 0x000000  # Cor preta (0x000000 em hexadecimal)
                # Alinhar à esquerda
                cell_range.api.HorizontalAlignment = -4131  # Constante para xlLeft
                ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address  # Ajuste o range conforme o necessário
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")

        elif json_filename.startswith("lista3"):
                        # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template3)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('B4').value = dados_json.get("Tema", "")
            ws.range('B5').value = dados_json.get("Palestrante", "")
            ws.range('B6').value = dados_json.get("Publico Alvo", "")
            ws.range('B7').value = dados_json.get("Data", "")
            ws.range('B8').value = dados_json.get("Horario de Inicio", "")
            ws.range('D8').value = dados_json.get("Horario de Fim", "")
            ws.range('B9').value = dados_json.get("Carga Horaria", "")
            ws.range('B10').value = dados_json.get("Local", "")
            ws.range('B11').value = dados_json.get("Contrato de Gestao", "")
            
            # Definir formatação: tamanho 20, fonte Calibri, cor preta, sem negrito
            cell_ranges = ['B4', 'B5', 'B6', 'B7', 'B8', 'D8', 'B9', 'B10', 'B11']
            for cell in cell_ranges:
                cell_range = ws.range(cell)
                cell_range.api.Font.Size = 14
                cell_range.api.Font.Name = "Calibri"
                cell_range.api.Font.Bold = True
                cell_range.api.Font.Color = 0x000000  # Cor preta (0x000000 em hexadecimal)
                # Alinhar à esquerda
                cell_range.api.HorizontalAlignment = -4131  # Constante para xlLeft
                ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address  # Ajuste o range conforme o necessário
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")

        elif json_filename.startswith("lista4"):
                        # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template4)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('B4').value = dados_json.get("Tema", "")
            ws.range('B5').value = dados_json.get("Palestrante", "")
            ws.range('B6').value = dados_json.get("Publico Alvo", "")
            ws.range('B7').value = dados_json.get("Data", "")
            ws.range('B8').value = dados_json.get("Horario de Inicio", "")
            ws.range('D8').value = dados_json.get("Horario de Fim", "")
            ws.range('B9').value = dados_json.get("Carga Horaria", "")
            ws.range('B10').value = dados_json.get("Plataforma Online", "")
            ws.range('B11').value = dados_json.get("Contrato de Gestao", "")
            
            # Definir formatação: tamanho 20, fonte Calibri, cor preta, sem negrito
            cell_ranges = ['B4', 'B5', 'B6', 'B7', 'B8', 'D8', 'B9', 'B10', 'B11']
            for cell in cell_ranges:
                cell_range = ws.range(cell)
                cell_range.api.Font.Size = 14
                cell_range.api.Font.Name = "Calibri"
                cell_range.api.Font.Bold = True
                cell_range.api.Font.Color = 0x000000  # Cor preta (0x000000 em hexadecimal)
                # Alinhar à esquerda
                cell_range.api.HorizontalAlignment = -4131  # Constante para xlLeft
            ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address  # Ajuste o range conforme o necessário
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")
        
        elif json_filename.startswith("lista5"):
                        # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template5)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('B4').value = dados_json.get("Tema", "")
            ws.range('B5').value = dados_json.get("Palestrante", "")
            ws.range('B6').value = dados_json.get("Publico Alvo", "")
            ws.range('B7').value = dados_json.get("Data", "")
            ws.range('B8').value = dados_json.get("Horario de Inicio", "")
            ws.range('D8').value = dados_json.get("Horario de Fim", "")
            ws.range('B9').value = dados_json.get("Carga Horaria", "")
            ws.range('B10').value = dados_json.get("Plataforma Online", "")
            ws.range('B11').value = dados_json.get("Contrato de Gestao", "")
            
            # Definir formatação: tamanho 20, fonte Calibri, cor preta, sem negrito
            cell_ranges = ['B4', 'B5', 'B6', 'B7', 'B8', 'D8', 'B9', 'B10', 'B11']
            for cell in cell_ranges:
                cell_range = ws.range(cell)
                cell_range.api.Font.Size = 14
                cell_range.api.Font.Name = "Calibri"
                cell_range.api.Font.Bold = True
                cell_range.api.Font.Color = 0x000000  # Cor preta (0x000000 em hexadecimal)
                # Alinhar à esquerda
                cell_range.api.HorizontalAlignment = -4131  # Constante para xlLeft
            ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address  # Ajuste o range conforme o necessário
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")

        elif json_filename.startswith("lista6"):
                        # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template6)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('B4').value = dados_json.get("Tema", "")
            ws.range('B5').value = dados_json.get("Palestrante", "")
            ws.range('B6').value = dados_json.get("Publico Alvo", "")
            ws.range('B7').value = dados_json.get("Data", "")
            ws.range('B8').value = dados_json.get("Horario de Inicio", "")
            ws.range('D8').value = dados_json.get("Horario de Fim", "")
            ws.range('B9').value = dados_json.get("Carga Horaria", "")
            ws.range('B10').value = dados_json.get("Local", "")
            ws.range('B11').value = dados_json.get("Contrato de Gestao", "")
            
            # Definir formatação: tamanho 20, fonte Calibri, cor preta, sem negrito
            cell_ranges = ['B4', 'B5', 'B6', 'B7', 'B8', 'D8', 'B9', 'B10', 'B11']
            for cell in cell_ranges:
                cell_range = ws.range(cell)
                cell_range.api.Font.Size = 14
                cell_range.api.Font.Name = "Calibri"
                cell_range.api.Font.Bold = True
                cell_range.api.Font.Color = 0x000000  # Cor preta (0x000000 em hexadecimal)
                # Alinhar à esquerda
                cell_range.api.HorizontalAlignment = -4131  # Constante para xlLeft
            ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address  # Ajuste o range conforme o necessário
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")

        elif json_filename.startswith("lista7"):
                        # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template7)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('A1').value = dados_json.get("Tema", "")
            ws.range('C2').value = dados_json.get("Programa", "")
            ws.range('C3').value = dados_json.get("Grupo", "")
            ws.range('C4').value = dados_json.get("Nome da Atividade", "")
            ws.range('C5').value = dados_json.get("Polos Participantes", "")
            ws.range('C6').value = dados_json.get("Local", "")
            ws.range('C7').value = dados_json.get("Data", "")
            ws.range('C8').value = dados_json.get("Horario de Inicio", "")
            ws.range('E8').value = dados_json.get("Horario de Fim", "")
            ws.range('C9').value = dados_json.get("Professor", "")
            
            # Definir formatação: tamanho 20, fonte Calibri, cor preta, sem negrito
            cell_ranges = ['A1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'E8', 'C9']
            for cell in cell_ranges:
                cell_range = ws.range(cell)
                cell_range.api.Font.Size = 19
                cell_range.api.Font.Name = "Calibri"
                cell_range.api.Font.Bold = True
                cell_range.api.Font.Color = 0x000000  # Cor preta (0x000000 em hexadecimal)
            # Centralizar apenas a célula A1, alinhar as outras à esquerda
            if cell == 'A1':
                cell_range.api.HorizontalAlignment = -4108  # Constante para xlCenter
            else:
                cell_range.api.HorizontalAlignment = -4131  # Constante para xlLeft
            ws.api.PageSetup.PrintArea = ws.range('A1:E1000').api.Address  # Ajuste o range conforme o necessário
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")

        elif json_filename.startswith("lista8"):
                        # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template8)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('A1').value = dados_json.get("Tema", "")
            ws.range('C2').value = dados_json.get("Programa", "")
            ws.range('C3').value = dados_json.get("Nome da Atividade", "")
            ws.range('C4').value = dados_json.get("Polos Participantes", "")
            ws.range('C5').value = dados_json.get("Local", "")
            ws.range('C6').value = dados_json.get("Data", "")
            ws.range('C7').value = dados_json.get("Horario de Inicio", "")
            ws.range('E7').value = dados_json.get("Horario de Fim", "")
            ws.range('C8').value = dados_json.get("Professor", "")
            
            # Definir formatação: tamanho 20, fonte Calibri, cor preta, sem negrito
            cell_ranges = ['A1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'E7']
            for cell in cell_ranges:
                cell_range = ws.range(cell)
                cell_range.api.Font.Size = 19
                cell_range.api.Font.Name = "Calibri"
                cell_range.api.Font.Bold = True
                cell_range.api.Font.Color = 0x000000  # Cor preta (0x000000 em hexadecimal)
            # Centralizar apenas a célula A1, alinhar as outras à esquerda
            if cell == 'A1':
                cell_range.api.HorizontalAlignment = -4108  # Constante para xlCenter
            else:
                cell_range.api.HorizontalAlignment = -4131  # Constante para xlLeft
            ws.api.PageSetup.PrintArea = ws.range('A1:E1000').api.Address  # Ajuste o range conforme o necessário
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")
    
        elif json_filename.startswith("lista9"):
                        # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template9)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('A1').value = dados_json.get("Meta", "")
            ws.range('B3').value = dados_json.get("Tema", "")
            ws.range('B4').value = dados_json.get("Convidado", "")
            ws.range('B5').value = dados_json.get("Regional", "")
            ws.range('B6').value = dados_json.get("Polo", "")
            ws.range('B7').value = dados_json.get("Publico Alvo", "")
            ws.range('B8').value = dados_json.get("Data", "")
            ws.range('B9').value = dados_json.get("Horario de Inicio", "")
            ws.range('D9').value = dados_json.get("Horario de Fim", "")
            ws.range('D4').value = dados_json.get("Responsável pela Atividade", "")
            ws.range('D5').value = dados_json.get("Local", "")
            ws.range('D6').value = dados_json.get("Remoto ou Presencial", "")
            ws.range('D8').value = dados_json.get("Carga Horaria", "")
            
            cell_ranges = ['A1', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'D4', 'D5', 'D6', 'D8', 'D9']
            xlCenter = -4108  # Alinhamento centralizado (horizontal)
            xlLeft = -4131    # Alinhamento à esquerda
            xlCenterVertical = -4108  # Alinhamento vertical centralizado

            for cell in cell_ranges:
                cell_range = ws.range(cell)
                
                # Definir o alinhamento padrão para o meio (horizontal e vertical)
                cell_range.api.HorizontalAlignment = xlCenter
                cell_range.api.VerticalAlignment = xlCenterVertical
                
                if cell == 'A1':
                    # Configurações específicas para a célula A1
                    cell_range.api.Font.Size = 32
                else:
                    # Configurações para as outras células
                    cell_range.api.Font.Size = 20
                    cell_range.api.Font.Name = "Calibri"
                    cell_range.api.Font.Bold = True
                    cell_range.api.Font.Color = 0x000000  # Cor preta

                    # Se a célula não for 'A1' e for 'B7' (célula mesclada), aplicar alinhamento à esquerda
                    if cell == 'B7':
                        ws.range('B7:D7').api.HorizontalAlignment = xlLeft
                    else:
                        # Alinhar à esquerda todas as células, exceto 'A1' e células mescladas como 'B7'
                        cell_range.api.HorizontalAlignment = xlLeft

            # Definir área de impressão (ajuste o range conforme necessário)
            ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address

            # Definir área de impressão (ajuste o range conforme necessário)
            ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")

        elif json_filename == "10lista.json":
            # Abrir o template Excel usando xlwings
            app = xw.App(visible=False)
            wb = xw.Book(template_10)
            ws = wb.sheets.active  # Usar a primeira planilha ativa
            
            # Preencher as células com os dados do JSON
            ws.range('A1').value = dados_json.get("Meta", "")
            ws.range('B3').value = dados_json.get("Tema", "")
            ws.range('B4').value = dados_json.get("Convidado", "")
            ws.range('B5').value = dados_json.get("Público Alvo", "")
            ws.range('B6').value = dados_json.get("Local", "")
            ws.range('B7').value = dados_json.get("Data", "")
            ws.range('B8').value = dados_json.get("Horario de Inicio", "")
            ws.range('D4').value = dados_json.get("Responsavel pela Atividade", "")
            ws.range('D5').value = dados_json.get("Parceiros", "")
            ws.range('D6').value = dados_json.get("Remoto ou Presencial", "")
            ws.range('D7').value = dados_json.get("Carga Horaria", "")
            ws.range('D8').value = dados_json.get("Horario de Fim", "")
            
            cell_ranges = ['A1', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'D4', 'D5', 'D6', 'D8', 'D7']
            xlCenter = -4108  # Alinhamento centralizado (horizontal)
            xlLeft = -4131    # Alinhamento à esquerda
            xlCenterVertical = -4108  # Alinhamento vertical centralizado

            for cell in cell_ranges:
                cell_range = ws.range(cell)
                
                # Definir o alinhamento padrão para o meio (horizontal e vertical)
                cell_range.api.HorizontalAlignment = xlCenter
                cell_range.api.VerticalAlignment = xlCenterVertical
                
                if cell == 'A1':
                    # Configurações específicas para a célula A1
                    cell_range.api.Font.Size = 32
                else:
                    # Configurações para as outras células
                    cell_range.api.Font.Size = 20
                    cell_range.api.Font.Name = "Calibri"
                    cell_range.api.Font.Bold = True
                    cell_range.api.Font.Color = 0x000000  # Cor preta
                    cell_range.api.HorizontalAlignment = xlLeft

            # Definir área de impressão (ajuste o range conforme necessário)
            ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address

            # Definir área de impressão (ajuste o range conforme necessário)
            ws.api.PageSetup.PrintArea = ws.range('A1:D1000').api.Address
            # Salvar o novo arquivo Excel, mantendo o cabeçalho e rodapé
            wb.save(novo_excel_filename)
            wb.close()  # Fechar o workbook após salvar
            app.quit()  # Fechar o aplicativo Excel completamente

            st.success(f"Novo arquivo Excel salvo como: {novo_excel_filename}")

        else:
            st.error("Houve erro no preenchimento. O json não foi achado.")
        
    
    except Exception as e:
        st.error(f'Houve erro no Excel: {e}')