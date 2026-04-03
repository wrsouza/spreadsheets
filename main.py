import pandas as pd

# Nome do arquivo
file_name = 'Planilha_Sorteio_Estatistico_V2.xlsx'

# 1. Configuração da Aba SORTEIO (Base de dados)
num_rows = 10000
sort_data = {'SORT': list(range(1, num_rows + 1))}
for i in range(1, 6):
    sort_data[str(i)] = [None] * num_rows # Colunas vazias para preenchimento manual

df_sorteio = pd.DataFrame(sort_data)

# 2. Gerar o arquivo com o motor XlsxWriter para inserir fórmulas
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    # Salva a aba SORTEIO
    df_sorteio.to_excel(writer, sheet_name='SORTEIO', index=False)
    
    workbook  = writer.book
    sheet_sorteio = writer.sheets['SORTEIO']
    sheet_analise = workbook.add_worksheet('ANALISE')

    # Cabeçalhos da aba ANALISE
    headers = ['SORT', '1', '2', '3', '4', '5', 'CONTAGEM', 'ESTUDO']
    for col_num, header in enumerate(headers):
        sheet_analise.write(0, col_num, header)

    # 3. Preenchimento da aba ANALISE com fórmulas dinâmicas
    for row in range(1, num_rows + 1):
        # Excel usa base 1 para linhas (row + 1)
        xl_row = row + 1
        
        # Colunas A a F: Espelhamento da aba SORTEIO
        # Ex: =SORTEIO!A2
        for col_idx in range(6): 
            col_letter = chr(65 + col_idx) # A, B, C, D, E, F
            sheet_analise.write_formula(row, col_idx, f"=SORTEIO!{col_letter}{xl_row}")

        # Coluna G: CONTAGEM
        # Lógica: Concatena o número se a célula for TRUE, e remove espaços extras com TRIM
        formula_contagem = (
            f"=TRIM("
            f"IF(B{xl_row}=TRUE; \"1 \"; \"\") & "
            f"IF(C{xl_row}=TRUE; \"2 \"; \"\") & "
            f"IF(D{xl_row}=TRUE; \"3 \"; \"\") & "
            f"IF(E{xl_row}=TRUE; \"4 \"; \"\") & "
            f"IF(F{xl_row}=TRUE; \"5 \"; \"\")"
            f")"
        )
        sheet_analise.write_formula(row, 6, formula_contagem)

        # Coluna H: ESTUDO (Contagem Incremental)
        # Lógica: COUNTIF desde o topo até a linha atual
        # Ex: =COUNTIF($G$2:G2; G2)
        formula_estudo = f"=IF(G{xl_row}=\"\"; \"\"; COUNTIF($G$2:G{xl_row}; G{xl_row}))"
        sheet_analise.write_formula(row, 7, formula_estudo)

    # Formatação visual
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    sheet_analise.set_column('A:F', 8)
    sheet_analise.set_column('G:H', 15)

print(f"Planilha '{file_name}' gerada com sucesso com as abas SORTEIO e ANALISE!")