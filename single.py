import pandas as pd
import numpy as np

# Configurações iniciais
file_name = 'Planilha_Sorteio_Numerica_V6.xlsx'
num_rows = 4000

# 1. Gerar Dados Aleatórios para teste (Apenas SORT e 1)
data_sorteio = {
    'SORT': list(range(1, num_rows + 1)),
    '1': [1 if val else None for val in np.random.choice([True, False], num_rows)]
}

df_sorteio = pd.DataFrame(data_sorteio)

with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    # --- ABA SORTEIO ---
    df_sorteio.to_excel(writer, sheet_name='SORTEIO', index=False)
    
    workbook = writer.book
    sheet_analise = workbook.add_worksheet('ANALISE')

    # Formatos
    fmt_int = workbook.add_format({'num_format': '0', 'align': 'center'})
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})

    # --- ABA ANALISE ---
    # Colunas reduzidas: SORT(A), 1(B), CONTAGEM(C)
    headers_an = ['SORT', '1', 'CONTAGEM']
    for c, h in enumerate(headers_an): 
        sheet_analise.write(0, c, h, fmt_header)

    for r in range(1, num_rows + 1):
        xl_r = r + 1 # Linha atual no Excel (2, 3, 4...)
        
        # Espelhamento da aba SORTEIO
        sheet_analise.write_formula(r, 0, f"=SORTEIO!A{xl_r}")
        sheet_analise.write_formula(r, 1, f"=SORTEIO!B{xl_r}")
        
        # --- Coluna C: CONTAGEM (Lógica de Streak/Sequência) ---
        if r == 1:
            # Primeira linha de dados: se tiver valor em B2, começa com 1, senão 0.
            f_contagem = f"=IF(B{xl_r}<>\"\"; 1; 0)"
        else:
            # Linhas subsequentes: se B atual preenchido, soma 1 ao anterior (C anterior), senão reseta.
            f_contagem = f"=IF(B{xl_r}<>\"\"; C{xl_r-1} + 1; 0)"
            
        sheet_analise.write_formula(r, 2, f_contagem, fmt_int)

    # Ajustes de layout
    sheet_sorteio = writer.sheets['SORTEIO']
    sheet_sorteio.set_column('A:B', 10, fmt_int)
    
    # Ajuste das colunas na ANALISE (A até C)
    sheet_analise.set_column('A:C', 12, fmt_int)

print(f"Planilha '{file_name}' gerada com sucesso! Coluna ESTUDO removida.")