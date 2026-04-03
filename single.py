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
    sheet_est = workbook.add_worksheet('ESTATISTICA')

    # Formatos
    fmt_int = workbook.add_format({'num_format': '0', 'align': 'center'})
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})

    # --- ABA ANALISE ---
    headers_an = ['SORT', '1', 'CONTAGEM']
    for c, h in enumerate(headers_an): 
        sheet_analise.write(0, c, h, fmt_header)

    for r in range(1, num_rows + 1):
        xl_r = r + 1 # Linha atual no Excel
        
        sheet_analise.write_formula(r, 0, f"=SORTEIO!A{xl_r}")
        sheet_analise.write_formula(r, 1, f"=SORTEIO!B{xl_r}")
        
        # Coluna C: Lógica de Streak (Sequência)
        if r == 1:
            f_contagem = f"=IF(B{xl_r}<>\"\"; 1; 0)"
        else:
            f_contagem = f"=IF(B{xl_r}<>\"\"; C{xl_r-1} + 1; 0)"
            
        sheet_analise.write_formula(r, 2, f_contagem, fmt_int)

    # --- ABA ESTATISTICA ---
    headers_es = ['SEQUENCIA', 'SAIU', 'FREQUENCIA', 'ULTIMO', 'FALTA']
    for c, h in enumerate(headers_es): 
        sheet_est.write(0, c, h, fmt_header)

    # Gerar de 1 a 20
    for i in range(1, 21):
        r = i 
        xl_r = r + 1
        
        # Coluna A: A sequência (1, 2, 3... 20)
        sheet_est.write(r, 0, i, fmt_int)
        
        # Coluna B: SAIU (Quantas vezes essa sequência específica foi atingida)
        # Analisa a coluna C da aba ANALISE
        f_saiu = f"=COUNTIF(ANALISE!$C$2:$C${num_rows+1}; A{xl_r})"
        sheet_est.write_formula(r, 1, f_saiu, fmt_int)
        
        # Coluna C: FREQUENCIA (Média de sorteios para cada ocorrência)
        f_freq = f"=IF(B{xl_r}=0; 0; INT({num_rows}/B{xl_r}))"
        sheet_est.write_formula(r, 2, f_freq, fmt_int)
        
        # Coluna D: ULTIMO (O número do sorteio/SORT da última vez que atingiu essa sequência)
        f_ultimo = f"=IF(B{xl_r}=0; 0; MAXIFS(ANALISE!$A$2:$A${num_rows+1}; ANALISE!$C$2:$C${num_rows+1}; A{xl_r}))"
        sheet_est.write_formula(r, 3, f_ultimo, fmt_int)
        
        # Coluna E: FALTA (Distância do último sorteio para o sorteio atual)
        f_falta = f"=IF(D{xl_r}=0; {num_rows}; {num_rows} - D{xl_r})"
        sheet_est.write_formula(r, 4, f_falta, fmt_int)

    # Ajustes finais de layout
    sheet_sorteio = writer.sheets['SORTEIO']
    sheet_sorteio.set_column('A:B', 10, fmt_int)
    sheet_analise.set_column('A:C', 12, fmt_int)
    sheet_est.set_column('A:E', 15, fmt_int)

print(f"Planilha '{file_name}' gerada com sucesso!")
print("Aba ESTATISTICA agora analisa sequências de 1 a 20.")