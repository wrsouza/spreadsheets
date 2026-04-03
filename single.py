import pandas as pd
import numpy as np

# Configurações iniciais
file_name = 'Planilha_Sorteio_Numerica_V6.xlsx'
num_rows = 4000

# 1. Gerar Dados Aleatórios para teste (Apenas SORT e Coluna 1)
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
    # A=SORT, B=1, C=CONTAGEM (Acertos), D=AUSENCIA (Vazios)
    headers_an = ['SORT', '1', 'CONTAGEM', 'AUSENCIA']
    for c, h in enumerate(headers_an): 
        sheet_analise.write(0, c, h, fmt_header)

    for r in range(1, num_rows + 1):
        xl_r = r + 1 
        
        sheet_analise.write_formula(r, 0, f"=SORTEIO!A{xl_r}")
        sheet_analise.write_formula(r, 1, f"=SORTEIO!B{xl_r}")
        
        # Lógica de Sequência de Acerto (Coluna C) e Ausência (Coluna D)
        if r == 1:
            f_contagem = f"=IF(B{xl_r}<>\"\"; 1; 0)"
            f_ausencia = f"=IF(B{xl_r}=\"\"; 1; 0)"
        else:
            f_contagem = f"=IF(B{xl_r}<>\"\"; C{xl_r-1} + 1; 0)"
            f_ausencia = f"=IF(B{xl_r}=\"\"; D{xl_r-1} + 1; 0)"
            
        sheet_analise.write_formula(r, 2, f_contagem, fmt_int)
        sheet_analise.write_formula(r, 3, f_ausencia, fmt_int)

    # --- ABA ESTATISTICA ---
    # Estrutura: SEQ | SAIU | FREQ | ULT | FALTA | NAO SAIU | FREQ(NS) | ULT(NS)
    headers_es = [
        'SEQUENCIA', 'SAIU', 'FREQUENCIA', 'ULTIMO', 'FALTA', 
        'NAO SAIU', 'FREQUENCIA', 'ULTIMO'
    ]
    for c, h in enumerate(headers_es): 
        sheet_est.write(0, c, h, fmt_header)

    for i in range(1, 21):
        r = i 
        xl_r = r + 1
        
        # A: SEQUENCIA (1 a 20)
        sheet_est.write(r, 0, i, fmt_int)
        
        # --- BLOCO DE ACERTOS (SAIU) ---
        # B: SAIU (Quantas vezes a sequência de acertos ocorreu)
        f_saiu = f"=COUNTIF(ANALISE!$C$2:$C${num_rows+1}; A{xl_r})"
        sheet_est.write_formula(r, 1, f_saiu, fmt_int)
        
        # C: FREQUENCIA (Média de acertos)
        f_freq_s = f"=IF(B{xl_r}=0; 0; INT({num_rows}/B{xl_r}))"
        sheet_est.write_formula(r, 2, f_freq_s, fmt_int)
        
        # D: ULTIMO (Último sorteio do acerto)
        f_ult_s = f"=IF(B{xl_r}=0; 0; MAXIFS(ANALISE!$A$2:$A${num_rows+1}; ANALISE!$C$2:$C${num_rows+1}; A{xl_r}))"
        sheet_est.write_formula(r, 3, f_ult_s, fmt_int)
        
        # E: FALTA (Distância do último acerto)
        f_falta = f"=IF(D{xl_r}=0; {num_rows}; {num_rows} - D{xl_r})"
        sheet_est.write_formula(r, 4, f_falta, fmt_int)

        # --- BLOCO DE AUSÊNCIAS (NAO SAIU) ---
        # F: NAO SAIU (Quantas vezes a sequência de vazios ocorreu)
        f_nao_saiu = f"=COUNTIF(ANALISE!$D$2:$D${num_rows+1}; A{xl_r})"
        sheet_est.write_formula(r, 5, f_nao_saiu, fmt_int)

        # G: FREQUENCIA (Média de ausências baseada na coluna F)
        f_freq_ns = f"=IF(F{xl_r}=0; 0; INT({num_rows}/F{xl_r}))"
        sheet_est.write_formula(r, 6, f_freq_ns, fmt_int)

        # H: ULTIMO (Último sorteio da ausência baseada na coluna D da ANALISE)
        f_ult_ns = f"=IF(F{xl_r}=0; 0; MAXIFS(ANALISE!$A$2:$A${num_rows+1}; ANALISE!$D$2:$D${num_rows+1}; A{xl_r}))"
        sheet_est.write_formula(r, 7, f_ult_ns, fmt_int)

    # Ajustes de layout e largura de colunas
    sheet_sorteio = writer.sheets['SORTEIO']
    sheet_sorteio.set_column('A:B', 10, fmt_int)
    sheet_analise.set_column('A:D', 12, fmt_int)
    sheet_est.set_column('A:H', 15, fmt_int)

print(f"Planilha '{file_name}' gerada com sucesso!")
print("Estatísticas de Acertos (B-E) e Ausências (F-H) configuradas.")