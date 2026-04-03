import pandas as pd
import numpy as np
import itertools

# Configurações iniciais
file_name = 'Planilha_Sorteio_Numerica_V6.xlsx'
num_rows = 4000

# 1. Gerar Dados Aleatórios para teste (Já com a nova regra: Número ou Vazio)
data_sorteio = {'SORT': list(range(1, num_rows + 1))}
for i in range(1, 6):
    # Gera o número da coluna se "sorteado", caso contrário fica vazio (None)
    data_sorteio[str(i)] = [i if val else None for val in np.random.choice([True, False], num_rows)]

df_sorteio = pd.DataFrame(data_sorteio)

# 2. Gerar as 32 combinações para a aba ESTATISTICA
combinações = []
for bits in itertools.product([0, 1], repeat=5):
    label = " ".join([str(i+1) if val else "0" for i, val in enumerate(bits)])
    combinações.append(label)

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
    headers_an = ['SORT', '1', '2', '3', '4', '5', 'CONTAGEM', 'ESTUDO']
    for c, h in enumerate(headers_an): 
        sheet_analise.write(0, c, h, fmt_header)

    for r in range(1, num_rows + 1):
        xl_r = r + 1
        # Espelhamento da aba SORTEIO
        for c in range(6): 
            sheet_analise.write_formula(r, c, f"=SORTEIO!{chr(65+c)}{xl_r}")
        
        # Coluna G: CONTAGEM (Ajustada para checar se a célula não está vazia)
        # Se na SORTEIO estiver o número 1, a ANALISE entende como "1", se estiver vazio, entende como "0"
        f_contagem = (
            f"=TRIM("
            f"IF(B{xl_r}<>\"\"; \"1 \"; \"0 \") & "
            f"IF(C{xl_r}<>\"\"; \"2 \"; \"0 \") & "
            f"IF(D{xl_r}<>\"\"; \"3 \"; \"0 \") & "
            f"IF(E{xl_r}<>\"\"; \"4 \"; \"0 \") & "
            f"IF(F{xl_r}<>\"\"; \"5 \"; \"0 \")"
            f")"
        )
        sheet_analise.write_formula(r, 6, f_contagem)
        sheet_analise.write_formula(r, 7, f"=COUNTIF($G$2:G{xl_r}; G{xl_r})", fmt_int)

    # --- ABA ESTATISTICA ---
    headers_es = ['COMBINA', 'SAIU', 'FREQUENCIA', 'ULTIMO', 'FALTA']
    for c, h in enumerate(headers_es): 
        sheet_est.write(0, c, h, fmt_header)

    for i, comb in enumerate(combinações):
        r = i + 1
        xl_r = r + 1
        sheet_est.write(r, 0, comb)
        
        # SAIU
        sheet_est.write_formula(r, 1, f"=COUNTIF(ANALISE!$G$2:$G${num_rows+1}; A{xl_r})", fmt_int)
        # FREQUENCIA
        sheet_est.write_formula(r, 2, f"=IF(B{xl_r}=0; 0; INT({num_rows}/B{xl_r}))", fmt_int)
        # ULTIMO (MAXIFS para Google Sheets)
        f_ultimo = f"=IF(B{xl_r}=0; 0; MAXIFS(ANALISE!$A$2:$A${num_rows+1}; ANALISE!$G$2:$G${num_rows+1}; A{xl_r}))"
        sheet_est.write_formula(r, 3, f_ultimo, fmt_int)
        # FALTA
        f_falta = f"=IF(B{xl_r}=0; 0; MAX(0; (D{xl_r} + C{xl_r}) - {num_rows}))"
        sheet_est.write_formula(r, 4, f_falta, fmt_int)

    # Ajustes finais
    sheet_sorteio = writer.sheets['SORTEIO']
    sheet_sorteio.set_column('A:F', 10, fmt_int)
    sheet_analise.set_column('G:H', 12)
    sheet_est.set_column('A:E', 15)

print(f"Planilha '{file_name}' gerada com sucesso!")