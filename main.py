import pandas as pd
import itertools

file_name = 'Planilha_Sorteio_Estatistico_V3.xlsx'
num_rows = 10000

# 1. Gerar as 32 combinações possíveis (Espaço Amostral)
combinações = []
for bits in itertools.product([0, 1], repeat=5):
    # Transforma (1, 0, 1, 0, 0) em "1 2 0 0 0" ou "0 0 0 0 0"
    label = " ".join([str(i+1) if val else "0" for i, val in enumerate(bits)])
    # Caso especial: se for tudo 0, mantemos "0 0 0 0 0"
    combinações.append(label)

with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    # --- ABA SORTEIO ---
    df_sorteio = pd.DataFrame({'SORT': range(1, num_rows + 1)})
    for i in range(1, 6): df_sorteio[str(i)] = None
    df_sorteio.to_excel(writer, sheet_name='SORTEIO', index=False)

    # --- ABA ANALISE (Fórmulas) ---
    sheet_analise = writer.book.add_worksheet('ANALISE')
    analise_headers = ['SORT', '1', '2', '3', '4', '5', 'CONTAGEM', 'ESTUDO']
    for c, h in enumerate(analise_headers): sheet_analise.write(0, c, h)
    
    for r in range(1, num_rows + 1):
        xl_r = r + 1
        for c in range(6): sheet_analise.write_formula(r, c, f"=SORTEIO!{chr(65+c)}{xl_r}")
        
        # Lógica CONTAGEM (ajustada para incluir '0' quando FALSE)
        f_contagem = f"=TRIM(IF(B{xl_r}=TRUE;\"1 \";\"0 \") & IF(C{xl_r}=TRUE;\"2 \";\"0 \") & IF(D{xl_r}=TRUE;\"3 \";\"0 \") & IF(E{xl_r}=TRUE;\"4 \";\"0 \") & IF(F{xl_r}=TRUE;\"5 \";\"0 \"))"
        sheet_analise.write_formula(r, 6, f_contagem)
        sheet_analise.write_formula(r, 7, f"=IF(G{xl_r}=\"\";\"\";COUNTIF($G$2:G{xl_r};G{xl_r}))")

    # --- ABA ESTATISTICA ---
    sheet_est = writer.book.add_worksheet('ESTATISTICA')
    est_headers = ['COMBINA', 'SAIU', 'FREQUENCIA', 'ULTIMO', 'FALTA']
    for c, h in enumerate(est_headers): sheet_est.write(0, c, h)

    for i, comb in enumerate(combinações):
        r = i + 1
        xl_r = r + 1
        sheet_est.write(r, 0, comb) # Coluna COMBINA
        
        # SAIU: Conta na aba ANALISE
        sheet_est.write_formula(r, 1, f"=COUNTIF(ANALISE!$G$2:$G${num_rows+1}; A{xl_r})")
        
        # FREQUENCIA: Total / SAIU (Inteiro)
        sheet_est.write_formula(r, 2, f"=IF(B{xl_r}=0; 0; INT({num_rows}/B{xl_r}))")
        
        # ULTIMO: Procura o último SORT onde a combinação ocorreu
        # Usamos AGGREGATE para ignorar erros e pegar o maior valor (MAX)
        f_ultimo = f"=IF(B{xl_r}=0; 0; AGGREGATE(14; 6; ANALISE!$A$2:$A${num_rows+1}/(ANALISE!$G$2:$G${num_rows+1}=A{xl_r}); 1))"
        sheet_est.write_formula(r, 3, f_ultimo)
        
        # FALTA: (Ultimo + Frequencia) - Max_Sort_Atual
        # Consideramos o Max Sort como o último valor preenchido na aba SORTEIO
        f_falta = f"=IF(B{xl_r}=0; 0; MAX(0; (D{xl_r} + C{xl_r}) - MAX(ANALISE!$A$2:$A${num_rows+1})))"
        sheet_est.write_formula(r, 4, f_falta)

    # Formatação Final
    fmt_int = writer.book.add_format({'num_format': '0'})
    sheet_est.set_column('A:E', 15, fmt_int)

print(f"Planilha '{file_name}' completa com as 3 abas gerada!")