import pandas as pd

# 1. Criar o intervalo sequencial para a coluna SORT (1 a 10.000)
sort_column = list(range(1, 10001))

# 2. Criar o DataFrame com a coluna inicial
df = pd.DataFrame({'SORT': sort_column})

# 3. Adicionar as colunas de 1 a 5 vazias (conforme solicitado: sem preencher TRUE/FALSE)
# Usamos None para que as células fiquem prontas para a entrada manual
for i in range(1, 6):
    df[str(i)] = None

# 4. Definir o nome do arquivo
file_name = 'Planilha_Sorteio_Estatistico.xlsx'

# 5. Gerar a planilha com a aba renomeada para SORTEIO
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='SORTEIO', index=False)
    
    # Acessar o objeto workbook e worksheet para ajustes finos (opcional)
    workbook  = writer.book
    worksheet = writer.sheets['SORTEIO']
    
    # Ajustar a largura das colunas para melhor visualização
    worksheet.set_column('A:A', 10)  # Coluna SORT
    worksheet.set_column('B:F', 8)   # Colunas 1 a 5

print(f"Planilha '{file_name}' criada com sucesso!")