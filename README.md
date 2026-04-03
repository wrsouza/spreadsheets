# Planilha Sorteio Numérica

## Visão Geral

Script Python que gera uma planilha `.xlsx` com **3 abas** para análise estatística de sorteios numéricos.
Arquivo gerado: `Planilha_Sorteio_Numerica_V6.xlsx` com 4000 linhas de dados simulados.

A lógica central: para cada linha de sorteio, verifica quais dos 5 números foram sorteados, monta uma string de combinação (ex: `"1 0 3 0 5"`), e na aba ESTATISTICA conta a frequência de cada uma das 32 combinações possíveis.

---

## Estrutura do Projeto

```
spreadsheets/
├── main.py          # Script principal — toda a lógica está aqui
├── requirements.txt # Dependências: pandas, xlsxwriter
└── .venv/           # Ambiente virtual Python
```

---

## Dependências

```
pandas
xlsxwriter
```

`itertools` e `numpy` são built-in/já incluídos — não precisam de instalação separada.

### Instalação

```bash
python3 -m venv .venv
source .venv/bin/activate.fish   # fish shell
# ou
source .venv/bin/activate        # bash/zsh

pip install -r requirements.txt
```

### Execução

```bash
python main.py
```

Gera o arquivo `Planilha_Sorteio_Numerica_V6.xlsx` no diretório atual.

---

## Constantes Configuráveis

| Variável    | Padrão                                | Descrição                       |
| ----------- | ------------------------------------- | ------------------------------- |
| `file_name` | `'Planilha_Sorteio_Numerica_V6.xlsx'` | Nome do arquivo gerado          |
| `num_rows`  | `4000`                                | Quantidade de linhas de sorteio |

Todas as fórmulas referenciam `num_rows` dinamicamente via f-strings.

---

## Estrutura da Planilha Gerada

### Aba SORTEIO

Gerada via `df_sorteio.to_excel(writer, sheet_name='SORTEIO', index=False)`.

| Coluna    | Descrição                                                                        |
| --------- | -------------------------------------------------------------------------------- |
| `SORT`    | Número sequencial (1 a `num_rows`)                                               |
| `1` a `5` | Se sorteado: inteiro do índice da coluna. Se não sorteado: célula vazia (`None`) |

**Geração dos dados aleatórios:**

```python
data_sorteio = {'SORT': list(range(1, num_rows + 1))}
for i in range(1, 6):
    data_sorteio[str(i)] = [i if val else None
                             for val in np.random.choice([True, False], num_rows)]
df_sorteio = pd.DataFrame(data_sorteio)
```

Para cada coluna `i` (1 a 5): se `True` → insere o inteiro `i`; se `False` → insere `None` (célula vazia no Excel).

---

### Aba ANALISE

Criada manualmente via `workbook.add_worksheet('ANALISE')`. Espelha a aba SORTEIO e adiciona 2 colunas calculadas.

| Col | Header     | Fórmula / Conteúdo                                                         |
| --- | ---------- | -------------------------------------------------------------------------- |
| A   | `SORT`     | `=SORTEIO!A{r}`                                                            |
| B   | `1`        | `=SORTEIO!B{r}`                                                            |
| C   | `2`        | `=SORTEIO!C{r}`                                                            |
| D   | `3`        | `=SORTEIO!D{r}`                                                            |
| E   | `4`        | `=SORTEIO!E{r}`                                                            |
| F   | `5`        | `=SORTEIO!F{r}`                                                            |
| G   | `CONTAGEM` | String de combinação — ver fórmula abaixo                                  |
| H   | `ESTUDO`   | `=COUNTIF($G$2:G{r}; G{r})` — total acumulado da combinação até essa linha |

**Fórmula CONTAGEM (coluna G) — chave de toda a análise:**

```excel
=TRIM(
  IF(B{r}<>""; "1 "; "0 ") &
  IF(C{r}<>""; "2 "; "0 ") &
  IF(D{r}<>""; "3 "; "0 ") &
  IF(E{r}<>""; "4 "; "0 ") &
  IF(F{r}<>""; "5 "; "0 ")
)
```

Produz strings como `"1 0 3 0 5"`, `"0 0 0 0 0"`, `"1 2 3 4 5"`.
O `TRIM` remove o espaço final. Essa string é a chave de lookup usada na aba ESTATISTICA.

No Python, a fórmula é montada assim:

```python
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
```

---

### Aba ESTATISTICA

Criada manualmente via `workbook.add_worksheet('ESTATISTICA')`. Lista as **32 combinações possíveis** (2⁵) com estatísticas de cada uma.

| Col | Header       | Fórmula                                                                                    | Descrição                                          |
| --- | ------------ | ------------------------------------------------------------------------------------------ | -------------------------------------------------- |
| A   | `COMBINA`    | valor fixo                                                                                 | String da combinação (ex: `"1 2 0 4 0"`)           |
| B   | `SAIU`       | `=COUNTIF(ANALISE!$G$2:$G${num_rows+1}; A{r})`                                             | Total de ocorrências nas `num_rows` linhas         |
| C   | `FREQUENCIA` | `=IF(B{r}=0; 0; INT({num_rows}/B{r}))`                                                     | A cada quantas linhas aparece em média             |
| D   | `ULTIMO`     | `=IF(B{r}=0; 0; MAXIFS(ANALISE!$A$2:$A${num_rows+1}; ANALISE!$G$2:$G${num_rows+1}; A{r}))` | SORT da última ocorrência                          |
| E   | `FALTA`      | `=IF(B{r}=0; 0; MAX(0; (D{r} + C{r}) - {num_rows}))`                                       | Linhas que faltam para próxima ocorrência esperada |

**Geração das 32 combinações via `itertools`:**

```python
combinações = []
for bits in itertools.product([0, 1], repeat=5):
    label = " ".join([str(i+1) if val else "0" for i, val in enumerate(bits)])
    combinações.append(label)
```

`itertools.product([0,1], repeat=5)` gera todas as combinações binárias de 5 posições (2⁵ = 32).
Para bit `1` → usa `i+1` (o número da posição); para bit `0` → usa `"0"`.
Resultado: de `"0 0 0 0 0"` até `"1 2 3 4 5"` (32 strings).

---

## Formatação Aplicada

| Objeto                           | Definição                            | Onde é aplicado       |
| -------------------------------- | ------------------------------------ | --------------------- |
| `fmt_int`                        | `num_format='0'`, `align='center'`   | Colunas numéricas     |
| `fmt_header`                     | bold, bg `#D7E4BC`, border 1, center | Cabeçalhos das 3 abas |
| `set_column('A:F', 10, fmt_int)` | largura 10 + fmt_int                 | Aba SORTEIO           |
| `set_column('G:H', 12)`          | largura 12                           | Aba ANALISE           |
| `set_column('A:E', 15)`          | largura 15                           | Aba ESTATISTICA       |

---

## Fluxo de Execução Completo

```
1. Criar df_sorteio com dados aleatórios (numpy)
2. Gerar lista de 32 combinações (itertools.product)
3. Abrir pd.ExcelWriter(file_name, engine='xlsxwriter')
   ├── df_sorteio.to_excel() → aba SORTEIO
   ├── workbook.add_worksheet('ANALISE')
   │   ├── Escrever cabeçalhos com fmt_header
   │   └── Para r em range(1, num_rows+1):  [xl_r = r+1]
   │       ├── Cols A-F: write_formula espelhando SORTEIO!{col}{xl_r}
   │       ├── Col G:    write_formula CONTAGEM (string combinação via TRIM+IF)
   │       └── Col H:    write_formula ESTUDO (COUNTIF acumulado)
   └── workbook.add_worksheet('ESTATISTICA')
       ├── Escrever cabeçalhos com fmt_header
       └── Para i, comb em enumerate(combinações):  [r=i+1, xl_r=r+1]
           ├── Col A: write() — string fixa da combinação
           ├── Col B: write_formula SAIU (COUNTIF em ANALISE!$G)
           ├── Col C: write_formula FREQUENCIA (INT division)
           ├── Col D: write_formula ULTIMO (MAXIFS em ANALISE!$A e $G)
           └── Col E: write_formula FALTA (MAX com subtração)
4. Fechar writer → salvar .xlsx → print confirmação
```

---

## Como Modificar

**Adicionar 6º número (coluna extra):**

- `range(1, 6)` → `range(1, 7)` na geração de `data_sorteio`
- `repeat=5` → `repeat=6` no `itertools.product`
- Adicionar `"6"` em `headers_an`
- Acrescentar mais um `IF(G{xl_r}<>""; "6 "; "0 ")` na fórmula CONTAGEM

**Mudar número de linhas:**

- Alterar apenas `num_rows = 4000` — todas as fórmulas se ajustam automaticamente

**Usar dados reais em vez de aleatórios:**

```python
# Substituir o bloco de geração de data_sorteio por:
df_sorteio = pd.read_csv('dados.csv')
# ou
df_sorteio = pd.read_excel('dados.xlsx')
# O DataFrame deve ter colunas: SORT, 1, 2, 3, 4, 5
# Células preenchidas com o número da coluna ou NaN/None para vazio
```
