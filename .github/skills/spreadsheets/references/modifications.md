# Modifications Guide

Step-by-step instructions for common changes to `main.py`.

---

## 1. Change Number of Rows

Edit only one line at the top of `main.py`:

```python
# Before
num_rows = 4000

# After (example: 5000 rows)
num_rows = 5000
```

All Excel formulas use `num_rows` via f-strings and update automatically.

---

## 2. Rename the Output File

Edit only one line at the top of `main.py`:

```python
# Before
file_name = 'Planilha_Sorteio_Numerica_V6.xlsx'

# After
file_name = 'MeuArquivo.xlsx'
```

---

## 3. Add a 6th Number Column

This affects 4 places in `main.py`:

**Step 1** — Data generation (add column `6`):

```python
# Before
for i in range(1, 6):

# After
for i in range(1, 7):
```

**Step 2** — ANALISE headers (add `'6'`):

```python
# Before
headers_an = ['SORT', '1', '2', '3', '4', '5', 'CONTAGEM', 'ESTUDO']

# After
headers_an = ['SORT', '1', '2', '3', '4', '5', '6', 'CONTAGEM', 'ESTUDO']
```

**Step 3** — ANALISE mirror loop (extend to 7 columns):

```python
# Before
for c in range(6):

# After
for c in range(7):
```

**Step 4** — CONTAGEM formula (add column G token, shift CONTAGEM to col H):

```python
# Before: f_contagem refs B-F, written to col 6
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

# After: add G token, shift write positions +1
f_contagem = (
    f"=TRIM("
    f"IF(B{xl_r}<>\"\"; \"1 \"; \"0 \") & "
    f"IF(C{xl_r}<>\"\"; \"2 \"; \"0 \") & "
    f"IF(D{xl_r}<>\"\"; \"3 \"; \"0 \") & "
    f"IF(E{xl_r}<>\"\"; \"4 \"; \"0 \") & "
    f"IF(F{xl_r}<>\"\"; \"5 \"; \"0 \") & "
    f"IF(G{xl_r}<>\"\"; \"6 \"; \"0 \")"
    f")"
)
sheet_analise.write_formula(r, 7, f_contagem)
sheet_analise.write_formula(r, 8, f"=COUNTIF($H$2:H{xl_r}; H{xl_r})", fmt_int)
```

**Step 5** — ESTATISTICA combinations (extend to 6 positions):

```python
# Before
for bits in itertools.product([0, 1], repeat=5):

# After
for bits in itertools.product([0, 1], repeat=6):
```

This expands from 32 to 64 combinations.

**Step 6** — Update ESTATISTICA formulas to use new column range:
The SAIU formula references `ANALISE!$G$2:$G$...` — update to `$H$2:$H$...` (CONTAGEM moved to col H).
Same for ULTIMO's second range argument.

---

## 4. Load Real Data Instead of Random

Replace the data generation block with file reading.

**Before (random data):**

```python
data_sorteio = {'SORT': list(range(1, num_rows + 1))}
for i in range(1, 6):
    data_sorteio[str(i)] = [i if val else None for val in np.random.choice([True, False], num_rows)]
df_sorteio = pd.DataFrame(data_sorteio)
```

**After (from CSV):**

```python
df_sorteio = pd.read_csv('dados.csv')
num_rows = len(df_sorteio)  # update num_rows to match real data
```

**After (from Excel):**

```python
df_sorteio = pd.read_excel('dados.xlsx', sheet_name='SORTEIO')
num_rows = len(df_sorteio)
```

**Required DataFrame format:**
| SORT | 1 | 2 | 3 | 4 | 5 |
|------|---|---|---|---|---|
| 1 | 1 | NaN | 3 | NaN | 5 |
| 2 | NaN | 2 | NaN | NaN | NaN |

- Column names must be strings: `'SORT'`, `'1'`, `'2'`, `'3'`, `'4'`, `'5'`
- Drawn numbers: the integer value of the column (`1` for col `'1'`, `2` for col `'2'`, etc.)
- Not drawn: `NaN` or `None`

---

## 5. Add a New Statistic Column to ESTATISTICA

Example: adding a `PERCENTUAL` column (how often this combo appears as % of total).

**Step 1** — Add to headers:

```python
# Before
headers_es = ['COMBINA', 'SAIU', 'FREQUENCIA', 'ULTIMO', 'FALTA']

# After
headers_es = ['COMBINA', 'SAIU', 'FREQUENCIA', 'ULTIMO', 'FALTA', 'PERCENTUAL']
```

**Step 2** — Write the formula after the FALTA formula block:

```python
# After the existing FALTA line:
f_pct = f"=IF(B{xl_r}=0; 0; ROUND(B{xl_r}/{num_rows}*100; 2))"
sheet_est.write_formula(r, 5, f_pct, fmt_int)
```

**Step 3** — Adjust column width if needed:

```python
# Before
sheet_est.set_column('A:E', 15)

# After
sheet_est.set_column('A:F', 15)
```

---

## 6. Change Output Format (e.g., apply color to ESTATISTICA rows)

Add a conditional format after the data-writing loop:

```python
# Highlight combinations that appeared 0 times (SAIU = 0)
fmt_zero = workbook.add_format({'bg_color': '#FFCCCC'})
sheet_est.conditional_format(1, 1, 32, 1, {
    'type': 'cell',
    'criteria': '==',
    'value': 0,
    'format': fmt_zero
})
```

Place this block inside the `with pd.ExcelWriter(...) as writer:` block, after writing all ESTATISTICA rows.
