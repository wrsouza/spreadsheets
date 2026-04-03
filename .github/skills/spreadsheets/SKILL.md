---
name: spreadsheets
description: "Use when working with this project: generating Excel spreadsheets, modifying sorteio/lottery analysis, adding columns, changing row count, editing formulas, running main.py, understanding the 3-tab structure (SORTEIO, ANALISE, ESTATISTICA), or extending the Python xlsxwriter script."
argument-hint: "Describe what you want to change or execute in the spreadsheet project"
---

# Spreadsheets — Planilha Sorteio Numérica

## What This Project Does

Generates `Planilha_Sorteio_Numerica_V6.xlsx` with 3 Excel sheets for lottery/draw statistical analysis.
All logic is in a single file: `main.py`.

## How to Run

```bash
# Activate virtual environment (fish shell)
source .venv/bin/activate.fish

# Run
python main.py
```

Output: `Planilha_Sorteio_Numerica_V6.xlsx` in the project root.

## Key Constants (top of main.py)

```python
file_name = 'Planilha_Sorteio_Numerica_V6.xlsx'  # output file name
num_rows  = 4000                                   # number of draw rows
```

All Excel formulas reference `num_rows` dynamically via f-strings — change only this variable to resize everything.

## Project Structure

```
main.py          ← entire logic (single file)
requirements.txt ← pandas, xlsxwriter
.venv/           ← Python virtual environment
```

## 3-Tab Spreadsheet Architecture

### Tab 1 — SORTEIO (input data)

- Col A `SORT`: row sequence 1..num_rows
- Cols B-F `1`..`5`: integer if drawn, empty (`None`) if not
- Generated with pandas DataFrame + `numpy.random.choice`

### Tab 2 — ANALISE (formula engine)

- Cols A-F: mirror SORTEIO via `=SORTEIO!{col}{row}`
- Col G `CONTAGEM`: key lookup string built with TRIM+IF (see [formulas reference](./references/formulas.md))
- Col H `ESTUDO`: `=COUNTIF($G$2:G{r}; G{r})` — running count of each combination

### Tab 3 — ESTATISTICA (statistics)

- 32 rows — one per possible combination (2⁵)
- Combinations generated with `itertools.product([0,1], repeat=5)`
- Cols: COMBINA, SAIU, FREQUENCIA, ULTIMO, FALTA (see [formulas reference](./references/formulas.md))

## Common Modifications

See [modifications guide](./references/modifications.md) for step-by-step instructions on:

- Adding a 6th number column
- Changing number of rows
- Loading real data instead of random
- Renaming output file
- Adding new statistics columns
