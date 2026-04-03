# Spreadsheets Project — AI Instructions

This project generates Excel spreadsheets for lottery/draw statistical analysis using Python.
All logic is in `main.py`. There are no other source files.

## Environment

- Python 3.12 in `.venv/`
- Activate: `source .venv/bin/activate.fish` (fish) or `source .venv/bin/activate` (bash/zsh)
- Run: `python main.py`
- Dependencies: `pandas`, `xlsxwriter` (see `requirements.txt`)

## Architecture Rules

- **Single-file project**: do not split `main.py` into modules unless explicitly asked
- **No classes**: the script is procedural — keep it that way
- **xlsxwriter only**: all sheets except SORTEIO are built manually with `workbook.add_worksheet()` and `write_formula()`. Do not use openpyxl.
- **Formula separator**: use `;` (semicolon) in Excel formulas, not `,` — this project targets locales that use semicolon as separator

## Key Variable

```python
num_rows = 4000  # controls everything — all formulas derive from this
```

## Tab Construction Pattern

```python
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='SORTEIO', index=False)   # tab 1: pandas
    workbook = writer.book
    sheet = workbook.add_worksheet('NAME')                    # tabs 2+: manual
    # write headers at row 0
    # write formulas for rows 1..num_rows
```

Row indexing: Python uses 0-based rows; Excel formula strings use 1-based (`xl_r = r + 1`, data starts at `xl_r = 2`).

## Before Making Changes

1. Read `main.py` fully — it is short (~80 lines)
2. Check [formulas reference](.github/skills/spreadsheets/references/formulas.md) before touching any `write_formula` call
3. Check [modifications guide](.github/skills/spreadsheets/references/modifications.md) for the specific change pattern

## After Making Changes

Run `python main.py` and verify:

- No Python errors
- Output file is created
- File name matches `file_name` variable
