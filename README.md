# Simplified Validation Pipeline (Cafe Orders)

A cut-down version of the NIHR INF annual reporting validation tool. It uses the same pipeline logic (fuzzy header detection → column mapping → per-column rules → value standardisation → conditional rules → colour-coded cells + audit report → parallel file processing) on a small 2-tab, 11-column dataset.

The data is generated rather than downloaded. The schema is modelled on Kaggle's *Cafe Sales – Dirty Data for Cleaning Training*, with deliberate mess added so every rule is exercised.

## Run it

```bash
pip install -r requirements.txt
python generate_messy_data.py   # -> input/Cafe_Orders_A.xlsx, input/Cafe_Orders_B.xlsx, input/answer_key.csv
python main.py                  # -> output/<file>.xlsx (labelled + reformatted), reports/Validation_Report_*.xlsx
python check_against_key.py     # compares output/ with the answer key
```

The files in `input/` are never modified, so you can re-run as often as you like.

## Structure

```text
├── config.py               # thresholds, fills, dropdowns, STANDARDISE_VALUES, VALIDATION_RULES + conditional_rules
├── detect_and_rename.py    # fuzzy header-row detection, column mapping, true max row
├── validator.py            # blank/number/date/dropdown/email checks, apply_conditions, Order ID cleaner
├── file_processing.py      # validate_file + save_report (openpyxl)
├── main.py                 # ProcessPoolExecutor entry point
├── generate_messy_data.py  # seeded messy data generator + answer key
└── check_against_key.py    # verification against the answer key
```

## What gets validated

**Store Details**: B3 must be a store from the `Stores` dropdown.

**Orders** (header on row 3 under a title banner; file A has misspelled headers, file B has columns in a different order):

| Column | Rules |
|---|---|
| Order ID | not blank; cleaned to `ORD-00000`, salvaged from Notes if blank |
| Order Date | not blank, valid date (standardised to `YYYY/MM/DD`), between 2020/01/01 and today |
| Item Category | not blank, dropdown |
| Item | not blank, dropdown |
| Quantity | number (strips `£`, `,`, spaces), 0–50 |
| Unit Price | not blank, number |
| Payment Method | not blank, dropdown (variants like `credit card` → `Card`) |
| Order Status | not blank, dropdown (`complete` → `Completed`) |
| Refund Date | valid date, plausible range |
| Customer Email | email format |

**Conditional rules**

- `required_if`: Refunded → Refund Date required
- `required_if_list`: Completed/Refunded → Quantity required
- `blanks_or_zero_list`: Cancelled → Quantity blank or 0
- `dropdown_dependency`: Item must belong to its Item Category
- `date_not_before`: Refund Date not before Order Date

## Colour key

| Colour | Meaning |
|---|---|
| Yellow `FFEA00` | Blank mandatory / conditionally required value |
| Green `50C878` | Invalid format (number, date, email), out of range, or broken date/zero rule |
| Light blue `00FFFF` | Value not in the dropdown, or not allowed for its parent column |

## Differences from the full pipeline

- A single rule set: there is no scheme switching and no special tabs (Biosamples, Publications and so on).
- The hard-coded per-list fixes have moved into `STANDARDISE_VALUES` in `config.py`.
- Conditional rules run once per row after all columns are cleaned. In the full pipeline they run inside the per-column loop.
- Validated copies are written to `output/` instead of overwriting the input.
- The report includes the row number for each issue.
