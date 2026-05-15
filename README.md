# portfolio-xlsx-price-loader

Python-based portfolio price loader and Excel updater for multi-asset tracking, with automatic CSV export and standalone HTML dashboard generation.

## Features

- Updates live prices for FX, gold, ETFs, stocks, and Bitcoin ETP instruments using Yahoo Finance.
- Writes current prices and performance metrics directly into a tracked Excel workbook.
- Exports a CSV snapshot after each run.
- Generates a standalone HTML dashboard with price cards, performance table, SWDA ATH panel, filters, theme toggle, and embedded charts.
- Works on both Windows and Linux with centralized path configuration.

## Current asset coverage

The script is currently configured for these instruments:

- GBP/EUR
- EUR/USD
- Gold Futures
- Invesco Physical Gold
- iShares Core MSCI World
- iShares Core MSCI EM IMI
- iShares MSCI World Small Cap
- Xtrackers MSCI World Energy
- iShares Developed Markets Property Yield
- Xtrackers MSCI World ESG
- Enel S.p.A.
- Bitcoin ETP

## Repository structure

```text
portfolio-xlsx-price-loader/
├── PAC_auto_update.py
├── README.md
├── LICENSE
├── .gitignore
└── requirements.txt
```

## Requirements

- Python 3.10 or newer recommended
- Microsoft Excel-compatible `.xlsx` or `.xlsm` tracker file
- Internet connection for Yahoo Finance data

Python packages:

- yfinance
- pandas
- openpyxl

Install dependencies with:

```bash
pip install -r requirements.txt
```

## Configuration

Edit the configuration block at the top of `PAC_auto_update.py`:

```python
# Windows: output folder, Excel file, CSV, and HTML
WIN_OUTPUT_DIR   = r'M:\\python\\portfolio'
WIN_TRACKER_FILE = r'M:\\pac.xlsx'
WIN_CSV_FILE     = r'M:\\python\\portfolio\\prezzi.csv'
WIN_HTML_FILE    = r'M:\\portfolio-tracker.html'

# Linux: output folder, Excel file, CSV, and HTML
LIN_OUTPUT_DIR   = '/home/user/path/to/output'
LIN_TRACKER_FILE = '/path/to/pac.xlsx'
LIN_CSV_FILE     = '/home/user/path/to/output/prezzi.csv'
LIN_HTML_FILE    = '/home/user/path/to/output/portfolio-tracker.html'

SHEET_NAME = 'portfolio OVERVIEW'
```

## How it works

1. Detects the running platform.
2. Loads the configured file paths.
3. Downloads live price data and historical series from Yahoo Finance.
4. Computes 1-day, 1-week, 1-month, 6-month, and 1-year performance.
5. Detects the SWDA all-time high using nominal prices with `auto_adjust=False`.
6. Updates the Excel workbook cells.
7. Saves a CSV export.
8. Generates a standalone HTML dashboard.

## Excel output mapping

The script currently writes values to these target cells:

- `P2` GBP_EUR
- `P3` USD_EUR
- `P6` XAU_USD
- `P7` GOLD_INVESCO
- `P9` GLOBALE
- `P10` EMERG
- `P11` SMALL_CAP
- `P12` ENERGY
- `P13` R.EST
- `P14` XDWS
- `P15` ENEL
- `P16` BTC_SPOT_IBIT
- `V9` SWDA ATH price
- `W9` SWDA ATH date
- `N19` update date
- `N20` update time

## Running the script

```bash
python PAC_auto_update.py
```

## Dashboard output

The generated HTML dashboard is standalone and can be opened directly in a browser. It includes:

- Price cards
- Performance table
- SWDA all-time high section
- Category filters
- Theme toggle
- CSV manual loader
- Four embedded performance charts

## Recommended public repo files

These files are recommended for a public GitHub repository:

- `README.md` for project overview and usage
- `LICENSE` to define reuse permissions
- `.gitignore` to avoid committing generated files, virtual environments, caches, and local spreadsheets
- `requirements.txt` for dependency installation

Optional later additions:

- `CHANGELOG.md`
- `CONTRIBUTING.md`
- `SECURITY.md`

## License choice

This repository is prepared with the MIT License because it is short, permissive, and commonly used for code that others may reuse, modify, and redistribute, including in closed-source projects, as long as the copyright and license notice are preserved.

## Notes before publishing

Before making the repository public, review and remove any private local paths, private file names, mount points, or personal environment details that you do not want to expose.
