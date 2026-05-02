# Automated Quotations

An AI-powered automation system that reads Herholz door product data and automatically responds to customer quotation requests received via message or email.

## What it does

Customers contact the business asking for door prices. Instead of manually looking up prices and replying, this system automates the process:

1. Reads the [Herholz product PDF](https://drive.google.com/file/d/1lqaFe1JBV6CVxDXyNLrbG0QyaESwv5Fj/view?usp=sharing) and price sheets
2. Parses the Excel price data and splits it into per-door CSV files
3. Feeds the relevant door data (including price tier colours) to an AI
4. The AI reads the customer's enquiry and generates a quotation response

## Supplier

**Herholz** — [herholz.de](http://herholz.de/)
Herholz is the parent company and manufacturer. The business resells Herholz doors to end customers as an authorised middleman.

## Project structure

```
Automated Quotations/
├── excel_inputs_files/         # Raw Excel price sheets from Herholz
│   └── Standard Decora Door Prices Gefalzt.xlsx
├── excel_output_files/         # One CSV per door, ready to feed to AI
│   └── <door name>.csv
├── split_doors.py              # Splits the master Excel into per-door CSVs
├── read_excel.py               # Utility to inspect cell colours in the Excel
└── README.md
```

## Price colour coding

The Excel price sheet uses cell background colours to indicate price tiers:

| Colour | Meaning |
|--------|---------|
| Green (`#92D050`) | Standard price |
| Yellow (`#FFFF00`) | Mid-tier price |
| Amber (`#FFC000`) | Premium / special price |

These colour labels are preserved in the CSV output as `price (colour)`, e.g. `191 (amber)`.

## Scripts

### `split_doors.py`
Reads the master Excel price file and writes one `.csv` file per door into `excel_output_files/`.
Each CSV includes the header row, both height variants (1985mm and 2110mm), and colour annotations on prices.

```bash
python split_doors.py
```

### `read_excel.py`
Utility script to inspect and verify cell fill colours in the source Excel file.

```bash
python read_excel.py
```

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install openpyxl
```

## How quotations work

1. A customer sends an enquiry specifying a door model, size, and finish
2. The relevant CSV from `excel_output_files/` is passed to the AI along with the customer's message
3. The AI reads the pricing data and the [Herholz product PDF](https://drive.google.com/file/d/1lqaFe1JBV6CVxDXyNLrbG0QyaESwv5Fj/view?usp=sharing), then drafts a quotation reply
