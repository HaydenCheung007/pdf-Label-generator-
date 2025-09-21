# pdf-Label-generator-
Through a input excel file you could generate labels with a specific format.
It uses **pandas**, **reportlab**, **pylibdmtx**, and **Pillow** to create carton and pallet labels with barcodes and DataMatrix codes.

## Features
- Reads input data from `input data.xlsx` (sheet: `Carton`or 'Pallet')
- Generates labels for **Cartons** and **Pallets** (Select though switching the dropdown cell on 'Select label type')
- After changing the fields(excel) change the responsible variable in the code
- Adds barcodes type(Code128) for item numbers, revisions, and Deleted
- Generates **DataMatrix ECC200(Bar code)**
- Handles missing data safely (blank cells don’t break the script)
- Exports a PDF named after: `{labeltype}Labels-{ItemNumber}-{SerTIN}.pdf`

## Requirements
- Python 3.9+  
- Libraries:
  - `pandas`
  - `reportlab`
  - `pylibdmtx`
  - `Pillow`
