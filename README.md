# Click & Drop Address Batch Converter

This repository contains a single Python script that converts plaintext address batches into a Royal Mail Click & Drop Desktop import XLSX.

## Requirements

- Python 3.11+
- `openpyxl` (required)
- `rich` (optional, for nicer progress output)

Install dependencies:

```bash
py -m pip install -r requirements.txt
```

## Usage

```bash
py click_drop_import.py --input "C:\path\addresses.txt"
```

Optional flags:

```bash
py click_drop_import.py \
  --input "C:\path\addresses.txt" \
  --outdir ".\output" \
  --filename "orders_custom.xlsx" \
  --default-weight 0.5 \
  --default-format "Parcel" \
  --service-code "" \
  --strict \
  --save-rejects \
  --debug \
  --split-name
```

Run the built-in self-test (no files written):

```bash
py click_drop_import.py --self-test
```

## Notes

- The script does **not** copy files into the Click & Drop watch folder. You should copy the XLSX manually.
- Input is read with UTF-8 (with BOM support) and errors are replaced to ensure compatibility with Windows terminals.
- Logs are written to `.\logs\app.log` with rotation.
- Use `--split-name` if your Click & Drop mapping expects separate first/last name fields.
