# Extract vs Rosetta Comparison

## Purpose

`compare_and_merge.py` compares the Rosetta/PMS truth sheet against the OCR contract extraction output.

This is the script used for the "OCR extract compared to Rosetta" step that produces the main comparison workbook.

## Inputs

The script expects these files in the original `Last One` folder:

```text
Last One/
  Senora Rosetta V5.xlsx
  Unified_Contracts_Data_v1.0.xlsx
```

`Senora Rosetta V5.xlsx` is the Rosetta sheet with NCs attached to PMS/Tenancy data.

`Unified_Contracts_Data_v1.0.xlsx` is the normalized OCR/extraction workbook from the contracts.

## Comparison Logic

The script indexes both files by NC and compares:

- Deposit
- Rent
- Start date
- End date

It normalizes NC formats, parses Euro/money strings, and normalizes date formats before comparing.

It also detects:

- NCs present in both files.
- NCs only in Rosetta.
- NCs only in the OCR/extraction file.
- Duplicate NCs.
- Field-level mismatches.

## Output

The script writes:

```text
Last One/Rosetta_vs_Contract_Comparison.xlsx
```

The output workbook includes detailed comparison sheets such as:

- Full comparison
- Mismatches only
- Summary
- Rosetta-only NCs
- Contract/OCR-only NCs

## Notes

This is different from `compare_rosetta_vs_tenancy.py`, which compares Rosetta back against the filtered tenancy report. This script is specifically for Rosetta vs OCR/extraction.

