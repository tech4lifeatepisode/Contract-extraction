# Rosetta Builder

## Purpose

`match_nc_latest.py` creates the Rosetta-style mapping by attaching NC IDs to the PMS/Tenancy export.

This is the final script used to solve the problem that the THM tenancy/PMS report has `Contract ID` and `Tenant ID`, but not the NC ID needed to compare against contract/OCR extraction data.

## Inputs

The script expects these files in the original recovered working layout:

```text
PMS Issue/
  Final Contract Extractions 10.04.2026.xlsx
  Tech Version - Ana_s Doc.xlsx
  Tech Version - Doc Seguimiento 10-04 v1.0.xlsx

Last One/
  Tenancy-Latest2.xlsx
```

## Matching Sources

It builds NC reference data from three sources:

- Final contract extraction: NC, tenant name, ID number.
- Ana report: NC, tenant name, DNI/passport.
- Seguimiento: NC, first name, last name.

It then maps those NCs onto the PMS/Tenancy export using:

- Tenant ID / DNI / passport matching.
- Exact normalized name matching.
- Fuzzy name matching.
- Positional pairing for people with multiple contracts and multiple NCs.

## Output

The script opens the original `Tenancy-Latest2.xlsx`, inserts a new first column named `NC`, fills it with matched NC IDs, and saves:

```text
Last One/Tenancy-Latest2_with_NC.xlsx
```

It is careful to preserve the original workbook values and formatting, shifting the original columns right by one.

## Related Validation

`verify_output.py` was used after this step to confirm the output workbook matched the original except for the inserted NC column.

