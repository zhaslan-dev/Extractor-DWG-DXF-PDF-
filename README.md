# Extractor-DWG-DXF-PDF-```markdown
# Generic DXF Table Parser

A Python tool that extracts structured table data from DXF files (AutoCAD drawings), identifies records with identifiers, and exports the results to Excel.

## Overview

This parser reads DXF files, locates tables that follow a specific layout (numbered rows, header row with 13 generic column names), and extracts values from designated columns. It can handle multi‑column tables (split across several vertical sections) and merged headers. The extracted data is aggregated by unique identifiers and exported into a ready‑to‑use Excel report.

The code has been deliberately cleaned of any domain‑specific terminology so that it can be adapted to various types of engineering drawings where similar table conventions are used.

## Features

- Finds table anchors using a `"number/number:Suffix"` pattern.
- Detects header rows containing 13 generic labels (`Col1` … `Col13`).
- Supports **multi‑column tables** (MCT) – automatically separates header groups by X‑coordinate gaps.
- Falls back to a content‑based search when column positions do not match perfectly.
- Extracts two types of structured identifiers (referred to as *item A* and *item B*) from the data rows.
- Aggregates unique identifiers and counts their occurrences.
- Exports two Excel sheets per file:
  - **Summary** – overall list of extracted identifiers.
  - **Details** – per‑record breakdown.
- Generates a combined summary across multiple files.

## How It Works

1. **Text extraction** – all `TEXT` and `MTEXT` entities are collected from the model space.
2. **Anchor search** – strings matching the `"number/number:Suffix"` pattern are treated as table anchors.
3. **Header detection** – immediately below each anchor, the first row of text is tested for the presence of all 13 generic header labels (ignoring spaces). If found, column positions are recorded.
4. **Multi‑column detection** – if the header row contains large gaps between groups of texts, the table is split into independent vertical sections, each with its own column map.
5. **Data row parsing** – for each row below the header, the value in the first column (`Col1`) is expected to be a sequential number. Two additional columns (`Col2`, `Col3`) are examined for structured identifiers (numeric‑hyphen patterns, etc.). When column positions are inaccurate, a dynamic fallback scans the entire row within the table boundaries.
6. **Aggregation** – found identifiers are counted and stored per record.
7. **Excel export** – the results are written to `.xlsx` files with formatting.

## Installation

Requires Python 3.10+ and the following packages:

```bash
pip install ezdxf openpyxl
```

## Usage

Place your DXF files in an input folder, then run:

```python
if __name__ == "__main__":
    input_folder = r"C:\path\to\dxf\files"
    output_folder = r"./output"
    process_folder(input_folder, output_folder)
```

Or use the command line after adjusting the paths inside the script.

## Output

For each processed file, an Excel workbook is created with two sheets:

- **Summary** – a consolidated BOM with columns: `ID`, `Description`, `Mfr`, `Qty`, `Type`.
- **Details** – a per‑record table with columns: `Position`, `Category`, `Description`, `Item ID`, `Mfr`, `Qty`.

A cross‑file summary is also generated, listing all unique identifiers, total quantities, and the files where they appear.

## Notes

- The parser assumes that tables have a left‑aligned anchor and a vertical numbering sequence in the first column.
- The header row must contain exactly 13 generic labels (or the sequence `1`, `2`, `3`… must be present to trigger fallback detection).
- Identifier formats are defined by regular expressions in `_ITEM_PATTERNS_A` and `_ITEM_PATTERNS_B`; they can be adjusted to match your specific drawing conventions.
- Multi‑column tables are automatically split when the gap between header texts exceeds an adaptive threshold.

## License

This project is provided as‑is for educational and adaptation purposes. Modify the pattern lists and logic to fit your own drawing standards.
```
