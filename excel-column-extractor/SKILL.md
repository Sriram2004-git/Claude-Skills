---
name: excel-column-extractor
description: Extracts specific product master data columns from an Excel file, excludes certain fields, enriches the data by matching with a Price List Excel file, and maps the result to a final Excel template with default values.
---

# Excel Column Extractor - Product Master Processing & Final Template Mapping

This skill automates the extraction, enrichment, and transformation of Product Master data into a final import-ready Excel template.

## Workflow

### NOTE - Don't search the Product Master and Price list excel by yourself
Always ask the user for path of the Product Master, Price List excel and Destination folder location. 

### Stage 1 – Product Master Extraction (In-Memory)
Only extract the following required fields from the Product Master source to minimize memory usage and skip disk writing:
* Material ID, Ordering Code Packed, MPG Product Description

### Stage 2 – Price List Comparison & Data Enrichment (In-Memory)
Match **Product Master (Material ID)** with **Price List (Material_NR)**. Retrieve only:
* Market_List_Price_Value, HSN, Techslice Price.

### Stage 3 – Final Excel Template Generation & Splitting
Transform the enriched dataset into the required 31-column template structure with specific mappings and default values. Save output directly in the destination folder without writing intermediate files.

### Indicator
Indicate the process while extraction happening in the background don't just give small message snippet, give simple live dashboard and at final give a summary messages in the console, showing the number of records processed and any errors encountered.

## Column Mapping Rules

| Template Column      | Source Value            |
| -------------------- | ----------------------- |
| Item Name            | Ordering Code Packed    |
| SKU                  | Material ID             |
| HSN/SAC              | HSN                     |
| Description          | MPG Product Description |
| Rate                 | Market_List_Price_Value |
| MRP                  | Market_List_Price_Value |
| Purchase Description | MPG Product Description |
| Purchase Rate        | Techslice Price         |

## Default Values

| Template Column      | Default Value                          |
| -------------------- | -------------------------------------- |
| Product Type         | goods                                  |
| Account              | Sales                                  |
| Usage unit           | pcs                                    |
| Item Type            | Sales                                  |
| Purchase Account     | Cost of Goods Sold                     |
| Inventory Account    | Inventory Asset                        |
| Vendor               | SANDVIK COROMANT INDIA PRIVATE LIMITED |
| Vendor Number        | 12025001                               |
| Status               | Active                                 |
| Inter/Intra State Tax| GST18/IGST18 (18%)                     |

> **Note – No-Overwrite Policy**: The final output is always saved as new files.
> * If total records <= 25,000, a single file is saved using a month-year name (e.g., `Product_Master_Final_June-2026.xlsx`). If a file with that name already exists, a numeric suffix is appended (e.g., `_1`, `_2`, …).
> * If total records > 25,000, the dataset is split into multiple files (each of the first files containing exactly 25,000 records, and the final file containing all remaining records). These are named using the Part-N convention (e.g., `Product_Master_Final_Part-1-June-2026.xlsx`). If any files in a part series already exist, a numeric suffix is appended to all files in the generated set (e.g., `Product_Master_Final_Part-1-June-2026_1.xlsx`).
> Source files and previously generated outputs are never modified or replaced.

## Usage

### Process, Enrich, & Transform (Stage 1, 2, & 3 Combined)
```powershell
python "<SKILL_PATH>/scripts/process_product_master.py" "<SOURCE_FILE_PATH>" "<PRICE_LIST_PATH>" "<DESTINATION_FOLDER_PATH>"
```

## Dependencies
Requires `pandas` and `openpyxl`.
