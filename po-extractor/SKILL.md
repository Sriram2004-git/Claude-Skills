---
name: po-extractor
description: >
  Extracts data from PO PDFs into the Tech Slice tracker at 'C:\Users\Durga Prakash\Documents\Tech Slice prjt\Business Database.xlsx'. 
  Supports multi-line items, duplicate detection, date normalization, and rigorous validation before writing.
---

# Purchase Order Extractor Skill

## Core Configuration
- **Default Path**: `C:\Users\Durga Prakash\Documents\Tech Slice prjt\Business Database.xlsx`
- **Dynamic Search**: If the file is missing or only a filename is given, recursively search `Documents`, `Desktop`, `Downloads`, and the user home folder.

---

## Field Mapping (Strict)

### Purchasedata Sheet
| Field | PO Extraction Key |
|---|---|
| Customer Name | `customer_name` |
| Customer PO Number | `po_number` |
| Customer Part Number | `line_item.part_number` |
| Order Date | `po_date` (Normalized to DD-MM-YYYY) |
| Order Qty | `line_item.qty` |
| Net Price | `line_item.net_price` |

### Sales Data Sheet
| Field | PO Extraction Key |
|---|---|
| Customer Name | `customer_name` |
| Customer PO Number | `po_number` |
| Customer PO Date | `po_date` (Normalized to DD-MM-YYYY) |
| Customer Part Number | `line_item.part_number` |
| Customer Order Qty | `line_item.qty` |
| Customer Unit Price | `line_item.unit_price` |

---

## Mandatory Workflows

### 1. Multi-line Item Handling
Every line item in a PO MUST be treated as a unique entry. Header fields (Customer Name, PO Number, Date) are repeated for every row generated from that PO.

### 2. Date Normalization
All dates MUST be converted to **DD-MM-YYYY** format (e.g., 06-01-2026). Never use "Jan", "January", or ISO formats in the final Excel output. Use a robust parser; if unparseable, flag for user correction.

### 3. Pre-Write Validation
Before writing ANY data, present the user with a validation table for EACH line item:
- **Status Table**: Show ✅ Found or ❌ Missing for every mapped field.
- **Action**: Ask the user: "Proceed with appending these items? (Y/N)"

### 4. Duplicate Detection
Search both sheets for the `Customer PO Number`. If found:
1. List matching PO Numbers and their row indices.
2. Present a menu:
   - `[1] Skip duplicates`
   - `[2] Append anyway`
   - `[3] Abort`
3. Wait for user input before proceeding.

### 5. Cell-Level Safety & Integrity
- **No Overwrites**: Before writing to a cell in the "next empty row", verify `cell.value is None`. If data exists, increment the row count and check again.
- **Isolated Update**: Do not touch, sort, or modify any existing records in the workbook.

### 6. Error Recovery
If a PDF fails to read mid-batch:
- Present the error.
- Ask: "Skip this file and continue batch? OR Abort entire process?"

### 7. Batch Summary Report
After completion, display a final table:
| PO File | PO Number | Items Added | Missing Fields | Status |
|---|---|---|---|---|
| toshiba.pdf | 76275 | 4 | None | Success |

---

## Python Logic Template (Internal Reference)

```python
import openpyxl
from datetime import datetime
import os

def normalize_date(date_str):
    # Try multiple formats (D-M-Y, M/D/Y, D-MMM-Y, etc.)
    # ALWAYS return DD-MM-YYYY
    pass

def find_workbook(filename):
    # Recursive search in Documents, Desktop, Downloads
    pass

def check_duplicates(ws, po_number):
    # Returns list of row indices matching the PO
    pass

def safe_append(ws, row_data, headers_map):
    # Check cell.value is None before writing
    pass
```

## Required Libraries
`pip install pdfplumber openpyxl python-dateutil --break-system-packages`
