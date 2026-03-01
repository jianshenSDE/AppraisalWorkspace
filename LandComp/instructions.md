# Land Comparable (LandComp) — Creation Instructions

## Purpose
Generate a Land Comparison Excel spreadsheet and matching Word document for appraisal assignments.
Each comparable sale = one MLS PDF + one optional TexasFile PNG → one column in the Excel / one section in the Word doc.

---

## Folder Structure

```
c:\AppraisalWorkspace\
├── Sources/
│   ├── MLS/                          # MLS PDF exports (one per comparable)
│   │   └── MLS-Corsicana, Texas, 75110.pdf
│   └── TexasFile/                    # TexasFile.com screenshots (one per comparable)
│       └── texasfile-navarroSearch.png
└── LandComp/
    ├── instructions.md               # This file
    ├── generate_land_comp.py         # Generator script
    ├── Examples/
    │   ├── LandComparisonExample.xlsx
    │   └── Land Comp write up example.docx
    └── Output/                       # Generated files go here
```

---

## Field Mapping — All 33 Fields

| # | Field | Source | Notes |
|---|---|---|---|
| 1 | Comparable Sale | Sequential ("Sale No. 1", "Sale No. 2", ...) | Auto-assigned |
| 2 | Property Type | MLS SubType | "Unimproved Land" → "Land Site" |
| 3 | CAD ID | MLS: `Parcel ID:` | |
| 4 | Street Address | MLS: first address line of listing | |
| 5 | City | MLS: first address line of listing | |
| 6 | State | MLS: first address line of listing | Always "Texas" |
| 7 | County | MLS: `County:` field | |
| 8 | Land Size (SF) | MLS: `Lot SqFt:` | Numeric only |
| 9 | Land Size (Acres) | MLS: `Acres:` | |
| 10 | Sales Price | MLS: `Close Price:` | |
| 11 | Date of Sale | MLS: `Closed Date:` | |
| 12 | Unit Price ($/SF) | **Calculated**: Close Price ÷ Lot SF | |
| 13 | Unit Price ($/Acre) | **Calculated**: Close Price × 43,560 ÷ Lot SF | Uses SF for accuracy |
| 14 | Zoning | MLS: `Zoning:` | Formatted with abbreviation code |
| 15 | Configuration | Default: "Regular" | Human override if irregular |
| 16 | Topography | **BLANK — Human fills** | Derived from site visit / Google Maps description |
| 17 | Min Topo (Ft) | **BLANK — Human fills** | From Google Maps elevation |
| 18 | Max Topo (Ft) | **BLANK — Human fills** | From Google Maps elevation |
| 19 | Topo % Change | **BLANK — Human fills** | Calculated: (Max−Min) ÷ Min |
| 20 | Access | Default: "Inside" | Human override if corner lot |
| 21 | Utilities | MLS: `Street/Utilities:` | Standardized ("All available", "Available", etc.) |
| 22 | Flood Zone | MLS: `Flood Zone Code:` | From page 4 of MLS PDF |
| 23 | Data Source | `[MLS Name] [MLS#], Texasfile.com` | MLS name detected from PDF copyright |
| 24 | Recorded Number | MLS Sale History: `Document #` | Also visible in TexasFile PNG |
| 25 | Grantor | MLS Sale History: Seller Name(s) | Reformatted from "Last First" to "First Last" |
| 26 | Grantee | MLS Sale History: Buyer Name(s) | Reformatted from "Last First" to "First Last" |
| 27 | Terms and Conditions | MLS: `Buyer Financing:` | "Cash" → "Cash or conventional financing" |
| 28 | Days on Market | MLS: `CDOM:` or `DOM:` | CDOM preferred |
| 29 | Property Rights Conveyed | Default: "Fee Simple" | Standard for Warranty Deed |
| 30 | Transactional Status | MLS Closed status → "Sold" | |
| 31 | Original Listing Price | MLS: `OLP:` | |
| 32 | Sale-to-List Ratio | **Calculated**: Close Price ÷ OLP | |
| 33 | Additional Comments | Default: "None." | Human override |

> **Rule: If any field cannot be reliably extracted, leave it blank. A human will fill it in manually.**

---

## What TexasFile PNG Provides

The TexasFile screenshot confirms/cross-references:
- **Recorded Number** (Document #) — same as in MLS Sale History
- **Grantor** and **Grantee** — same as in MLS Sale History
- **Document Type** (Warranty Deed) — confirms Fee Simple rights
- **Date Filed** — county recording date (may differ from closing date)

Since all of this data is also available in the MLS PDF's "Sale History from Public Records" section, the TexasFile PNG is used as **secondary verification** and for the `Data Source` field notation.

---

## Step-by-Step Process

### Step 1 — Gather Source Files

**MLS PDF(s)** — export from the MLS system (NTREIS, HAR, etc.)
- Must include: listing details, sale history, flood zone pages
- Save to `c:\AppraisalWorkspace\Sources\MLS\`
- **One property per PDF** — provide one file per comparable, OR
- **Multiple properties in one PDF** — export a bulk/combined PDF; the script auto-detects and splits each listing by its MLS# + address header. One combined PDF for all comparables is fine.

**TexasFile PNG(s)** — screenshot of TexasFile.com search results showing deed records
- Save to `c:\AppraisalWorkspace\Sources\TexasFile\`
- **One screenshot can cover multiple properties** — a single PNG showing all deed records in the search table is sufficient. The script uses it only for the `Data Source` field notation (`Texasfile.com`), not for per-property data extraction.

### Step 2 — Configure the Script
Open `generate_land_comp.py` and update the `CONFIG` section at the top:
```python
MLS_PDF_PATHS = [
    r"c:\AppraisalWorkspace\Sources\MLS\MLS-NewCity.pdf",
    # Each entry can be a single-listing OR multi-listing PDF.
    # Add multiple entries here if you have separate PDFs per property.
]
TEXASFILE_PNG_PATHS = [
    r"c:\AppraisalWorkspace\Sources\TexasFile\texasfile-newSearch.png",
    # optional — one PNG covering all properties is fine. Can be empty list [].
]
OUTPUT_NAME = "NewCity"   # used in output filenames
```

### Step 3 — Run the Script

**Option A — Edit CONFIG and run (no arguments needed):**
```
cd c:\AppraisalWorkspace\LandComp
python generate_land_comp.py
```

**Option B — Pass file paths directly as arguments (no file editing needed):**
```
cd c:\AppraisalWorkspace\LandComp
python generate_land_comp.py --name "CityName" --mls "c:\AppraisalWorkspace\Sources\MLS\MLS-File.pdf" --texasfile "c:\AppraisalWorkspace\Sources\TexasFile\tf-search.png"
```

For multiple comparables, repeat `--mls` for each one:
```
python generate_land_comp.py --name "Austin" --mls "Sources\MLS\MLS-Prop1.pdf" --mls "Sources\MLS\MLS-Prop2.pdf"
```

**Option C — Ask the AI assistant (Cascade/Windsurf):**
Simply tell it: *"Generate a LandComp for [location] using `Sources/MLS/[filename].pdf` and `Sources/TexasFile/[filename].png`"*
and it will update the CONFIG and run the script for you.

### Step 4 — Review Output
Output files appear in `c:\AppraisalWorkspace\LandComp\Output\`:
- `LandComp_NewCity.xlsx` — comparison spreadsheet
- `LandComp_NewCity.docx` — Word doc with tables + MLS photos

### Step 5 — Fill Blank Fields Manually
In the generated files, manually fill in:
- `Topography` — describe site terrain (e.g., "Level", "Gently sloping", "Rolling")
- `Min Topo (Ft)` — use Google Maps elevation tool
- `Max Topo (Ft)` — use Google Maps elevation tool
- `Topo % Change` — calculate: `(Max - Min) / Min × 100%`
- `Configuration` — override if lot is irregular, triangular, etc.
- `Access` — override if corner lot ("Corner") or other access type
- `Additional Comments` — add any relevant notes

---

## Outputs

### Excel Spreadsheet
- **Columns**: A = field labels, B = Sale No. 1, C = Sale No. 2, etc.
- **Formatting**: bold headers, borders, alternating row shading
- **Calculated fields**: Unit Price $/SF, Unit Price $/Acre, Sale-to-List Ratio

### Word Document
- **Per comparable**: "Land Sale Comp No. X" heading → data table → property photos
- **Photos**: large MLS listing images extracted from page 2 of MLS PDF, displayed in 2-column grid
- **Multiple comparables**: stack sequentially (Comp 1 table + photos → Comp 2 table + photos → ...)

---

## Dependencies

| Package | Install | Purpose |
|---|---|---|
| `python-docx` | `pip install python-docx` | Word document generation |
| `openpyxl` | `pip install openpyxl` | Excel generation |
| `pymupdf` | `pip install pymupdf` | PDF text + image extraction |

---

## Zoning Code Reference

| MLS Zoning Text | Formatted Output |
|---|---|
| General Retail | GR (General Retail) |
| Commercial | C (Commercial) |
| Light Industrial | LI (Light Industrial) |
| Heavy Industrial | HI (Heavy Industrial) |
| Residential | R (Residential) |
| Agricultural | A (Agricultural) |
| Mixed Use | MU (Mixed Use) |
| Office | O (Office) |
| Planned Development | PD (Planned Development) |

> If the zoning code is not in this table, the raw MLS value is used. Human should verify and abbreviate.
