# Improved Comparable (ImprovedComp) — Creation Instructions

## Purpose
Generate an Improved Comparable Sale spreadsheet and matching Word document for appraisal assignments.
Each comparable = one MLS PDF + one CAD PDF (recommended) → one column in the Excel / one section in the Word doc.

---

## Folder Structure

```
c:\AppraisalWorkspace\
├── Sources/
│   ├── MLS/          # MLS PDF exports (one per comparable sale)
│   ├── CAD/          # County Appraisal District PDFs (one per comparable, recommended)
│   └── TexasFile/    # TexasFile.com screenshots (optional alternative to CAD)
└── ImprovedComp/
    ├── instructions.md               # This file
    ├── generate_improved_comp.py     # Generator script
    ├── README.txt
    ├── Example/
    │   ├── Improved comp example.xlsx
    │   └── improved comp example.docx
    └── Output/                       # Generated files go here
```

---

## Field Mapping — All 38 Fields

| # | Field | Source | Notes |
|---|---|---|---|
| 1 | Comparable Sale | Sequential ("Sale No. 1", ...) | Auto-assigned |
| 2 | Property Type | MLS description keywords | "Restaurant", "Retail", "Office", etc. — human verify |
| 3 | CAD ID | CAD: `Property ID:` or MLS: `Parcel ID:` | CAD preferred |
| 4 | Street Address | MLS: first address line / CAD: `Situs Address:` | MLS preferred |
| 5 | City | MLS: first address line / CAD: `Situs Address:` | |
| 6 | State | Always "Texas" | |
| 7 | County | CAD: entity table (`KAUFMAN COUNTY` → "Kaufman") | MLS fallback |
| 8 | Land Size (SF) | CAD: Property Land table sqft / MLS: `Lot SqFt:` | CAD preferred (more precise) |
| 9 | Land Size (Acres) | **Calculated**: Land SF ÷ 43,560 | |
| 10 | Gross Building Area (SF) | CAD: `COMMERCIAL MAIN` sqft / MLS: `Building Sq Ft:` | |
| 11 | Net Rentable Area (SF) | Same as Gross for single-tenant; analyst adjusts for multi-tenant | |
| 12 | Rentable Unit Number | **BLANK** — analyst fills for multi-unit | |
| 13 | Average Unit Size | **BLANK** — analyst fills for multi-unit | |
| 14 | Unit Mix | **BLANK** — analyst fills for multi-unit | |
| 15 | Project Amenities | **BLANK** — analyst fills | |
| 16 | Unit Amenities | **BLANK** — analyst fills | |
| 17 | Finish-out Percentage | Default 100% if occupied/leased | Analyst may adjust |
| 18 | Land to Building Ratio | **Calculated**: Land SF ÷ Gross Bldg SF | |
| 19 | Year Built/Renovated | CAD: building table `Year Built` | |
| 20 | Date of Sale | MLS: `Closed Date:` | |
| 21 | Sales Price | MLS: `Close Price:` | |
| 22 | Unit Price ($/Net SF) | **Calculated**: Sales Price ÷ Net Rentable SF | |
| 23 | Unit Price ($/Gross SF) | **Calculated**: Sales Price ÷ Gross Building SF | |
| 24 | Unit Price ($/Unit) | **BLANK** — analyst fills for multi-unit | |
| 25 | Cap Rate | **BLANK** — requires NOI data not in MLS | |
| 26 | NOI ($/Net SF) | **BLANK** — requires income data | |
| 27 | NOI ($/Unit) | **BLANK** — analyst fills for multi-unit | |
| 28 | Occupancy Rate at Time of Sale | MLS description: "currently leased" / "vacant" → 100% / 0% | Analyst verify |
| 29 | Data Source | `[MLS Name] [MLS#], [CAD Name]` or `Texasfile.com` | |
| 30 | Recorded Number | MLS Sale History: `Document #` | May be volume-page format (e.g. 8178-232) |
| 31 | Grantor | MLS Sale History: Seller Name(s) | All names included |
| 32 | Grantee | MLS Sale History: Buyer Name(s) | All names included |
| 33 | Terms and Conditions | MLS: `Buyer Financing:` | "Cash" → "Cash or conventional financing" |
| 34 | Time on Market | MLS: `CDOM:` preferred, fallback `DOM:` | Days |
| 35 | Property Rights Conveyed | Default "Fee Simple" | Standard for Warranty Deed |
| 36 | Transactional Status | MLS Closed status → "Sold" | |
| 37 | Original Listing Price | MLS: `OLP:` | |
| 38 | Sale-to-List Ratio | **Calculated**: Sales Price ÷ OLP | |
| 39 | Additional Comments | From MLS property description (truncated) | Human should expand |

> **Rule: If any field cannot be reliably extracted, leave it blank. A human will fill it in manually.**

---

## What CAD PDF Provides (vs MLS)

| Field | MLS | CAD | Priority |
|---|---|---|---|
| CAD ID / Property ID | `Parcel ID:` | `Property ID:` | CAD |
| Land SF | `Lot SqFt:` | Property Land table | CAD (more precise) |
| Land Acres | `Acres:` | Property Land table | CAD |
| County | `County:` | Entity table (e.g. KAUFMAN COUNTY) | CAD |
| Gross Building Area | `Building Sq Ft:` | `COMMERCIAL MAIN [sqft]` | CAD |
| Year Built | (not usually in MLS) | Building table `Year Built` | CAD |
| Building Type / Use | MLS description | Building type classification | Both |

---

## Property Type Detection Keywords

The script attempts to detect property type from MLS description text:

| Keywords in description | Property Type assigned |
|---|---|
| taqueria, restaurant, bar, grill, tavern, diner | Restaurant |
| retail, store, shop, boutique | Retail |
| office, professional, suite | Office |
| warehouse, storage, industrial, flex | Industrial/Warehouse |
| medical, clinic, dental, health | Medical Office |
| hotel, motel, hospitality, inn | Hospitality |
| church, worship, ministry | Religious |
| auto, car wash, tire, mechanic | Automotive |

> If no keyword matches, the property type is left **blank** for the analyst.

---

## Step-by-Step Process

### Step 1 — Gather Source Files

**MLS PDF** (required)
- Export from NTREIS, HAR, or other MLS system
- Must include: listing detail, sale history, building characteristics
- Save to `c:\AppraisalWorkspace\Sources\MLS\`

**CAD PDF** (strongly recommended)
- Download from the County Appraisal District website (e.g. Kaufman CAD, DCAD, TCAD)
  - Search the property at the CAD website → "Print" or "Export" the property detail page
- Provides building SF, year built, land dimensions, county
- Save to `c:\AppraisalWorkspace\Sources\CAD\`

**TexasFile PNG** (alternative to CAD, for Recorded Number cross-reference)
- Save to `c:\AppraisalWorkspace\Sources\TexasFile\`

### Step 2 — Configure the Script

Open `generate_improved_comp.py` and update the `CONFIG` section:
```python
MLS_PDF_PATHS = [
    r"c:\AppraisalWorkspace\Sources\MLS\808 E Moore Avenue, Terrell MLS.pdf",
]
CAD_PDF_PATHS = [
    r"c:\AppraisalWorkspace\Sources\CAD\808 E Moore Avenue, Terrell Kaufmann CAD.pdf",
]
OUTPUT_NAME = "Terrell"
```

### Step 3 — Run the Script

**Edit CONFIG then run:**
```
cd c:\AppraisalWorkspace\ImprovedComp
python generate_improved_comp.py
```

**Or use command-line arguments:**
```
python generate_improved_comp.py --name "Terrell" --mls "Sources/MLS/808 E Moore Avenue, Terrell MLS.pdf" --cad "Sources/CAD/808 E Moore Avenue, Terrell Kaufmann CAD.pdf"
```

For multiple comparables:
```
python generate_improved_comp.py --name "DFW" --mls "MLS1.pdf" --mls "MLS2.pdf" --cad "CAD1.pdf" --cad "CAD2.pdf"
```

### Step 4 — Review Output

Files appear in `c:\AppraisalWorkspace\ImprovedComp\Output\`:
- `ImprovedComp_[Name].xlsx` — full 38-field spreadsheet
- `ImprovedComp_[Name].docx` — Word doc with data table + MLS photos

### Step 5 — Fill Blank Fields Manually

| Field | What to fill in |
|---|---|
| Property Type | Verify/refine the auto-detected type (e.g. "Restaurant", "Retail Strip") |
| Rentable Unit Number | Number of units/suites if multi-tenant |
| Average Unit Size | SF per unit if multi-tenant |
| Unit Mix | Description of suite types if multi-tenant |
| Project Amenities | Parking, signage, etc. |
| Unit Amenities | HVAC, loading dock, drive-thru, etc. |
| Finish-out Percentage | % of space that is finished/leaseable |
| Cap Rate | If income data is available |
| NOI ($/Net SF) | If income data is available |
| Recorded Number | If not auto-extracted, look up at TexasFile.com or CAD |
| Grantee | If not auto-extracted |
| Additional Comments | Expand the auto-generated summary with relevant details |

---

## Outputs

### Excel Spreadsheet
- **Columns**: A = blank, B = field labels, C = Sale No. 1, D = Sale No. 2, etc.
- **Formulas**: Land Size (Acres), Land-to-Building Ratio, Unit Prices, Sale-to-List Ratio all use cell formulas
- **Formatting**: bold headers, borders, label background shading

### Word Document
- **Per comparable**: "Improved Comparable Sale No. X" heading → data table → property photos
- **Data table**: shows only fields with non-blank values (blank fields are Excel-only)
- **Photos**: large MLS listing images from page 2, displayed in 2-column grid
- **Multiple comparables**: stack sequentially

---

## Dependencies

| Package | Install |
|---|---|
| `python-docx` | `pip install python-docx` |
| `openpyxl` | `pip install openpyxl` |
| `pymupdf` | `pip install pymupdf` |
