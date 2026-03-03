LEASE COMPARABLE GENERATOR
==========================

Parses a Collier lease comps PDF and generates:
  - LeaseComp_[name].xlsx   (comparison spreadsheet matching example format)
  - LeaseComp_[name].docx   (Word doc with data tables + PDF page images)


QUICK START
-----------
    cd c:\AppraisalWorkspace\LeaseComps
    python generate_lease_comp.py

Default input:  Sources\Comps\lease cstore comps - collier.pdf
Default output: LeaseComps\Output\LeaseComp_Collier_CStore.xlsx / .docx


USAGE
-----
    python generate_lease_comp.py [--name <name>] [--pdf <path>]

    --name      Output name used in filenames (default: Collier_CStore)
    --pdf       Path to the Collier lease comps PDF


EXAMPLES
--------
Default (uses CONFIG paths):
    python generate_lease_comp.py

Custom PDF and name:
    python generate_lease_comp.py --name "Valero_CC" --pdf "C:\Sources\Comps\lease_comps.pdf"


REQUIREMENTS
------------
Python 3.8+ with:
    pip install python-docx PyMuPDF Pillow openpyxl


EXCEL FORMAT (matches example LeaseCompWriteUpExcel.xlsx)
---------------------------------------------------------
Fields per comparable (25 rows):
    Rent Comp, Status, Property Name, Location, City, State,
    Net Rentable Area (SF), Estimated Year of Construction,
    Rentable Unit Size (SF), Tenant, Data Source,
    Lease Rate ($/SF/YR), Adj Lease Rate ($/SF/YR),
    Leased Date, Lease Start, Lease Term (months), Lease End,
    Months on Market, Months Vacant, Expense Structure,
    Building Type, Land Size (SF), Land Size (Acre),
    Additional Comments

IMPORTANT — Lease Rate vs. Adjusted Lease Rate:
    The Collier PDF contains TWO rate columns:
      - LEASE RATE:     The actual contract rent per SF per year.
      - ADJ LEASE RATE: An appraiser-adjusted/normalized rate used for
                        comparison analysis (accounts for differences in
                        lease terms, concessions, market conditions, etc.).
    The script populates "Lease Rate ($/SF/YR)" with the CONTRACT rate
    (the actual rent), and includes "Adj Lease Rate ($/SF/YR)" as a
    separate field for reference. Do NOT confuse the two — the adjusted
    rate does not reflect what the tenant actually pays.

Fields left blank for analyst to fill:
    Leased Date, Lease Start, Lease End, Months on Market, Months Vacant


WORD DOC FORMAT
---------------
Same layout as LandComp / ImprovedComp Word docs:
  - "Lease Comparable No. X" heading (12pt Bold Calibri)
  - Two-column data table (Table Grid style, 9pt Calibri)
  - PDF page images rendered as compressed JPEG (200 DPI, quality 85)
    embedded at 6.5" width


HOW IT WORKS
------------
1. Opens the Collier lease comps PDF and splits each page into individual
   comparable blocks (typically 2 per page).

2. Parses structured fields: Physical Information (Name, Address, NRA,
   Year Built, Occupancy, Site Size, Construction), Confirmation source,
   and the lease data table (Tenant, Rate Type, Size, Term, Lease Rate,
   Adjusted Lease Rate).

3. Maps parsed fields to the 25-field LeaseComp structure matching the
   example Excel format. Uses the contract LEASE RATE (not ADJ LEASE RATE)
   for the primary rate field; the adjusted rate is included separately.

4. Generates Excel with the same styling as ImprovedComp (blue headers,
   alternating row fills, frozen panes).

5. Generates Word doc from a blank Document() — no template loaded as base.
   PDF pages are rendered as compressed JPEG images and embedded after the
   data tables.

IMPORTANT — Template handling:
   The Word doc is built from Document() (blank). This avoids orphaned image
   blobs that would bloat the file and cause Word to crash on copy/paste.
   Do NOT change this to Document(template_path) or shutil.copy() followed
   by clearing paragraphs. See RetailMarketReport/README.txt for details.


SOURCE PDF FORMAT
-----------------
Colliers International Valuation & Advisory Services lease comparable pages.
Each page contains 2 comparables with this structure:
  - COMPARABLE N header
  - PHYSICAL INFORMATION block (Name, Address, City/State/Zip, MSA, NRA,
    Year Built, Occupancy, Site Size, Site Coverage, Construction)
  - CONFIRMATION block (Company, Source)
  - Lease table: TENANT NAME / RATE TYPE / SIZE / START DATE / TERM /
    LEASE RATE / ADJ LEASE RATE
  - Remarks sentence


FILE STRUCTURE
--------------
LeaseComps/
  generate_lease_comp.py    -- Main generation script
  README.txt                -- This file
  Example/
    LeaseCompWriteUpExcel.xlsx  -- Example Excel format reference
  Output/
    LeaseComp_Collier_CStore.xlsx
    LeaseComp_Collier_CStore.docx
