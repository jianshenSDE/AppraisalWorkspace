================================================================================
  IMPROVED COMPARABLE (ImprovedComp) — HOW TO GENERATE A NEW REPORT
================================================================================

WHAT THIS GENERATES
-------------------
  1. An Excel spreadsheet with all 38 fields per comparable sale
  2. A Word document with the data table + property photos from the MLS listing


BEFORE YOU START — GATHER THESE FILES
--------------------------------------

  FILE 1:  MLS PDF  (required)
           Export the MLS listing from NTREIS, HAR, or other MLS system.
           Must include: listing detail, sale history, building characteristics.
           Save to:  c:\AppraisalWorkspace\Sources\MLS\

  FILE 2:  CAD PDF  (strongly recommended)
           Download the property detail page from the County Appraisal District
           website (e.g. kaufmancad.org, dcad.org, tcad.org).
           Provides: building SF, year built, land dimensions, county.
           Save to:  c:\AppraisalWorkspace\Sources\CAD\

           How to get it:
             1. Go to the CAD website for the county
             2. Search by address or property ID
             3. Open the property detail page
             4. Print / Save as PDF

  FILE 3:  TexasFile Screenshot  (optional — alternative to CAD for deed lookup)
           Save to:  c:\AppraisalWorkspace\Sources\TexasFile\

  For multiple comparable sales, provide one MLS PDF and one CAD PDF per property.


HOW TO GENERATE THE REPORT
---------------------------

  Copy a prompt below, paste it into the AI chat (Windsurf/Cascade),
  swap in your actual file names, and send it. That's it!


----------------------------------------------------------------
  OPTION A — SINGLE PROPERTY (MLS + CAD)
----------------------------------------------------------------

  Generate an ImprovedComp report for Terrell.

  MLS file: Sources/MLS/808 E Moore Avenue, Terrell MLS.pdf
  CAD file: Sources/CAD/808 E Moore Avenue, Terrell Kaufmann CAD.pdf


----------------------------------------------------------------
  OPTION B — MULTIPLE PROPERTIES
----------------------------------------------------------------

  Generate an ImprovedComp report for DFW with 3 comparable sales.

  MLS files:
    Sources/MLS/MLS-Property1.pdf
    Sources/MLS/MLS-Property2.pdf
    Sources/MLS/MLS-Property3.pdf

  CAD files:
    Sources/CAD/Property1-CAD.pdf
    Sources/CAD/Property2-CAD.pdf
    Sources/CAD/Property3-CAD.pdf


----------------------------------------------------------------
  OPTION C — MLS ONLY (no CAD available)
----------------------------------------------------------------

  Generate an ImprovedComp report for Keller.

  MLS file: Sources/MLS/MLS-Keller.pdf
  (No CAD PDF available)


WHERE TO FIND YOUR OUTPUT
--------------------------
  After the AI runs, your files will appear here:

      c:\AppraisalWorkspace\ImprovedComp\Output\

      ImprovedComp_[Name].xlsx   <-- Excel spreadsheet (all 38 fields)
      ImprovedComp_[Name].docx   <-- Word doc with table + photos


AFTER GENERATION — FILL IN THESE FIELDS MANUALLY
--------------------------------------------------
  The script fills in everything it can confidently extract.
  The following fields require analyst judgment or additional data:

  +----------------------------+--------------------------------------------------+
  | Field                      | What to fill in                                  |
  +----------------------------+--------------------------------------------------+
  | Property Type              | Verify auto-detected type (e.g. "Restaurant",    |
  |                            | "Retail Strip", "Office Building")               |
  +----------------------------+--------------------------------------------------+
  | Rentable Unit Number       | Number of suites/units if multi-tenant           |
  | Average Unit Size          | Average SF per unit if multi-tenant              |
  | Unit Mix                   | Description of suite types if multi-tenant       |
  +----------------------------+--------------------------------------------------+
  | Project Amenities          | Parking, signage, monument, drive-thru, etc.     |
  | Unit Amenities             | HVAC type, loading dock, grade level doors, etc. |
  +----------------------------+--------------------------------------------------+
  | Finish-out Percentage      | % of space that is finished/leaseable            |
  |                            | (script defaults to 100% if tenant in place)     |
  +----------------------------+--------------------------------------------------+
  | Cap Rate                   | If NOI / income data is available                |
  | NOI ($/Net SF)             | Net Operating Income per SF                      |
  | NOI ($/Unit)               | NOI per unit if multi-tenant                     |
  +----------------------------+--------------------------------------------------+
  | Recorded Number            | If not auto-extracted, look up at TexasFile.com  |
  | Grantee                    | If not auto-extracted                            |
  +----------------------------+--------------------------------------------------+
  | Additional Comments        | Expand the auto-generated text with relevant     |
  |                            | details about condition, tenancy, renovations    |
  +----------------------------+--------------------------------------------------+


FIELDS FILLED AUTOMATICALLY
----------------------------
  The following are extracted or calculated without any manual work:

  - Comparable Sale, CAD ID, Street Address, City, State, County
  - Land Size (SF and Acres), Gross Building Area, Net Rentable Area
  - Land to Building Ratio, Year Built
  - Date of Sale, Sales Price, Original Listing Price
  - Unit Price ($/Net SF and $/Gross SF), Sale-to-List Ratio
  - Occupancy Rate at Time of Sale (from MLS description)
  - Data Source, Grantor, Terms and Conditions
  - Time on Market, Property Rights Conveyed, Transactional Status
  - Property Type (keyword detection — verify accuracy)
  - Additional Comments (starter text from MLS description)

================================================================================
