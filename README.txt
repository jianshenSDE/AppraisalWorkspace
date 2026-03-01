================================================================================
 APPRAISAL WORKSPACE — README
================================================================================

FOLDER STRUCTURE
----------------
c:\AppraisalWorkspace\
  Sources\
    MLS\              <-- Drop MLS PDF exports here (one per comparable sale)
    CAD\              <-- Drop County Appraisal District PDFs here
    TexasFile\        <-- Drop TexasFile.com screenshots here (one per comparable)
    WDA\              <-- Drop TWC WDA/MSA Profile PDFs here (one per region)
  LandComp\           <-- Land Comparable report generator (vacant/unimproved land)
  ImprovedComp\       <-- Improved Comparable report generator (buildings/commercial)
  RegionalOverviewAndMarketAnalysis\  <-- Regional Overview & Market Area report generator


================================================================================
 REPORT TYPE 1: LAND COMPARABLE (LandComp)
================================================================================

FILES YOU NEED TO PROVIDE
--------------------------
1. MLS PDF (required, one per comparable sale)
   - Export from NTREIS, HAR, or other MLS system
   - Must include: listing detail page, sale history, flood zone pages
   - Save to: c:\AppraisalWorkspace\Sources\MLS\

2. TexasFile PNG (optional but recommended, one per comparable)
   - Screenshot of TexasFile.com search results showing the deed record
   - Save to: c:\AppraisalWorkspace\Sources\TexasFile\

OUTPUTS GENERATED
-----------------
- LandComp\Output\LandComp_[Name].xlsx   (comparison spreadsheet, all 33 fields)
- LandComp\Output\LandComp_[Name].docx   (Word doc with data table + MLS photos)

FIELDS LEFT BLANK (fill manually after generation)
----------------------------------------------------
- Topography         -- describe terrain from site visit (e.g. "Level", "Gently sloping")
- Min Topo (Ft)      -- look up on Google Maps elevation tool
- Max Topo (Ft)      -- look up on Google Maps elevation tool
- Topo % Change      -- calculated: (Max - Min) / Min * 100%

SAMPLE PROMPTS
--------------

Single comparable:
  "Generate a LandComp for Corsicana using
   Sources/MLS/MLS-Corsicana, Texas, 75110.pdf
   and Sources/TexasFile/texasfile-navarroSearch.png"

Multiple comparables:
  "Generate a LandComp for Waco using these files:
   MLS PDFs:
     Sources/MLS/MLS-Waco-Prop1.pdf
     Sources/MLS/MLS-Waco-Prop2.pdf
     Sources/MLS/MLS-Waco-Prop3.pdf
   TexasFile PNGs:
     Sources/TexasFile/texasfile-mclennan-1.png
     Sources/TexasFile/texasfile-mclennan-2.png
     Sources/TexasFile/texasfile-mclennan-3.png"

No TexasFile (MLS only):
  "Generate a LandComp for Austin using Sources/MLS/MLS-Austin.pdf
   (no TexasFile available)"


================================================================================
 REPORT TYPE 2: IMPROVED COMPARABLE (ImprovedComp)
================================================================================

Use this for commercial properties that have buildings/improvements on them
(restaurants, retail, office, warehouse, etc.).
Use LandComp (above) for vacant/unimproved land sales.

FILES YOU NEED TO PROVIDE
--------------------------
1. MLS PDF (required, one per comparable sale)
   - Export from NTREIS, HAR, or other MLS system
   - Save to: c:\AppraisalWorkspace\Sources\MLS\

2. CAD PDF (strongly recommended, one per comparable)
   - Download from the County Appraisal District website
   - Provides: building SF, year built, land dimensions, county
   - Save to: c:\AppraisalWorkspace\Sources\CAD\

OUTPUTS GENERATED
-----------------
- ImprovedComp\Output\ImprovedComp_[Name].xlsx   (38-field spreadsheet)
- ImprovedComp\Output\ImprovedComp_[Name].docx   (Word doc with table + photos)

FIELDS LEFT BLANK (fill manually after generation)
----------------------------------------------------
- Rentable Unit Number / Average Unit Size / Unit Mix  -- for multi-unit buildings
- Project Amenities / Unit Amenities                   -- from site inspection
- Cap Rate / NOI                                       -- requires income/rent data
- Recorded Number / Grantee                            -- if not in MLS sale history
- Additional Comments                                  -- expand starter text

SAMPLE PROMPTS
--------------

Single comparable (MLS + CAD):
  "Generate an ImprovedComp report for Terrell using
   Sources/MLS/808 E Moore Avenue, Terrell MLS.pdf
   and Sources/CAD/808 E Moore Avenue, Terrell Kaufmann CAD.pdf"

Multiple comparables:
  "Generate an ImprovedComp report for DFW with 3 comparable sales.
   MLS files: Sources/MLS/Prop1.pdf, Sources/MLS/Prop2.pdf, Sources/MLS/Prop3.pdf
   CAD files: Sources/CAD/CAD1.pdf, Sources/CAD/CAD2.pdf, Sources/CAD/CAD3.pdf"


================================================================================
 REPORT TYPE 3: REGIONAL OVERVIEW AND MARKET AREA ANALYSIS
================================================================================

FILES YOU NEED TO PROVIDE
--------------------------
1. TWC WDA/MSA Profile PDF (required)
   - Download from lmi.twc.texas.gov (Workforce Development Area Profiles)
   - Must include: labor force stats, employment by industry, wages, firm size
   - Save to: c:\AppraisalWorkspace\Sources\WDA\

OUTPUTS GENERATED
-----------------
- RegionalOverviewAndMarketAnalysis\Output\[Name]Regional Overview and Market Area Analysis.docx

FIELDS LEFT BLANK (insert manually after generation)
-----------------------------------------------------
- Regional Map         -- search "<TODO: Insert Map>" in the Word doc (2 instances)
- Market Area Map      -- same as above

SAMPLE PROMPTS
--------------

  "Generate a Regional Overview and Market Area Analysis for the
   Dallas-Fort Worth WDA using Sources/WDA/DallasFortWorth.pdf.
   The market area county is Tarrant County."

  "Generate a Regional Overview and Market Area Analysis for the
   Houston-Galveston MSA using Sources/WDA/Houston-Galveston.pdf.
   The market area is Harris County."


================================================================================
 NOTES
================================================================================

- If a field cannot be reliably extracted from source files, it is left blank.
  A human should review and fill in any blank fields before finalizing the report.

- For LandComp, the TexasFile PNG is used for the "Data Source" field notation
  and as secondary verification of deed record data. All key data is sourced
  from the MLS PDF.

- Grantor/Grantee names are formatted from MLS source order (Last, First & Second
  --> First and Second Last). Minor reordering differences from examples are
  acceptable -- all names are included.

- For the Regional Overview report, industry projections (2022-2032) are sourced
  from TWC Long-Term Projections. If MSA-level data is unavailable, statewide
  projections adapted to the region's industry mix are used.
