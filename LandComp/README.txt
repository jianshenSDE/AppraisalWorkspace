================================================================================
  LAND COMPARABLE (LandComp) — HOW TO GENERATE A NEW REPORT
================================================================================

WHAT THIS GENERATES
-------------------
  1. An Excel spreadsheet with all comparable sale data (33 fields per property)
  2. A Word document with the data table + property photos from the MLS listing


BEFORE YOU START — GATHER THESE TWO FILES
------------------------------------------

  FILE 1:  MLS PDF  (one or more files, one or more properties per file)
           Export the MLS listing(s) from the MLS system (NTREIS, HAR, etc.)
           Save to:
               c:\AppraisalWorkspace\Sources\MLS\

           How many PDFs do I need?
           - ONE property per PDF  -->  provide one PDF file
           - MULTIPLE properties   -->  you can either:
               a) Export them all into ONE combined PDF  (the script will
                  automatically detect and split each property), OR
               b) Export each property as its own separate PDF file
           Either way works!

  FILE 2:  TexasFile Screenshot  (optional but recommended)
           Screenshot of TexasFile.com search results showing the deed records.
           Save to:
               c:\AppraisalWorkspace\Sources\TexasFile\

           One screenshot can show multiple properties in the results table.
           You can provide one screenshot covering all properties, or one
           per property — it doesn't matter, as the script just notes
           "Texasfile.com" as the data source.


HOW TO GENERATE THE REPORT
---------------------------

  Just copy the sample prompt below, paste it into the AI chat (Windsurf/Cascade),
  swap in your actual file names, and send it. That's it!


----------------------------------------------------------------
  OPTION A — ONE PROPERTY, ONE PDF
----------------------------------------------------------------

  Generate a LandComp report for Corsicana.

  MLS file:       Sources/MLS/MLS-Corsicana, Texas, 75110.pdf
  TexasFile file: Sources/TexasFile/texasfile-navarroSearch.png


----------------------------------------------------------------
  OPTION B — MULTIPLE PROPERTIES IN ONE COMBINED PDF
          (the script automatically detects and separates them)
----------------------------------------------------------------

  Generate a LandComp report for Navarro County.

  MLS file:       Sources/MLS/MLS-NavarroBulkExport.pdf
  TexasFile file: Sources/TexasFile/texasfile-navarroSearch.png

  Note: the MLS PDF contains 3 properties exported together.


----------------------------------------------------------------
  OPTION C — MULTIPLE PROPERTIES, SEPARATE PDF PER PROPERTY
----------------------------------------------------------------

  Generate a LandComp report for Waco with 3 comparable sales.

  MLS files:
    Sources/MLS/MLS-Waco-Property1.pdf
    Sources/MLS/MLS-Waco-Property2.pdf
    Sources/MLS/MLS-Waco-Property3.pdf

  TexasFile file: Sources/TexasFile/texasfile-mclennan.png
  (one screenshot showing all 3 deed records is fine)


----------------------------------------------------------------
  OPTION D — NO TEXASFILE AVAILABLE
----------------------------------------------------------------

  Generate a LandComp report for Austin.

  MLS file: Sources/MLS/MLS-Austin.pdf
  (No TexasFile screenshot available)


WHERE TO FIND YOUR OUTPUT
--------------------------
  After the AI runs, your files will appear here:

      c:\AppraisalWorkspace\LandComp\Output\

      LandComp_[CityName].xlsx   <-- Excel comparison sheet
      LandComp_[CityName].docx   <-- Word doc with table + photos


AFTER GENERATION — FILL IN THESE FIELDS MANUALLY
--------------------------------------------------
  A few fields require a quick manual lookup and cannot be auto-filled:

  +------------------+----------------------------------------------+
  | Field            | Where to get it                              |
  +------------------+----------------------------------------------+
  | Topography       | Site visit or aerial view description        |
  |                  | (e.g. "Level", "Gently sloping", "Rolling")  |
  +------------------+----------------------------------------------+
  | Min Topo (Ft)    | Google Maps -> right-click any point ->      |
  | Max Topo (Ft)    | "What's here?" shows elevation in feet       |
  +------------------+----------------------------------------------+
  | Topo % Change    | Formula: (Max - Min) / Min x 100%            |
  +------------------+----------------------------------------------+

  Everything else is filled in automatically from the MLS PDF.

================================================================================
