================================================================================
  REGIONAL OVERVIEW & MARKET AREA ANALYSIS — HOW TO GENERATE A NEW REPORT
================================================================================

WHAT THIS GENERATES
-------------------
  A Word document containing:
    - Regional Overview  (labor force, industries, employment trends, wages)
    - Market Area Analysis  (county description, roads, land use, life stage)

  The document follows the same format as the example in the Examples folder.


BEFORE YOU START — GATHER THESE FILES
--------------------------------------

  FILE 1:  TWC WDA/MSA Profile PDF  (required)
           Download the Workforce Development Area profile for your region
           from the Texas Workforce Commission website:

               https://lmi.twc.texas.gov  -->  Workforce Area Profiles

           Save it to this folder:
               c:\AppraisalWorkspace\Sources\WDA\

  That's all you need! The AI will research the county-level market area
  details (roads, land use, demographics) from online sources automatically.


HOW TO GENERATE THE REPORT
---------------------------

  Copy the sample prompt below, paste it into the AI chat (Windsurf/Cascade),
  swap in your actual file name and location details, and send it. That's it!


----------------------------------------------------------------
  SAMPLE PROMPT — COPY THIS AND FILL IN YOUR DETAILS
----------------------------------------------------------------

  Generate a Regional Overview and Market Area Analysis report.

  WDA/MSA file: Sources/WDA/CorpusChristi.pdf

  Region name:  Corpus Christi MSA
  Counties in the MSA:  Aransas, Nueces, and San Patricio

  Market area county:  Aransas County
  County seat:  Rockport


----------------------------------------------------------------
  ANOTHER EXAMPLE — DIFFERENT REGION
----------------------------------------------------------------

  Generate a Regional Overview and Market Area Analysis report.

  WDA/MSA file: Sources/WDA/DallasFortWorth.pdf

  Region name:  Dallas-Fort Worth MSA
  Counties in the MSA:  Dallas, Tarrant, Collin, and Denton

  Market area county:  Tarrant County
  County seat:  Fort Worth


WHERE TO FIND YOUR OUTPUT
--------------------------
  After the AI runs, your file will appear here:

      c:\AppraisalWorkspace\RegionalOverviewAndMarketAnalysis\Output\

      [RegionName]Regional Overview and Market Area Analysis.docx


AFTER GENERATION — INSERT THESE ITEMS MANUALLY
------------------------------------------------
  Two map images need to be inserted manually into the Word document:

  +-------------------+-----------------------------------------------+
  | Placeholder       | What to insert                                |
  +-------------------+-----------------------------------------------+
  | <TODO: Insert Map>| Regional Map — WDA/MSA boundary map showing   |
  | (1st occurrence)  | all counties in the workforce area            |
  +-------------------+-----------------------------------------------+
  | <TODO: Insert Map>| Market Area Map — county-level map showing    |
  | (2nd occurrence)  | the specific market area county               |
  +-------------------+-----------------------------------------------+

  To find them: open the Word doc, press Ctrl+F, search for:
      <TODO: Insert Map>

  Everything else is written automatically from the WDA PDF and
  online research.

================================================================================
