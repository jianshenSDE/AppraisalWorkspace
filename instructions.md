# Regional Overview and Market Area Analysis — Creation Instructions

## Purpose
This document provides step-by-step instructions for creating a **Regional Overview and Market Area Analysis** Word document for any Texas region. It is designed to be reusable across different appraisal assignments by substituting the relevant WDA/MSA and county data.

---

## Document Structure (Exact Format)

The output DOCX must follow this exact structure with these style names (inherited from the template):

| Element | Style Name | Font | Size | Weight |
|---|---|---|---|---|
| Major section title | `Style 1 - heading 1` | Calibri | 20pt | Bold |
| Section sub-header | `Style2 - heading 2` | Calibri | 12pt | Bold |
| Sub-heading within body | Run in `Normal` para | Calibri | 14pt | Bold |
| Map label | Run in `Normal` para | Calibri | 12pt | Bold |
| Land use sub-type (Residential, etc.) | Run in `Normal` para | (inherited) | 10pt | Bold |
| Body text | `Normal` | (inherited) | 10pt | Regular |
| Bulleted list items | `List Paragraph` | (inherited) | 10pt | Regular |

**Page setup:** 8.5" × 11", all margins 1.0 inch, line spacing 1.1 for Normal, space after 6pt.

**Bullet character** (Normal inline bullets): `•` (U+2022)  
**List dash** (List Paragraph items): `–` (U+2013)

---

## Two-Part Document Structure

### Part 1 — Regional Overview

**Heading:** `Regional Overview` (Style 1 - heading 1)

Sub-sections (all 14pt Bold Calibri run inside Normal paragraph):

1. **Introduction** — 2–3 sentences describing the WDA/MSA, counties included, primary city, and economic character.
2. **Regional Map** (Style2 - heading 2) — Placeholder: `<TODO: Insert Map>`
3. *Source citation* — Normal paragraph citing the TWC WDA/MSA Profile and publication date.
4. **Labor Force** — Bullet stats (Normal with • prefix) + 2 analytical sentences.
5. **Industry Overview** — List Paragraph items (sector – %) + 2–3 analysis sentences.
6. **Employment Trends and Projection** — List Paragraph items (industry – % growth) + 2 analysis sentences.
7. **Employment by Firm Size and Ownership** — List Paragraph items + 2 analysis sentences.
8. **Unemployment Trends** — 3 narrative sentences covering pandemic peak, recovery, current level.
9. **Conclusion** — 3 sentences summarizing the regional economy, key strengths, and outlook.

---

### Part 2 — Market Area Analysis

**Heading:** `Market Area Analysis` (Style 1 - heading 1)

Two fixed boilerplate definition paragraphs (Normal) — copy verbatim from template.

Sub-sections (all 14pt Bold Calibri run inside Normal paragraph, except noted):

1. **Market area Map** (12pt Bold) — Placeholder: `<TODO: Insert Map>`
2. **General Description** (14pt Bold) — 3–4 sentences: location, county seat, character, regional role.
3. **Access And Major Roadways** (14pt Bold) — 1 intro sentence + Normal bullets (• prefix) for each highway.
4. **Land Use And Supportive Development** (14pt Bold) — 1 intro sentence, then 5 bold sub-headers (10pt Bold in Normal):
   - **Residential** — 2–3 sentences
   - **Commercial** — 2–3 sentences
   - **Industrial** — 2–3 sentences
   - **Agricultural and Open Land** — 2–3 sentences
   - **Recreational and Institutional** — 2–3 sentences
5. **Life Stage and Trends** (14pt Bold) — 2 sentences, then Normal bullet list of stability factors, then 3 future-trend sentences.
6. **Conclusion** (14pt Bold) — 1 sentence + List Paragraph key factors + 2 closing sentences.

---

## Data Sources Required

### Regional Overview Data
| Data Point | Source |
|---|---|
| WDA/MSA name and counties | Texas Workforce Commission WDA Profile |
| Labor force stats (CLF, Employed, Unemployed, UR) | TWC WDA/MSA Profile — most recent month |
| Employment by industry (sector counts and %) | TWC WDA/MSA Profile — most recent month |
| Average weekly wage | TWC WDA/MSA Profile Q3 wages section OR QCEW via BLS |
| Employment by firm size class | TWC WDA/MSA Profile |
| Employment by ownership (private/govt) | TWC WDA/MSA Profile |
| Projected fastest growing industries 2022–2032 | TWC Long-Term Industry Projections (lmi.twc.texas.gov) |
| Historical unemployment rates | TWC WDA/MSA Profile |
| Texas & US comparators | Same TWC profile |

### Market Area Data (County Level)
| Data Point | Source |
|---|---|
| County location, description, county seat | Wikipedia / Texas Almanac |
| Population and demographics | U.S. Census Bureau QuickFacts |
| Adjacent counties | Wikipedia |
| Major highways | TxDOT / Wikipedia |
| Airport | Wikipedia |
| Land use patterns | Local chamber / county assessor / news |
| Commercial development | Local EDC / news / Rockport-Fulton Chamber |
| Recreational/environmental resources | USFWS / Texas Parks & Wildlife |

---

## Step-by-Step Creation Process

### Step 1 — Gather WDA/MSA Profile PDF
- Download the latest TWC WDA/MSA Profile PDF from `lmi.twc.texas.gov` or receive from client.
- Extract all tabular data: labor force, employment by industry, wages, firm size, historical UR.

### Step 2 — Research County-Level Data
- Pull Census QuickFacts for the specific market area county.
- Note population, median age, housing units, owner/renter split.
- Research major highways via TxDOT or Wikipedia.
- Research commercial development, land use, and economic drivers via local EDC, news, and chamber of commerce sites.

### Step 3 — Gather Industry Projections
- Visit `lmi.twc.texas.gov` → Industry Projections → select the relevant MSA.
- Export or note the top 10 fastest growing industries (% change 2022–2032).
- If MSA-level data is unavailable, use statewide projections with a note.

### Step 4 — Run the Document Generator Script
- Ensure `pymupdf` is installed: `pip install pymupdf`
- Open `generate_regional_analysis.py` in `c:\AppraisalWorkspace\`.
- Update the following constants at the top of the script for the new area:
  - `WDA_PDF_PATH` — path to the new area's TWC WDA/MSA Profile PDF
  - `OUTPUT_PATH` — desired output filename
  - All narrative content variables (MSA name, counties, stats, county description, etc.)
- Run: `python generate_regional_analysis.py`
- Output will appear in `c:\AppraisalWorkspace\Output\`.

**What the script does with the WDA PDF:**
All pages of the WDA/MSA Profile PDF are automatically rendered as high-resolution images (150 DPI) and embedded directly into the Word document at 6.362" width, immediately after the source citation line. This mirrors the example document structure where the Deep East Texas WDA PDF pages are embedded as images before the analyst narrative begins. The script uses `PyMuPDF` (fitz) to render each PDF page to a temporary PNG, inserts it via `python-docx`, then deletes the temp files.

### Step 5 — Insert Maps
- Open the generated DOCX.
- Search for `<TODO: Insert Map>` (there are 2 instances).
- Replace each with the appropriate regional map image.
- Regional Map = WDA/MSA boundary map.
- Market Area Map = County map.

### Step 6 — Review and Finalize
- Confirm all numbers match the source WDA/MSA PDF.
- Verify county description accuracy.
- Adjust tone and detail as needed for the specific appraisal assignment.

---

## Formatting Rules (Do Not Deviate)

1. Never change style names — always use `Style 1 - heading 1`, `Style2 - heading 2`, `Normal`, `List Paragraph`.
2. Sub-headings (Introduction, Labor Force, etc.) are **NOT** a separate style — they are Bold 14pt Calibri runs inside a `Normal` paragraph.
3. Bullet items in body text use `•` (U+2022) prefix inside a Normal paragraph.
4. Bulleted list items under conclusions/key factors use `List Paragraph` style with `–` (U+2013) separator.
5. "Market area Map" uses 12pt Bold (not 14pt) — exception to sub-heading rule.
6. Land use sub-headers (Residential, Commercial, etc.) are Bold 10pt inside Normal paragraphs.
7. No tables in the document — all data is presented as narrative or bulleted lists.
8. Page margins: 1.0 inch all sides.
9. WDA/MSA PDF pages are embedded as inline images at **6.362" width** in `Normal` paragraphs, placed immediately after the source citation and before the analyst narrative sections.

---

## Files Reference

| File | Purpose |
|---|---|
| `instructions.md` | This file — reusable creation guide |
| `generate_regional_analysis.py` | Python script to generate DOCX output |
| `Examples/RegionalOverviewAndMarketAnalysis/DeepEastTexasRegional Overview and Market Area Analysis - Completed.docx` | Format template / reference example |
| `WDA/CorpusChristi.pdf` | Source data for Corpus Christi MSA (also embedded as images) |
| `WDA/Deep East Texas WDA.pdf` | Source data used in the example document |
| `Output/CorpusChristiRegional Overview and Market Area Analysis.docx` | Generated Corpus Christi report |

## Dependencies

| Package | Install | Purpose |
|---|---|---|
| `python-docx` | `pip install python-docx` | Read/write DOCX files |
| `pymupdf` | `pip install pymupdf` | Render PDF pages as images for embedding |
