================================================================================
  Retail Market Report — CoStar PDF to Word Converter
================================================================================

PURPOSE
-------
Converts a CoStar Retail Market Report PDF into a formatted Word document (.docx)
that matches the Austin MSA example template exactly (fonts, headings, paragraph
styles, image placement, margins, etc.).


QUICK START
-----------
To convert a single CoStar PDF:

    python convert_costar_to_docx.py "<path_to_costar_pdf>"

The output .docx will be saved in the "output/" folder with an auto-generated
name like "<MSA Name> MSA Retail Market Report.docx".


OPTIONS
-------
    python convert_costar_to_docx.py <input_pdf> [--output <path>] [--template <path>]

    input_pdf       Path to the CoStar Retail Market Report PDF
    --output, -o    Custom output .docx path (optional)
    --template, -t  Custom template .docx path (optional, defaults to the
                    example/Austin MSA Retail Market Report.docx)


EXAMPLES
--------
Single file:
    python convert_costar_to_docx.py "C:\Sources\CoStar\Houston - TX USA-Retail-Market-2026-03-01.pdf"

Custom output path:
    python convert_costar_to_docx.py "C:\Sources\CoStar\Houston - TX USA-Retail-Market-2026-03-01.pdf" -o "C:\Reports\Houston.docx"

Batch (PowerShell — all PDFs in a folder):
    Get-ChildItem "C:\Sources\CoStar\*.pdf" | ForEach-Object { python convert_costar_to_docx.py $_.FullName }


REQUIREMENTS
------------
Python 3.8+ with:
    pip install python-docx PyMuPDF Pillow lxml


HOW IT WORKS
------------
1. Opens the CoStar PDF and identifies sections by page headers:
   Overview, Leasing, Rent, Construction, Under Construction Properties,
   Sales, Sales Past 12 Months, Economy, Submarkets, Supply & Demand Trends,
   Rent & Vacancy, Sale Trends.

2. Classifies each page as "narrative" (flowing two-column paragraph text) or
   "image" (charts, tables, maps) using font analysis — body narrative text is
   Arial ~10pt non-bold; everything else is treated as a chart/table page.

3. For narrative pages: extracts the two-column text, reflows it into single-
   column paragraphs, and renders the metrics summary bar as an image.

4. For image pages: renders the page as a 200 DPI compressed JPEG (quality 85),
   cropping out the CoStar footer (date, copyright notice, page number)
   at the bottom, and embeds it in the Word document at 6.50" width.

5. Creates a BLANK Word document, then copies only the style definitions and
   section properties (margins, page size) from the template via XML deep-copy.
   This is critical — loading the template directly and clearing its paragraphs
   would leave its original image blobs orphaned inside the docx package,
   bloating the file (~40 MB) and causing Word to crash on copy/paste.

6. Applies the template's custom styles:
   - "Style 1 - heading 1" — Title (Calibri 20pt, bold, small caps)
   - "No Spacing"          — Source line (italic)
   - "Style2 - heading 2"  — Section headings (Calibri 12pt, bold)
   - "Normal"              — Body paragraphs (Calibri 10pt, justified)

7. Page layout: 8.5" x 11", 1" margins all around.


FILE STRUCTURE
--------------
RetailMarketReport/
  convert_costar_to_docx.py   — Main conversion script
  README.txt                  — This file
  example/
    Austin MSA Retail Market Report.docx         — Template (styles only; its
                                                   images are NOT carried over)
    Austin - TX USA-Retail-Market-2026-01-11.pdf  — Example source PDF
  output/
    <generated .docx files>


NOTES
-----
- The script auto-detects the MSA name and report date from the PDF.
- CoStar PDFs with more/fewer narrative pages are handled automatically;
  sections with no narrative text will be entirely image-based.
- Output files are typically 6-10 MB (JPEG images, no orphaned media).
- The template .docx is used ONLY as a style source. It is never used as
  the base document for output — this prevents orphaned image blobs.
- If CoStar changes their PDF layout significantly, the font-based page
  classifier (classify_page_by_font) or section header detection may need
  updating.
- The CoStar footer is cropped at y=738 of the 792pt page (constant
  FOOTER_CROP_Y in the script). Adjust if CoStar changes footer position.
