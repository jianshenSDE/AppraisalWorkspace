================================================================================
  Retail Market Report — CoStar PDF to Word Converter
================================================================================

For comprehensive instructions (formatting rules, styling details, troubleshooting),
see: instructions.md  (in this same folder)

PURPOSE
-------
Converts a CoStar Retail Market Report PDF into a formatted Word document (.docx)
that matches the Austin MSA example template exactly (fonts, headings, paragraph
styles, image placement, margins, etc.).

Also handles non-standard CoStar PDFs (e.g., Single Tenant Net Lease reports)
by rendering all content pages as images.


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

5. Creates a new empty Word document from scratch, then reads the template
   .docx (the Austin example) and copies ONLY its style/formatting definitions
   (fonts, heading sizes, spacing, margins, page size) into the new doc.
   The template's actual content and images are never carried over.
   NOTE: Do NOT use Document(template_path) as the base document — that
   would embed the template's 33 images as invisible orphaned blobs (~20 MB),
   bloating the file and causing Word to crash on copy/paste. Always start
   from Document() and copy styles separately via _copy_styles_from_template().

6. The styles copied from the template are:
   - "Heading 1"    — Title (Calibri 20pt Bold, black)
   - "No Spacing"   — Source line (Calibri 11pt Italic)
   - "Heading 2"    — Section headings (Calibri Bold, black)
   - "Normal"       — Body paragraphs

7. CRITICAL FORMATTING — all narrative body paragraphs must be:
   - Font:      Calibri 10pt (explicitly set on every run)
   - Alignment: JUSTIFY (WD_ALIGN_PARAGRAPH.JUSTIFY)
   These are set explicitly in code, not inherited from styles, to ensure
   they are always correct regardless of template style definitions.

8. Page layout: 8.5" x 11", 1" margins all around.

9. Non-standard PDFs (no section headers detected): all content pages are
   rendered as images. Blank/cover pages (< 20 chars) are skipped.


FILE STRUCTURE
--------------
RetailMarketReport/
  convert_costar_to_docx.py   — Main conversion script
  instructions.md             — Comprehensive instructions (for AI + human)
  README.txt                  — This file (quick-start guide)
  example/
    Example - Austin MSA Retail Market Report (1).docx  — Template (styles only)
    Austin - TX USA-Retail-Market-2026-01-11.pdf        — Example source PDF
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
