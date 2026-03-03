# Retail Market Report — Creation Instructions

## Purpose
This document provides comprehensive instructions for converting CoStar Retail Market Report PDFs into formatted Word documents. It is designed as a reference for both human operators and AI assistants (e.g., Claude Opus via Windsurf/Cascade).

---

## Document Styling (Exact Format — Do Not Deviate)

The output DOCX must match the example document's styling precisely:

| Element | Style | Font | Size | Weight | Alignment |
|---|---|---|---|---|---|
| Report title | `Heading 1` | Calibri | 20pt | Bold | Left |
| Source/date line | `No Spacing` | Calibri | 11pt | Italic | Left |
| Section heading | `Heading 2` | Calibri | (inherited) | Bold | Left |
| Plain section heading* | `Normal` | Calibri | (inherited) | Bold | Left |
| Body/narrative paragraph | `Normal` | Calibri | **10pt** | Regular | **Justify** |
| Image paragraph | `Normal` | — | — | — | Center |

*\* "Rent & Vacancy" and "Sale Trends" sections use plain Normal style headings instead of Heading 2.*

### Critical Styling Rules

1. **Font:** All text must use **Calibri**. Set `run.font.name = 'Calibri'` explicitly on every run.
2. **Body text alignment:** All narrative/body paragraphs must use **JUSTIFY** alignment (`WD_ALIGN_PARAGRAPH.JUSTIFY`). This is the most visible formatting requirement.
3. **Body text size:** Narrative paragraphs use **10pt** (127000 EMU). Set `run.font.size = Pt(10)`.
4. **Title size:** 20pt Bold Calibri, black color.
5. **Source line:** 11pt Calibri Italic.
6. **Section headings:** Calibri Bold, black color (`RGBColor(0, 0, 0)`).
7. **Heading colors:** Explicitly set `run.font.color.rgb = RGBColor(0, 0, 0)` on Heading 1 and Heading 2 runs to override any theme colors.
8. **Page margins:** Inherited from template (1.0 inch all sides, 8.5" × 11").
9. **Image width:** 6.50 inches for all embedded images.

---

## How the Script Works

### Two Processing Modes

**Mode 1 — Standard CoStar MSA Reports** (e.g., "Austin - TX USA-Retail-Market-2026-03-03.pdf"):
1. Identifies sections by matching the first line of each page against known section names (Overview, Leasing, Rent, Construction, etc.).
2. Classifies each page as **narrative** (body text in Arial ~10pt non-bold) or **image** (charts/tables/maps) using font analysis.
3. For narrative pages: extracts two-column text, reflows into single-column paragraphs, and renders the metrics summary bar as an image.
4. For image pages: renders the full page as a 200 DPI compressed JPEG (quality 85), cropping out the CoStar footer at y=738.
5. Produces a Word document with section headings, narrative text (Calibri 10pt Justified), and embedded JPEG images.

**Mode 2 — Non-standard CoStar Reports** (e.g., "United States - Single Tenant Net Lease Retail Report"):
- When no standard section headers are detected (0 sections found), the script falls back to rendering **all content pages as images**.
- Blank/cover pages (< 20 characters of text) are automatically skipped.
- A title and source line are still added at the top.

### Section Detection

The script recognizes these section names (in order):
```
Overview, Leasing, Rent, Construction, Under Construction Properties,
Sales, Sales Past 12 Months, Economy, Submarkets,
Supply & Demand Trends, Rent & Vacancy, Sale Trends
```

### Page Classification (Font Analysis)

A page is classified as **narrative** if:
- More than 50% of body characters are Arial ~10pt non-bold
- At least 200 such characters exist
- Header (top 60pt) and footer (bottom 60pt) zones are excluded

Otherwise it is classified as **image** and rendered as a full-page JPEG.

---

## IMPORTANT — Template Handling (Orphaned Image Prevention)

The script creates a **new empty Word document** via `Document()` and copies ONLY the style definitions and section properties (margins, page size) from the template using `_copy_styles_from_template()`.

**Do NOT:**
- Use `Document(template_path)` as the base document
- Use `shutil.copy()` to copy the template then clear paragraphs
- Load the template and modify its contents

**Why:** Loading the template directly embeds all of its images as invisible orphaned blobs inside the output DOCX. This bloats the file (sometimes 50+ MB) and causes Word to crash on copy/paste operations. The blank-document approach ensures only the styles (font definitions, heading formats, margins) are carried over — no image data.

---

## Usage

### Basic Command
```
cd c:\AppraisalWorkspace\RetailMarketReport
python convert_costar_to_docx.py <input_pdf>
```

### With Options
```
python convert_costar_to_docx.py <input_pdf> [--output <path>] [--template <path>]
```

| Argument | Description |
|---|---|
| `input_pdf` | Path to the CoStar PDF (required) |
| `--output`, `-o` | Custom output path (default: auto-generated in `output/`) |
| `--template`, `-t` | Custom template DOCX path (default: example template) |

### Output Filename Auto-Generation
- **Standard MSA PDFs** (filename contains `- TX USA`): Output is `<MSA> MSA Retail Market Report.docx`
  - Example: `Austin - TX USA-Retail-Market-2026-03-03.pdf` → `Austin MSA Retail Market Report.docx`
- **Non-standard PDFs**: Output uses the PDF's stem filename
  - Example: `United States - Single Tenant Net Lease Retail Report (1).pdf` → `United States - Single Tenant Net Lease Retail Report (1).docx`

### Examples
```
# Standard CoStar MSA report
python convert_costar_to_docx.py "c:\AppraisalWorkspace\Sources\CoStar\Austin - TX USA-Retail-Market-2026-03-03.pdf"

# Non-standard CoStar report
python convert_costar_to_docx.py "c:\AppraisalWorkspace\Sources\CoStar\United States - Single Tenant Net Lease Retail Report (1).pdf"

# Custom output path
python convert_costar_to_docx.py "c:\AppraisalWorkspace\Sources\CoStar\Houston - TX USA-Retail-Market-2026-03-01.pdf" -o "c:\AppraisalWorkspace\RetailMarketReport\output\Houston Report.docx"
```

---

## Constants and Tunable Parameters

| Constant | Value | Description |
|---|---|---|
| `JPEG_QUALITY` | 85 | JPEG compression quality (0–100) |
| `RENDER_DPI` | 200 | PDF page rendering resolution |
| `IMAGE_WIDTH_EMU` | 6.50 inches | Embedded image width in Word |
| `FOOTER_CROP_Y` | 738 | Y-coordinate (of 792pt page) where CoStar footer is cropped |
| `TEMPLATE_PATH` | `example/Example - Austin MSA Retail Market Report (1).docx` | Style source template |

If CoStar changes their footer position, adjust `FOOTER_CROP_Y`. If CoStar changes their PDF layout significantly, the font-based page classifier (`classify_page_by_font`) or section header detection may need updating.

---

## Folder Structure

```
c:\AppraisalWorkspace\
├── Sources/
│   └── CoStar/                              # Input CoStar PDFs
│       ├── Austin - TX USA-Retail-Market-2026-03-03.pdf
│       ├── Corpus Christi - TX USA-Retail-Market-2026-03-01.pdf
│       ├── Dallas-Fort Worth - TX USA-Retail-Market-2026-03-01.pdf
│       ├── Houston - TX USA-Retail-Market-2026-03-01.pdf
│       └── United States - Single Tenant Net Lease Retail Report (1).pdf
└── RetailMarketReport/
    ├── instructions.md                      # This file
    ├── README.txt                           # Quick-start guide
    ├── convert_costar_to_docx.py            # Main conversion script
    ├── example/
    │   ├── Example - Austin MSA Retail Market Report (1).docx  # Style template
    │   └── Austin - TX USA-Retail-Market-2026-01-11.pdf        # Example source PDF
    └── output/                              # Generated Word documents
        ├── Austin MSA Retail Market Report.docx
        └── United States - Single Tenant Net Lease Retail Report (1).docx
```

---

## Files Reference

| File | Purpose |
|---|---|
| `instructions.md` | This file — comprehensive creation guide for AI and human reference |
| `README.txt` | Quick-start guide with brief usage notes |
| `convert_costar_to_docx.py` | Python conversion script |
| `example/Example - Austin MSA Retail Market Report (1).docx` | Template (styles only; images NOT carried over) |
| `example/Austin - TX USA-Retail-Market-2026-01-11.pdf` | Example CoStar source PDF |

---

## Dependencies

| Package | Install | Purpose |
|---|---|---|
| `python-docx` | `pip install python-docx` | Create/modify Word DOCX files |
| `PyMuPDF` | `pip install PyMuPDF` | Parse PDFs, extract text, render pages as images |
| `Pillow` | `pip install Pillow` | Image format conversion (PNG → JPEG) |
| `lxml` | `pip install lxml` | XML manipulation for style copying |

All four: `pip install python-docx PyMuPDF Pillow lxml`

---

## Troubleshooting

| Issue | Cause | Fix |
|---|---|---|
| `PackageNotFoundError` on template | Template DOCX file renamed or missing | Update `TEMPLATE_PATH` constant in script |
| `KeyError: no style with name '...'` | Template uses different style names | Check available styles in template with `[s.name for s in doc.styles]` and update references |
| `PermissionError` on save | Output file is open in Word/Excel | Close the file, or the script auto-saves to `(new).docx` |
| Very large output (>20 MB) | Orphaned images from template | Ensure `Document()` blank construction is used, not `Document(template_path)` |
| Empty output (no sections) | Non-standard PDF without section headers | Script auto-falls back to all-image mode |
| Footer visible in images | CoStar changed footer position | Adjust `FOOTER_CROP_Y` constant |
| Missing narrative text | CoStar changed font sizes | Adjust font size thresholds in `classify_page_by_font` and `extract_narrative_paragraphs` |

---

## Formatting Checklist (Post-Generation Review)

- [ ] Title is Calibri 20pt Bold
- [ ] Source line is Calibri 11pt Italic
- [ ] Section headings are Calibri Bold
- [ ] All body paragraphs are Calibri 10pt **Justified**
- [ ] Images are centered and 6.50" wide
- [ ] No CoStar footer visible in images
- [ ] File size is reasonable (6–10 MB typical for MSA reports)
- [ ] No orphaned images (check with: open in Word, Ctrl+A, check file properties)
