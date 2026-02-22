#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DocuFlow — Complete Installer for OpenAI Codex CLI
===================================================
Installs the Python package, MCP server, agent instructions,
and optional dependencies for Codex.

Usage:
    python install_codex.py              # Interactive install
    python install_codex.py --auto       # Non-interactive, all defaults
    python install_codex.py --uninstall  # Remove DocuFlow from Codex
"""

import json
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

# ── Constants ──────────────────────────────────────────────────────────────

DOCUFLOW_DIR = Path(__file__).resolve().parent
PACKAGE_NAME = "docuflow-mcp"
MCP_SERVER_NAME = "docuflow"
MIN_PYTHON = (3, 10)
CODEX_HOME = Path(os.environ.get("CODEX_HOME", str(Path.home() / ".codex"))).expanduser()
CODEX_SKILLS_DIR = CODEX_HOME / "skills"

CODEX_SKILLS = {
    "ppt-slide-generator": {
        "display_name": "PPT Slide Generator",
        "short_description": "Generate and polish PPT slides from user goals",
        "default_prompt": "Use $ppt-slide-generator to design a professional presentation from my outline.",
        "description": "Generate single-slide or multi-slide PowerPoint decks from user goals by producing slide HTML and converting it to PPTX with DocuFlow tools. Use when users ask to create PPT/PPTX presentations, meeting decks, pitch/report slides, or to turn outlines/content into visual slides.",
        "content": """\
# PPT Slide Generator

## Goal
Create professional presentation slides with a deterministic HTML-to-PPTX pipeline.

## Workflow
1. Clarify scope before generation: topic, audience, slide count, tone, language, and output file path.
2. Draft slide plan first for multi-slide decks (title, section slides, closing).
3. Generate one HTML document per slide.
4. Save HTML under `./output/` using `slide.html` or `slide_1.html`, `slide_2.html`, etc.
5. Convert HTML to PPTX with DocuFlow tools.
6. Report outputs and offer revision passes.

## HTML Contract
Use these hard constraints for conversion compatibility:
- Use only `<div>` and `<p>` content tags.
- Use inline styles only.
- Use `position: absolute` for all placed elements.
- Use a root canvas of `1920x1080` for 16:9 by default.
- Use `1080x1080` only when the user explicitly requests square slides.
- Avoid animation, hover, transitions, and dynamic behavior.

## Layout and Readability Defaults
- Keep margins generous: about 100px horizontal and 80px vertical minimum.
- Keep strong contrast between text and background.
- Keep typography hierarchy clear:
  - Main title: ~72-120px
  - Subtitle: ~36-48px
  - Body text: ~28-36px
  - Caption/footnote: ~20-24px
- Prefer concise bullets and short text blocks.

## Tool Calls
- Single slide conversion: use `mcp__docuflow__html_to_pptx_convert`.
- Multi-slide conversion: prefer `mcp__docuflow__html_to_pptx_convert_multi` when generating all slides together.
- If needed, convert slide-by-slide for easier debugging.

## Quality Checklist
Before conversion, verify:
- Canvas dimensions are correct.
- Only supported tags are used.
- All placement styles are absolute and inline.
- Text remains readable at presentation distance.
- Visual balance is preserved (no overcrowded slide).

## Failure Recovery
If conversion fails:
1. Check malformed HTML structure.
2. Remove unsupported styling or tags.
3. Re-run conversion with the simplified slide.
4. Return a clear error summary and regenerated files.

## Response Pattern
After generation:
- State how many slides were generated.
- Return exact output paths.
- Ask whether to revise style, wording, or layout.
""",
    },
    "report-generator": {
        "display_name": "Report Generator",
        "short_description": "Create professional Word reports with templates",
        "default_prompt": "Use $report-generator to create a professional document.",
        "description": "Create professional Word documents and reports using DocuFlow tools. Handles template selection, document structure, tables, images, TOC, and PDF export.",
        "content": """\
# Report Generator

## Goal
Create professional Word documents with proper structure, formatting, and metadata.

## Workflow
1. Clarify requirements: topic, audience, structure, style, language, output path.
2. Select template with `template_list_presets`, then `template_create_from_preset`.
3. Build structure with `heading_add` for each section.
4. Fill content with `paragraph_add`, `list_add_bullet`/`list_add_numbered`.
5. Insert data with `table_add` + `table_set_cell`, `image_add` for figures.
6. Add navigation with `toc_add`, `page_number_add`.
7. Set metadata with `doc_set_properties`.
8. Export if needed with `convert` to PDF or other formats.

## Document Structure Patterns

### Business Report
1. Title page (heading level 0 + subtitle)
2. Table of Contents
3. Executive Summary
4. Background / Introduction
5. Findings / Analysis (with sub-headings)
6. Data tables and charts
7. Recommendations
8. Appendix

### Technical Document
1. Title + version info
2. Table of Contents
3. Overview
4. Architecture / Design
5. Implementation Details (with sub-headings)
6. Testing / Validation
7. References

## Formatting Conventions
- Use `heading_add` with level 1 for main sections, level 2 for sub-sections.
- Keep paragraphs concise; prefer bullet lists for 3+ related points.
- Use tables for structured data; set column widths with `table_set_column_width`.
- Add `page_add_break` before major sections.
- Use `header_set` / `footer_set` for running headers and page numbers.

## Quality Checklist
- Heading hierarchy is consistent (no skipped levels).
- Tables have clear headers and aligned data.
- TOC is present for documents > 3 pages.
- Metadata (title, author) is set.
""",
    },
    "excel-dashboard": {
        "display_name": "Excel Dashboard",
        "short_description": "Build data dashboards with charts and analysis",
        "default_prompt": "Use $excel-dashboard to build a data dashboard.",
        "description": "Build Excel data dashboards with formulas, charts, conditional formatting, pivot tables, and statistical analysis using DocuFlow tools.",
        "content": """\
# Excel Dashboard

## Goal
Build Excel workbooks with data, formulas, charts, and visual analysis.

## Workflow
1. Clarify requirements: data source, metrics, chart types, audience, output path.
2. Create workbook with `excel_create` with named sheets.
3. Input data with `cell_write`; use `data_fill` for series/patterns.
4. Add formulas with `formula_batch` for bulk calculations.
5. Visualize with `chart_create` for bar/line/pie/scatter charts.
6. Format with `conditional_format` for highlights, `cell_format` for styles.
7. Analyze with `stats_summary` for descriptive stats, `pivot_create` for summaries.
8. Finalize with `sheet_rename` for clear tab names.

## Sheet Organization
- Sheet 1: "Raw Data" — source data with headers in row 1
- Sheet 2: "Calculations" — formulas referencing Raw Data
- Sheet 3: "Dashboard" — charts + KPI summary cells
- Sheet 4: "Pivot" — pivot table summaries (optional)

## Chart Types Guide
- Trend over time → `line`
- Category comparison → `bar` or `column`
- Part of whole → `pie` or `doughnut`
- Correlation → `scatter`

## Tool Notes
- Use `formula_batch` for efficiency (not repeated `cell_formula`).
- Use `formula_quick` for single common operations (sum, average, count).
- Use `named_range` for frequently referenced ranges.
- Use `data_validate` to add dropdown lists or constraints.

## Quality Checklist
- All formulas compute correctly.
- Charts have titles, axis labels, and legends.
- Number formats are appropriate.
- Sheet tabs have descriptive names.
""",
    },
    "pdf-toolkit": {
        "display_name": "PDF Toolkit",
        "short_description": "Merge, split, encrypt, extract, and annotate PDFs",
        "default_prompt": "Use $pdf-toolkit to process my PDF documents.",
        "description": "Perform PDF operations: merge, split, extract pages/text/tables/images, encrypt/decrypt, watermark, redact, fill forms, and convert to editable formats.",
        "content": """\
# PDF Toolkit

## Goal
Perform common PDF manipulation tasks efficiently.

## Workflow
1. Always start with `pdf_info` to understand the document.
2. Determine the operation needed.
3. Execute the appropriate tool(s).
4. Verify with `pdf_info` on the output.
5. Report results and output paths.

## Operations

### Extract Content
- `pdf_extract_text` — get all text (specify page range for large PDFs)
- `pdf_extract_tables` — get structured table data
- `pdf_extract_images` — extract embedded images
- `pdf_get_outline` — get bookmarks/TOC structure

### Manipulate Pages
- `pdf_merge` — combine multiple PDFs
- `pdf_split` — split by page ranges
- `pdf_extract_pages` — pull specific pages
- `pdf_rotate` — rotate 90/180/270 degrees
- `pdf_delete_pages` — remove pages

### Annotate & Edit
- `pdf_add_watermark` — text or image watermark
- `pdf_text_replace` — find and replace text
- `pdf_redact` — permanently remove content (irreversible!)
- `pdf_annotate_text` — add notes/comments

### Security
- `pdf_encrypt` — add password protection
- `pdf_decrypt` — remove password (requires current password)

### Forms
- `pdf_form_get_fields` — list fillable fields
- `pdf_form_fill` — fill fields by name-value mapping

### Convert
- `pdf_to_editable` — PDF to Word/Markdown
- `pdf_tables_to_word` / `pdf_tables_to_excel` — extract tables
- `pdf_to_text` — plain text extraction

## Error Recovery
- Encrypted PDF: ask for password, use `pdf_decrypt` first.
- Empty text: PDF may be scanned images, suggest `ocr_pdf`.
- Form fill fails: use `pdf_form_get_fields` to verify field names.
""",
    },
    "doc-convert": {
        "display_name": "Document Converter",
        "short_description": "Convert between 40+ document formats",
        "default_prompt": "Use $doc-convert to convert my document.",
        "description": "Convert documents between 40+ formats (docx/pdf/md/html/latex/epub/odt/rst/pptx/csv...) using DocuFlow and pandoc.",
        "content": """\
# Document Converter

## Goal
Convert documents between formats accurately while preserving structure.

## Workflow
1. Confirm input file path and format.
2. Clarify desired output format.
3. Use `convert_formats` to verify support if uncertain.
4. Convert with `convert` (single) or `convert_batch` (multiple).
5. Verify output file exists and is valid.
6. Report output path and any warnings.

## Common Conversions
- docx → pdf, md, html, txt, epub, latex, odt, rtf
- md → docx, pdf, html, epub, latex, pptx, rst
- html → docx, pdf, md, epub
- latex → pdf, docx, html, epub
- xlsx → csv (via `excel_save_as`)

## Tool Selection
- Single file: `convert`
- Multiple files: `convert_batch`
- List formats: `convert_formats`
- With template styling: `convert_with_template`
- PDF to editable: `pdf_to_editable` (better than pandoc for PDFs)
- Scanned PDF: `ocr_pdf` first, then convert

## Template-Based Conversion
Use `convert_with_template` when converting Markdown to styled docx:
1. Create reference docx with `template_create_from_preset`.
2. Pass as template parameter during conversion.

## Tips
- PDF output requires pandoc + LaTeX; if unavailable, create docx first.
- Source files should be UTF-8.
- For images in Markdown, use absolute paths.
""",
    },
    "ocr-extract": {
        "display_name": "OCR Extractor",
        "short_description": "Extract text from images and scanned PDFs",
        "default_prompt": "Use $ocr-extract to extract text from my image or scanned document.",
        "description": "Extract text from images (png/jpg/tiff) and scanned PDFs using Tesseract OCR. Supports Chinese, English, Japanese, Korean and 100+ languages.",
        "content": """\
# OCR Extractor

## Goal
Extract text from images and scanned documents using OCR.

## Workflow
1. Identify input type: image file(s) or scanned PDF.
2. Detect or ask for language (Chinese, English, Japanese, etc.).
3. Check OCR availability with `ocr_status`.
4. Run the appropriate OCR tool.
5. Post-process and deliver results.

## Tool Selection
- Single image (png/jpg/tiff/bmp): `ocr_image`
- Multi-page scanned PDF: `ocr_pdf`
- Scanned PDF to editable Word: `ocr_to_docx`
- Check installation: `ocr_status`

## Language Codes
- English: `eng`
- Simplified Chinese: `chi_sim`
- Traditional Chinese: `chi_tra`
- Japanese: `jpn`
- Korean: `kor`
- Mixed Chinese+English: `chi_sim+eng`

## Common Workflows

### Image to Text
1. `ocr_status` → verify Tesseract
2. `ocr_image(path, language)` → extracted text

### Scanned PDF to Editable Word
1. `ocr_to_docx(path, language, output_path)` → .docx file
2. `doc_info` → verify output

## Tips
- 300 DPI is optimal for OCR accuracy.
- High contrast (dark text on white) works best.
- Always specify the correct language code.
- Use `chi_sim+eng` for mixed Chinese-English documents.

## Error Recovery
- Tesseract not installed: provide install instructions.
- Empty result: wrong language or low resolution.
- PDF has selectable text: use `pdf_extract_text` instead (no OCR needed).
""",
    },
}

OPTIONAL_TOOLS = {
    "pandoc": {
        "check_cmd": "pandoc --version",
        "description": "40+ format conversion (docx/pdf/md/html/latex/epub...)",
        "install_hint": {
            "Windows": "winget install JohnMacFarlane.Pandoc   or   choco install pandoc",
            "Darwin":  "brew install pandoc",
            "Linux":   "sudo apt install pandoc   or   sudo pacman -S pandoc",
        },
    },
    "tesseract": {
        "check_cmd": "tesseract --version",
        "description": "OCR text recognition (Chinese/English/Japanese/Korean...)",
        "install_hint": {
            "Windows": "winget install UB-Mannheim.TesseractOCR   or   choco install tesseract",
            "Darwin":  "brew install tesseract tesseract-lang",
            "Linux":   "sudo apt install tesseract-ocr tesseract-ocr-chi-sim",
        },
    },
}

TOTAL_STEPS = 7


# ── Helpers ────────────────────────────────────────────────────────────────

class Color:
    if sys.stdout.isatty():
        GREEN  = "\033[92m"
        YELLOW = "\033[93m"
        RED    = "\033[91m"
        CYAN   = "\033[96m"
        BOLD   = "\033[1m"
        RESET  = "\033[0m"
    else:
        GREEN = YELLOW = RED = CYAN = BOLD = RESET = ""


def info(msg: str):
    print(f"  {Color.GREEN}[OK]{Color.RESET}  {msg}")


def warn(msg: str):
    print(f"  {Color.YELLOW}[!!]{Color.RESET}  {msg}")


def fail(msg: str):
    print(f"  {Color.RED}[ERR]{Color.RESET} {msg}")


def header(msg: str):
    print(f"\n{Color.BOLD}{Color.CYAN}{'─' * 60}")
    print(f"  {msg}")
    print(f"{'─' * 60}{Color.RESET}\n")


def run(cmd: str, capture=False, check=True) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd, shell=True, capture_output=capture, text=True, check=check
    )


def write_utf8(path: Path, text: str):
    """Write UTF-8 text with LF newlines and trailing newline."""
    normalized = text.replace("\r\n", "\n").rstrip("\n") + "\n"
    path.write_text(normalized, encoding="utf-8", newline="\n")


def ask_yes_no(prompt: str, default=True) -> bool:
    if "--auto" in sys.argv:
        return default
    hint = "[Y/n]" if default else "[y/N]"
    try:
        answer = input(f"  {prompt} {hint}: ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return default
    if not answer:
        return default
    return answer in ("y", "yes")


def get_tool_names() -> list:
    """Dynamically discover all registered tool names."""
    result = run(
        f'"{sys.executable}" -c "'
        f"import sys; sys.path.insert(0, r'{DOCUFLOW_DIR / 'src'}'); "
        f"from docuflow_mcp.tools import get_all_tools; "
        f"names = sorted(set(t.name for t in get_all_tools())); "
        f"print('\\n'.join(names))"
        f'"',
        capture=True,
    )
    return [n for n in result.stdout.strip().split("\n") if n]


# ── Steps ──────────────────────────────────────────────────────────────────

def check_prerequisites():
    header(f"Step 1/{TOTAL_STEPS}  Check Prerequisites")

    # Python
    v = sys.version_info
    if (v.major, v.minor) >= MIN_PYTHON:
        info(f"Python {v.major}.{v.minor}.{v.micro}")
    else:
        fail(f"Python {v.major}.{v.minor} — requires >= {MIN_PYTHON[0]}.{MIN_PYTHON[1]}")
        return False

    # Node.js
    try:
        result = run("node --version", capture=True, check=False)
        if result.returncode == 0:
            info(f"Node.js {result.stdout.strip()}")
        else:
            raise FileNotFoundError
    except (FileNotFoundError, OSError):
        fail("Node.js not found — required for Codex CLI")
        system = platform.system()
        hints = {
            "Windows": "winget install OpenJS.NodeJS.LTS   or   https://nodejs.org",
            "Darwin":  "brew install node",
            "Linux":   "sudo apt install nodejs npm   or   https://nodejs.org",
        }
        print(f"         Install:  {hints.get(system, hints['Linux'])}")
        return False

    # Codex CLI
    codex_path = shutil.which("codex")
    if codex_path:
        result = run("codex --version", capture=True, check=False)
        info(f"Codex CLI {result.stdout.strip()}  ({codex_path})")
    else:
        warn("Codex CLI not found — installing now...")
        try:
            run("npm install -g @openai/codex")
            codex_path = shutil.which("codex")
            if codex_path:
                result = run("codex --version", capture=True, check=False)
                info(f"Codex CLI {result.stdout.strip()} installed")
            else:
                fail("Installation succeeded but 'codex' not found on PATH")
                warn("Try restarting your terminal, then re-run this script")
                return False
        except subprocess.CalledProcessError:
            fail("Failed to install Codex CLI")
            print("         Try manually:  npm install -g @openai/codex")
            return False

    return True


def install_docuflow_package():
    header(f"Step 2/{TOTAL_STEPS}  Install DocuFlow Package")

    # Check if already installed
    result = run(
        f'"{sys.executable}" -m pip show {PACKAGE_NAME}',
        capture=True, check=False
    )
    if result.returncode == 0:
        for line in result.stdout.split("\n"):
            if line.startswith("Version:"):
                version = line.split(":", 1)[1].strip()
                info(f"{PACKAGE_NAME} {version} already installed")
                if ask_yes_no("Reinstall / update?", default=False):
                    break
                else:
                    return True

    print(f"  Installing from {DOCUFLOW_DIR} ...\n")
    try:
        run(f'"{sys.executable}" -m pip install -e "{DOCUFLOW_DIR}"')
        info("docuflow-mcp installed successfully")
        return True
    except subprocess.CalledProcessError:
        fail("pip install failed")
        return False


def check_optional_tools():
    header(f"Step 3/{TOTAL_STEPS}  Check Optional Tools")
    system = platform.system()
    all_ok = True

    for name, tool in OPTIONAL_TOOLS.items():
        try:
            result = run(tool["check_cmd"], capture=True, check=False)
            if result.returncode == 0:
                version_line = result.stdout.strip().split("\n")[0]
                info(f"{name}: {version_line}")
            else:
                raise FileNotFoundError
        except (FileNotFoundError, OSError):
            all_ok = False
            warn(f"{name} not found — {tool['description']}")
            hint = tool["install_hint"].get(system, tool["install_hint"].get("Linux", ""))
            if hint:
                print(f"         Install:  {hint}")

    if all_ok:
        info("All optional tools available")
    else:
        print()
        warn("Missing tools are optional. Core features work without them.")

    return True


def register_mcp_server():
    header(f"Step 4/{TOTAL_STEPS}  Register MCP Server")

    # Check if already registered
    result = run(f"codex mcp get {MCP_SERVER_NAME}", capture=True, check=False)
    if result.returncode == 0:
        info(f"'{MCP_SERVER_NAME}' is already registered in Codex")
        print(f"\n{result.stdout.strip()}\n")
        if not ask_yes_no("Re-register (remove and add)?", default=False):
            return True
        run(f"codex mcp remove {MCP_SERVER_NAME}", capture=True, check=False)
        info("Removed existing registration")

    # Determine command
    entry_point = shutil.which("docuflow-mcp")
    if entry_point:
        cmd = f'codex mcp add {MCP_SERVER_NAME} -- docuflow-mcp'
        info(f"Using entry-point: {entry_point}")
    else:
        server_py = str(DOCUFLOW_DIR / "src" / "docuflow_mcp" / "server.py")
        cmd = f'codex mcp add {MCP_SERVER_NAME} -- "{sys.executable}" "{server_py}"'
        warn("Entry-point not on PATH, using direct Python invocation")

    try:
        run(cmd)
        info(f"MCP server '{MCP_SERVER_NAME}' registered in Codex")
    except subprocess.CalledProcessError as e:
        fail(f"Registration failed: {e}")
        print(f"\n  Manual:  codex mcp add {MCP_SERVER_NAME} -- docuflow-mcp\n")
        return False

    # Show result
    result = run(f"codex mcp get {MCP_SERVER_NAME}", capture=True, check=False)
    if result.returncode == 0:
        print(f"\n{result.stdout.strip()}\n")

    return True


def install_agent():
    header(f"Step 5/{TOTAL_STEPS}  Install Agent Instructions")

    codex_md = DOCUFLOW_DIR / "CODEX.md"

    if codex_md.exists():
        info(f"CODEX.md exists ({codex_md.stat().st_size} bytes)")
        if not ask_yes_no("Regenerate with latest tool list?", default=False):
            return True

    # Generate CODEX.md dynamically
    try:
        tool_names = get_tool_names()
    except Exception:
        tool_names = []

    content = f"""# DocuFlow — All-in-One Document Processing

You have access to the **DocuFlow** MCP server with {len(tool_names)} tools for document processing.
Use these tools to help users create, edit, convert, and analyze documents.

## Available Modules

### Word (.docx)
doc_create, doc_read, doc_info, doc_set_properties, doc_merge,
paragraph_add/modify/delete/get, heading_add, heading_get_outline,
table_add/get/set_cell/add_row/add_column/delete_row/merge_cells/delete,
image_add, image_add_to_paragraph, list_add_bullet, list_add_numbered,
page_set_margins, page_set_size, page_add_break, page_add_section_break,
header_set, footer_set, page_number_add, search_find, search_replace,
hyperlink_add, toc_add, line_break_add, horizontal_line_add,
style_create/modify/export/import, template_list_presets/create_from_preset/apply_styles,
comment_add, comment_list, export_to_text, export_to_markdown

### Excel (.xlsx)
excel_create, excel_read, excel_info, excel_save_as,
sheet_list/add/delete/rename/copy,
cell_read/write/format/merge/formula, row_insert/delete, col_insert/delete,
formula_batch, formula_quick, data_sort/filter/validate/deduplicate/fill,
stats_summary, conditional_format, named_range, pivot_create,
chart_create, excel_chart_modify, excel_to_word

### PowerPoint (.pptx)
ppt_create, ppt_read, ppt_info, ppt_set_properties, ppt_merge,
slide_add/delete/duplicate/get_layouts,
shape_add_text/image/table/shape, placeholder_list/set,
slide_set_background, slide_add_notes,
animation_add/list/remove, slide_set_transition/remove_transition,
chart_add/get_data/list/delete/modify, master_list/get_info

### PDF
pdf_info, pdf_extract_text/tables/images, pdf_get_outline,
pdf_merge, pdf_split, pdf_extract_pages, pdf_rotate, pdf_delete_pages,
pdf_add_watermark, pdf_text_replace, pdf_redact, pdf_annotate_text,
pdf_encrypt, pdf_decrypt, pdf_form_get_fields, pdf_form_fill,
pdf_tables_to_word/excel, pdf_to_text, pdf_to_editable

### Others
convert, convert_batch, convert_formats, convert_with_template,
ocr_image, ocr_pdf, ocr_to_docx,
html_to_pptx_convert, html_to_pptx_convert_multi,
image_generate, image_generate_for_ppt

## Best Practices

1. Always use absolute paths for file operations
2. Check file existence with doc_info/excel_info/ppt_info/pdf_info before editing
3. Use templates (template_create_from_preset) for professional documents
4. Use formula_batch instead of repeated cell_formula calls
5. Create in one format, then convert to another if needed

## Common Workflows

### Create a Professional Report
```
1. template_create_from_preset -> base document
2. heading_add -> sections
3. paragraph_add -> content
4. table_add + table_set_cell -> data tables
5. image_add -> figures
6. toc_add -> table of contents
7. convert -> export to PDF
```

### Build a Data Dashboard (Excel)
```
1. excel_create -> new workbook
2. cell_write -> input data
3. formula_batch -> calculations
4. chart_create -> visualizations
5. conditional_format -> highlights
6. pivot_create -> summary table
```

### Create a Presentation
```
1. ppt_create -> new presentation
2. slide_add -> add slides
3. shape_add_text/image/table -> content
4. chart_add -> data charts
5. animation_add -> entrance effects
6. slide_set_transition -> slide transitions
```

### Process a PDF
```
1. pdf_info -> check structure
2. pdf_extract_text -> get content
3. pdf_extract_tables -> get tabular data
4. pdf_to_editable -> convert to Word/Markdown
5. pdf_encrypt -> secure the document
```
"""

    with open(codex_md, "w", encoding="utf-8") as f:
        f.write(content)

    info(f"CODEX.md written ({len(tool_names)} tools documented)")

    # Also trust this project in Codex if possible
    codex_config = Path.home() / ".codex" / "config.toml"
    project_key = str(DOCUFLOW_DIR).replace("\\", "\\\\")
    if codex_config.exists():
        try:
            config_text = codex_config.read_text(encoding="utf-8")
            # Check if project is already trusted
            if str(DOCUFLOW_DIR) in config_text and "trusted" in config_text.split(str(DOCUFLOW_DIR).replace("\\", "\\"))[-1][:50]:
                info("Project already trusted in Codex")
            else:
                info("Tip: run 'codex' in this directory and approve trust when prompted")
        except Exception:
            pass

    return True


def install_codex_skill():
    header(f"Step 6/{TOTAL_STEPS}  Install Codex Skills")

    try:
        CODEX_SKILLS_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        fail(f"Failed to create Codex skills directory: {e}")
        return False

    installed = 0
    for skill_name, skill_data in CODEX_SKILLS.items():
        skill_dir = CODEX_SKILLS_DIR / skill_name

        if skill_dir.exists():
            info(f"Skill '{skill_name}' already exists")
            if not ask_yes_no(f"Overwrite '{skill_name}'?", default=False):
                installed += 1
                continue
            try:
                shutil.rmtree(skill_dir)
            except Exception as e:
                warn(f"Failed to remove existing '{skill_name}': {e}")
                continue

        try:
            agents_dir = skill_dir / "agents"
            agents_dir.mkdir(parents=True, exist_ok=True)

            # Build SKILL.md with frontmatter
            skill_md = (
                f"---\n"
                f"name: {skill_name}\n"
                f"description: {skill_data['description']}\n"
                f"---\n\n"
                f"{skill_data['content']}"
            )
            write_utf8(skill_dir / "SKILL.md", skill_md)

            # Build openai.yaml
            openai_yaml = (
                f"interface:\n"
                f"  display_name: \"{skill_data['display_name']}\"\n"
                f"  short_description: \"{skill_data['short_description']}\"\n"
                f"  default_prompt: \"{skill_data['default_prompt']}\"\n"
            )
            write_utf8(agents_dir / "openai.yaml", openai_yaml)

            info(f"Skill '{skill_name}' installed")
            installed += 1
        except Exception as e:
            warn(f"Failed to install '{skill_name}': {e}")

    info(f"{installed}/{len(CODEX_SKILLS)} skills installed to {CODEX_SKILLS_DIR}")
    if installed > 0:
        warn("Restart Codex to pick up new skills.")
    return True


def verify_installation():
    header(f"Step 7/{TOTAL_STEPS}  Verify Installation")

    # Check tool count
    try:
        tool_names = get_tool_names()
        count = len(tool_names)
        info(f"DocuFlow loaded — {count} MCP tools available")
    except Exception as e:
        fail(f"Verification failed: {e}")
        return False

    # Check codex mcp
    result = run("codex mcp list", capture=True, check=False)
    if MCP_SERVER_NAME in result.stdout:
        info(f"'{MCP_SERVER_NAME}' visible in codex mcp list")
    else:
        warn(f"'{MCP_SERVER_NAME}' not found in codex mcp list")

    # Check CODEX.md
    codex_md = DOCUFLOW_DIR / "CODEX.md"
    if codex_md.exists():
        info(f"Agent instructions: CODEX.md")
    else:
        warn("Agent instructions: CODEX.md missing")

    # Check Codex skills
    installed_skills = []
    for skill_name in CODEX_SKILLS:
        skill_md = CODEX_SKILLS_DIR / skill_name / "SKILL.md"
        if skill_md.exists():
            installed_skills.append(skill_name)
    if installed_skills:
        info(f"Skills installed: {', '.join(installed_skills)}")
    else:
        warn("No Codex skills found")

    print(f"""
  {Color.BOLD}Available modules:{Color.RESET}
    Word (docx)        paragraphs, tables, images, styles, comments...
    Excel (xlsx)       cells, formulas, charts, pivot, conditional format...
    PowerPoint (pptx)  slides, shapes, animations, transitions, charts...
    PDF                extract, merge, split, encrypt, forms, OCR...
    Conversion         40+ formats via pandoc
    AI Image Gen       text-to-image generation
""")
    return True


# ── Banner & completion ───────────────────────────────────────────────────

def print_banner():
    if platform.system() == "Windows":
        os.system("")
    print(f"""
{Color.BOLD}{Color.CYAN}
  ╔══════════════════════════════════════════════════════════╗
  ║                                                          ║
  ║      DocuFlow MCP  —  Codex CLI Installer  v1.1          ║
  ║                                                          ║
  ║   All-in-one document processing for OpenAI Codex        ║
  ║   Word  |  Excel  |  PPT  |  PDF  |  OCR                ║
  ║                                                          ║
  ╚══════════════════════════════════════════════════════════╝
{Color.RESET}""")


def print_done():
    print(f"""
{Color.BOLD}{Color.GREEN}  Installation complete!{Color.RESET}

  What was installed:
    [package]      docuflow-mcp (pip, editable)
    [mcp server]   codex mcp — docuflow registered
    [agent]        CODEX.md — project-level instructions
    [skill]        {len(CODEX_SKILLS)} skills installed to {CODEX_SKILLS_DIR}

  Skills (invoke in Codex):
    $ppt-slide-generator   Generate HTML-to-PPTX presentations
    $report-generator      Create professional Word reports
    $excel-dashboard       Build data dashboards with charts
    $pdf-toolkit           Merge, split, encrypt, extract PDFs
    $doc-convert           Convert between 40+ document formats
    $ocr-extract           OCR text from images and scanned PDFs

  Usage:
    codex "Create a Word document with a quarterly report table"
    codex "Convert my report.docx to PDF"
    codex "Add a pie chart to slide 2 of presentation.pptx"
    codex "$ppt-slide-generator Build a 5-slide product pitch deck"
    codex "$pdf-toolkit Merge and encrypt these three PDFs"

  Management:
    codex mcp list              List registered MCP servers
    codex mcp get docuflow      Show DocuFlow config
    codex mcp remove docuflow   Unregister DocuFlow

  Uninstall:
    python install_codex.py --uninstall

  Note:
    Restart Codex to pick up newly installed skills.
""")


def uninstall():
    print_banner()
    header("Uninstall DocuFlow from Codex")

    # Remove MCP registration
    result = run(f"codex mcp get {MCP_SERVER_NAME}", capture=True, check=False)
    if result.returncode == 0:
        if ask_yes_no(f"Remove '{MCP_SERVER_NAME}' from Codex MCP?", default=True):
            run(f"codex mcp remove {MCP_SERVER_NAME}", check=False)
            info("MCP server removed from Codex")
    else:
        info("DocuFlow not registered in Codex (nothing to remove)")

    # Remove CODEX.md
    codex_md = DOCUFLOW_DIR / "CODEX.md"
    if codex_md.exists():
        if ask_yes_no("Remove CODEX.md agent instructions?", default=False):
            codex_md.unlink()
            info("CODEX.md removed")

    # Remove Codex skills
    installed_skills = [
        name for name in CODEX_SKILLS
        if (CODEX_SKILLS_DIR / name).exists()
    ]
    if installed_skills:
        if ask_yes_no(f"Remove {len(installed_skills)} Codex skills?", default=True):
            for skill_name in installed_skills:
                try:
                    shutil.rmtree(CODEX_SKILLS_DIR / skill_name)
                except Exception:
                    pass
            info(f"Removed {len(installed_skills)} skills")

    # Optionally uninstall Python package
    if ask_yes_no("Also uninstall docuflow-mcp Python package?", default=False):
        run(f'"{sys.executable}" -m pip uninstall -y {PACKAGE_NAME}', check=False)
        info("Python package removed")

    print(f"\n  {Color.GREEN}DocuFlow removed from Codex.{Color.RESET}\n")


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    if "--uninstall" in sys.argv:
        uninstall()
        return

    print_banner()

    steps = [
        check_prerequisites,
        install_docuflow_package,
        check_optional_tools,
        register_mcp_server,
        install_agent,
        install_codex_skill,
        verify_installation,
    ]

    for step in steps:
        if not step():
            fail("Installation aborted.")
            sys.exit(1)

    print_done()


if __name__ == "__main__":
    main()
