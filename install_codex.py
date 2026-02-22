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
PPT_SKILL_NAME = "ppt-slide-generator"
PPT_SKILL_DIR = CODEX_SKILLS_DIR / PPT_SKILL_NAME

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

PPT_SKILL_MD = """---
name: ppt-slide-generator
description: Generate single-slide or multi-slide PowerPoint decks from user goals by producing slide HTML and converting it to PPTX with DocuFlow tools. Use when users ask to create PPT/PPTX presentations, meeting decks, pitch/report slides, or to turn outlines/content into visual slides.
---

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
"""

PPT_SKILL_OPENAI_YAML = """interface:
  display_name: "PPT Slide Generator"
  short_description: "Generate and polish PPT slides from user goals"
  default_prompt: "Use $ppt-slide-generator to design a professional presentation from my outline."
"""

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
    header(f"Step 6/{TOTAL_STEPS}  Install Codex Skill")

    try:
        CODEX_SKILLS_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        fail(f"Failed to create Codex skills directory: {e}")
        return False

    if PPT_SKILL_DIR.exists():
        info(f"Skill '{PPT_SKILL_NAME}' already exists at {PPT_SKILL_DIR}")
        if not ask_yes_no("Overwrite with bundled version?", default=False):
            return True
        try:
            shutil.rmtree(PPT_SKILL_DIR)
            info("Removed existing skill directory")
        except Exception as e:
            fail(f"Failed to remove existing skill directory: {e}")
            return False

    try:
        agents_dir = PPT_SKILL_DIR / "agents"
        agents_dir.mkdir(parents=True, exist_ok=True)
        write_utf8(PPT_SKILL_DIR / "SKILL.md", PPT_SKILL_MD)
        write_utf8(agents_dir / "openai.yaml", PPT_SKILL_OPENAI_YAML)
    except Exception as e:
        fail(f"Failed to install skill files: {e}")
        return False

    # Minimal self-check
    skill_file = PPT_SKILL_DIR / "SKILL.md"
    if not skill_file.exists():
        fail("Skill installation failed: SKILL.md missing")
        return False

    info(f"Codex skill installed: {PPT_SKILL_DIR}")
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

    # Check Codex skill
    skill_md = PPT_SKILL_DIR / "SKILL.md"
    if skill_md.exists():
        info(f"Skill installed: {PPT_SKILL_NAME}")
    else:
        warn(f"Skill missing: {PPT_SKILL_NAME}")

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
    [skill]        {PPT_SKILL_NAME} — installed to {PPT_SKILL_DIR}

  Usage:
    codex "Create a Word document with a quarterly report table"
    codex "Convert my report.docx to PDF"
    codex "Add a pie chart to slide 2 of presentation.pptx"
    codex "$ppt-slide-generator Build a 5-slide product pitch deck"

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

    # Remove Codex skill
    if PPT_SKILL_DIR.exists():
        if ask_yes_no(f"Remove Codex skill '{PPT_SKILL_NAME}'?", default=True):
            try:
                shutil.rmtree(PPT_SKILL_DIR)
                info(f"Skill '{PPT_SKILL_NAME}' removed")
            except Exception as e:
                warn(f"Failed to remove skill directory: {e}")

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
