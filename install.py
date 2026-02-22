#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DocuFlow — Complete Installer for Claude Code
==============================================
Installs the Python package, MCP server, agent instructions,
tool permissions, and optional dependencies.

Usage:
    python install.py              # Full install (interactive)
    python install.py --auto       # Non-interactive, all defaults
    python install.py --uninstall  # Remove DocuFlow
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
MIN_PYTHON = (3, 10)

MCP_SERVER_CONFIG = {"command": "docuflow-mcp"}
MCP_SERVER_CONFIG_FALLBACK = {
    "command": sys.executable,
    "args": [str(DOCUFLOW_DIR / "src" / "docuflow_mcp" / "server.py")],
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

# Total step count
TOTAL_STEPS = 6


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

def check_python_version():
    header(f"Step 1/{TOTAL_STEPS}  Check Python")
    v = sys.version_info
    if (v.major, v.minor) >= MIN_PYTHON:
        info(f"Python {v.major}.{v.minor}.{v.micro}  ({sys.executable})")
        return True
    else:
        fail(f"Python {v.major}.{v.minor} — requires >= {MIN_PYTHON[0]}.{MIN_PYTHON[1]}")
        return False


def install_package():
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
                if not ask_yes_no("Reinstall / update?", default=False):
                    return True
                break

    print(f"  Installing from {DOCUFLOW_DIR} ...\n")
    try:
        run(f'"{sys.executable}" -m pip install -e "{DOCUFLOW_DIR}"')
        info("docuflow-mcp installed successfully")
        return True
    except subprocess.CalledProcessError:
        # Common on Windows: entry-point .exe is locked by running MCP server
        warn("pip install failed — the MCP server may be locking files")
        warn("This is normal if DocuFlow is already running in Claude Code")
        # Check if package is actually usable despite the error
        check = run(
            f'"{sys.executable}" -c "from docuflow_mcp.tools import get_all_tools; print(len(get_all_tools()))"',
            capture=True, check=False
        )
        if check.returncode == 0:
            info(f"Package is functional ({check.stdout.strip()} tools loaded)")
            return True
        fail("Package is not importable. Close Claude Code and retry.")
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


def configure_mcp_server():
    header(f"Step 4/{TOTAL_STEPS}  Configure MCP Server")

    entry_point = shutil.which("docuflow-mcp")
    if entry_point:
        mcp_config = MCP_SERVER_CONFIG.copy()
        info(f"Found entry-point: {entry_point}")
    else:
        mcp_config = MCP_SERVER_CONFIG_FALLBACK.copy()
        warn("Entry-point not on PATH, using direct Python invocation")

    # ── 1. Project .mcp.json ──
    project_mcp = DOCUFLOW_DIR / ".mcp.json"
    try:
        if project_mcp.exists():
            with open(project_mcp, "r", encoding="utf-8") as f:
                settings = json.load(f)
        else:
            settings = {}
        settings.setdefault("mcpServers", {})
        settings["mcpServers"]["docuflow"] = mcp_config
        with open(project_mcp, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
            f.write("\n")
        info(f".mcp.json configured")
    except Exception as e:
        warn(f"Failed to write .mcp.json: {e}")

    # ── 2. Global settings (optional) ──
    global_settings = Path.home() / ".claude" / "settings.json"
    if ask_yes_no(f"Also register globally in {global_settings}?", default=False):
        try:
            global_settings.parent.mkdir(parents=True, exist_ok=True)
            if global_settings.exists():
                with open(global_settings, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = {}
            data.setdefault("mcpServers", {})
            data["mcpServers"]["docuflow"] = mcp_config
            with open(global_settings, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
                f.write("\n")
            info(f"Global settings configured")
        except Exception as e:
            warn(f"Failed to write global settings: {e}")

    return True


def install_agent():
    header(f"Step 5/{TOTAL_STEPS}  Install Agent & Permissions")

    claude_dir = DOCUFLOW_DIR / ".claude"
    agents_dir = claude_dir / "agents"
    agents_dir.mkdir(parents=True, exist_ok=True)

    # ── 1. CLAUDE.md — project-level agent instructions ──
    claude_md = DOCUFLOW_DIR / "CLAUDE.md"
    if claude_md.exists():
        info(f"CLAUDE.md exists ({claude_md.stat().st_size} bytes)")
    else:
        warn("CLAUDE.md not found — creating from template")
        _write_agent_instructions(claude_md)
        info("CLAUDE.md created")

    # ── 2. PPT Agent ──
    ppt_agent = agents_dir / "ppt-slide-generator.md"
    if ppt_agent.exists():
        info(f"PPT slide generator agent exists ({ppt_agent.stat().st_size} bytes)")
    else:
        warn("PPT agent not found (skipped — not critical)")

    # ── 3. settings.local.json — auto-allow all DocuFlow tools ──
    settings_file = claude_dir / "settings.local.json"
    try:
        tool_names = get_tool_names()
        allow_list = [f"mcp__docuflow__{n}" for n in tool_names]
        allow_list.extend(["WebSearch", "WebFetch"])

        if settings_file.exists():
            with open(settings_file, "r", encoding="utf-8") as f:
                settings = json.load(f)
        else:
            settings = {}

        settings["permissions"] = {"allow": allow_list}
        settings["enableAllProjectMcpServers"] = True
        settings["enabledMcpjsonServers"] = ["docuflow"]

        with open(settings_file, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
            f.write("\n")

        info(f"Tool permissions: {len(tool_names)} tools auto-allowed")
        info(f"Settings written to {settings_file}")

    except Exception as e:
        warn(f"Failed to configure permissions: {e}")
        warn("You may need to manually approve tools on first use")

    return True


def verify_installation():
    header(f"Step 6/{TOTAL_STEPS}  Verify Installation")

    try:
        tool_names = get_tool_names()
        count = len(tool_names)
        info(f"DocuFlow loaded — {count} MCP tools available")
    except Exception as e:
        fail(f"Verification failed: {e}")
        return False

    # Check all files
    checks = [
        (DOCUFLOW_DIR / ".mcp.json",                         "MCP server config"),
        (DOCUFLOW_DIR / "CLAUDE.md",                          "Agent instructions"),
        (DOCUFLOW_DIR / ".claude" / "settings.local.json",    "Tool permissions"),
    ]
    for path, label in checks:
        if path.exists():
            info(f"{label}: {path.name}")
        else:
            warn(f"{label}: missing ({path})")

    print(f"""
  {Color.BOLD}Modules:{Color.RESET}
    Word (docx)        paragraphs, tables, images, styles, comments...
    Excel (xlsx)       cells, formulas, charts, pivot, conditional format...
    PowerPoint (pptx)  slides, shapes, animations, transitions, charts...
    PDF                extract, merge, split, encrypt, forms, OCR...
    Conversion         40+ formats via pandoc
    AI Image Gen       text-to-image generation
""")
    return True


# ── Agent instructions template ───────────────────────────────────────────

def _write_agent_instructions(path: Path):
    """Generate CLAUDE.md with current tool inventory."""
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
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


# ── Banner & completion ───────────────────────────────────────────────────

def print_banner():
    if platform.system() == "Windows":
        os.system("")
    print(f"""
{Color.BOLD}{Color.CYAN}
  ╔══════════════════════════════════════════════════════════╗
  ║                                                          ║
  ║       DocuFlow MCP  —  Claude Code Installer v1.1        ║
  ║                                                          ║
  ║   All-in-one document processing for Claude Code         ║
  ║   Word  |  Excel  |  PPT  |  PDF  |  OCR                ║
  ║                                                          ║
  ╚══════════════════════════════════════════════════════════╝
{Color.RESET}""")


def print_done():
    print(f"""
{Color.BOLD}{Color.GREEN}  Installation complete!{Color.RESET}

  What was installed:
    [package]      docuflow-mcp (pip, editable)
    [mcp server]   .mcp.json
    [agent]        CLAUDE.md — project-level instructions
    [agent]        .claude/agents/ppt-slide-generator.md
    [permissions]  .claude/settings.local.json — all tools auto-allowed

  Next steps:
    1. Restart Claude Code (or run: claude)
    2. Ask Claude to create a document, e.g.:
       "Create a new Word document with a sales report table"

  Useful commands:
    docuflow-mcp                    Start the MCP server directly
    pip show docuflow-mcp           Show package info
    python install.py --uninstall   Remove DocuFlow
""")


def uninstall():
    print_banner()
    header("Uninstall DocuFlow")

    if not ask_yes_no("Remove docuflow-mcp package?", default=True):
        return

    run(f'"{sys.executable}" -m pip uninstall -y {PACKAGE_NAME}', check=False)
    info("Package removed")

    # Clean .mcp.json
    project_mcp = DOCUFLOW_DIR / ".mcp.json"
    if project_mcp.exists():
        try:
            with open(project_mcp, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "mcpServers" in data and "docuflow" in data["mcpServers"]:
                del data["mcpServers"]["docuflow"]
                with open(project_mcp, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                    f.write("\n")
                info(f"Removed docuflow from .mcp.json")
        except Exception:
            pass

    # Clean global settings
    global_settings = Path.home() / ".claude" / "settings.json"
    if global_settings.exists():
        try:
            with open(global_settings, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "mcpServers" in data and "docuflow" in data["mcpServers"]:
                if ask_yes_no(f"Also remove from global settings?", default=True):
                    del data["mcpServers"]["docuflow"]
                    with open(global_settings, "w", encoding="utf-8") as f:
                        json.dump(data, f, indent=2, ensure_ascii=False)
                        f.write("\n")
                    info("Removed from global settings")
        except Exception:
            pass

    # Clean settings.local.json permissions
    settings_local = DOCUFLOW_DIR / ".claude" / "settings.local.json"
    if settings_local.exists():
        if ask_yes_no("Remove .claude/settings.local.json?", default=False):
            settings_local.unlink()
            info("Removed settings.local.json")

    print(f"\n  {Color.GREEN}DocuFlow uninstalled.{Color.RESET}\n")


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    if "--uninstall" in sys.argv:
        uninstall()
        return

    print_banner()

    steps = [
        check_python_version,
        install_package,
        check_optional_tools,
        configure_mcp_server,
        install_agent,
        verify_installation,
    ]

    for step in steps:
        if not step():
            fail("Installation aborted.")
            sys.exit(1)

    print_done()


if __name__ == "__main__":
    main()
