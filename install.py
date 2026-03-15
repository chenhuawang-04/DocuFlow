#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DocuFlow — Complete Installer for Claude Code
==============================================
Installs the Python package, MCP server, agent instructions,
skills, tool permissions, and optional dependencies.

Usage:
    python install.py              # Full install (interactive)
    python install.py --auto       # Non-interactive, all defaults
    python install.py --uninstall  # Remove DocuFlow
"""

import json
import os
import platform
import shlex
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

GLOBAL_COMMANDS_DIR = Path.home() / ".claude" / "commands"
PROJECT_COMMANDS_DIR = DOCUFLOW_DIR / ".claude" / "commands"

SKILL_NAMES = [
    "ppt-slide-generator",
    "report-generator",
    "excel-dashboard",
    "pdf-toolkit",
    "doc-convert",
    "ocr-extract",
]


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


def run(cmd, capture=False, check=True) -> subprocess.CompletedProcess:
    if os.name == 'nt':
        # Windows: pass raw string to shell so cmd.exe handles quoting
        return subprocess.run(
            cmd if isinstance(cmd, str) else subprocess.list2cmdline(cmd),
            shell=True, capture_output=capture, text=True, check=check
        )
    if isinstance(cmd, str):
        args = shlex.split(cmd)
    else:
        args = list(cmd)
    return subprocess.run(
        args, shell=False, capture_output=capture, text=True, check=check
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
    python_code = (
        f"import sys; sys.path.insert(0, r'{DOCUFLOW_DIR / 'src'}'); "
        f"from docuflow_mcp.tools import get_all_tools; "
        f"names = sorted(set(t.name for t in get_all_tools())); "
        f"print('\\n'.join(names))"
    )
    result = subprocess.run(
        [sys.executable, "-c", python_code],
        capture_output=True, text=True, check=True,
    )
    return [n for n in result.stdout.strip().split("\n") if n]


# ── Steps ──────────────────────────────────────────────────────────────────

def check_python_version():
    v = sys.version_info
    if (v.major, v.minor) >= MIN_PYTHON:
        info(f"Python {v.major}.{v.minor}.{v.micro}  ({sys.executable})")
        return True
    else:
        fail(f"Python {v.major}.{v.minor} — requires >= {MIN_PYTHON[0]}.{MIN_PYTHON[1]}")
        return False


def install_package():

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

            # Register MCP server
            data.setdefault("mcpServers", {})
            data["mcpServers"]["docuflow"] = mcp_config

            # Merge tool permissions so DocuFlow works outside this project
            try:
                tool_names = get_tool_names()
                allow_list = [f"mcp__docuflow__{n}" for n in tool_names]
                existing = data.get("permissions", {}).get("allow", [])
                merged = sorted(set(existing) | set(allow_list))
                data.setdefault("permissions", {})
                data["permissions"]["allow"] = merged
                info(f"Global permissions: {len(tool_names)} tools auto-allowed")
            except Exception as e:
                warn(f"Failed to set global permissions: {e}")

            with open(global_settings, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
                f.write("\n")
            info(f"Global settings configured")
        except Exception as e:
            warn(f"Failed to write global settings: {e}")

    return True


def install_agent():

    claude_dir = DOCUFLOW_DIR / ".claude"

    # ── 1. CLAUDE.md — project-level agent instructions ──
    claude_md = DOCUFLOW_DIR / "CLAUDE.md"
    if not claude_md.exists():
        # Try restoring from git first
        try:
            run(f"git -C \"{DOCUFLOW_DIR}\" checkout HEAD -- CLAUDE.md",
                capture=True, check=True)
            info(f"CLAUDE.md restored from git ({claude_md.stat().st_size} bytes)")
        except (subprocess.CalledProcessError, FileNotFoundError):
            warn("CLAUDE.md not found — creating from template")
            _write_agent_instructions(claude_md)
            info("CLAUDE.md created")
    else:
        info(f"CLAUDE.md exists ({claude_md.stat().st_size} bytes)")

    # ── 2. Check project-level skills (restore from git if missing) ──
    PROJECT_COMMANDS_DIR.mkdir(parents=True, exist_ok=True)
    missing_skills = [
        name for name in SKILL_NAMES
        if not (PROJECT_COMMANDS_DIR / f"{name}.md").exists()
    ]
    if missing_skills:
        # Try restoring from git (files may have been deleted in working tree)
        try:
            run(
                f"git -C \"{DOCUFLOW_DIR}\" checkout HEAD -- .claude/commands/",
                capture=True, check=True,
            )
            info(f"Restored {len(missing_skills)} skill(s) from git")
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass  # Not a git repo or files not tracked — skip silently

    for skill_name in SKILL_NAMES:
        skill_file = PROJECT_COMMANDS_DIR / f"{skill_name}.md"
        if skill_file.exists():
            info(f"Skill /{skill_name} ({skill_file.stat().st_size} bytes)")
        else:
            warn(f"Skill /{skill_name} not found (skipped — not critical)")

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

        existing = settings.get("permissions", {}).get("allow", [])
        merged = sorted(set(existing) | set(allow_list))
        settings["permissions"] = {"allow": merged}
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


def install_global_skills():

    if not ask_yes_no(
        f"Install skills globally to {GLOBAL_COMMANDS_DIR}?\n"
        "         (makes /ppt-slide-generator etc. available in all projects)",
        default=True,
    ):
        info("Skipped global skill installation")
        return True

    try:
        GLOBAL_COMMANDS_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        fail(f"Failed to create {GLOBAL_COMMANDS_DIR}: {e}")
        return False

    installed = 0
    for skill_name in SKILL_NAMES:
        src = PROJECT_COMMANDS_DIR / f"{skill_name}.md"
        dst = GLOBAL_COMMANDS_DIR / f"{skill_name}.md"

        if not src.exists():
            warn(f"Source missing: {src.name} (skipped)")
            continue

        if dst.exists():
            # Compare content — skip if identical
            if src.read_bytes() == dst.read_bytes():
                info(f"/{skill_name} already up to date")
                installed += 1
                continue
            if not ask_yes_no(f"Overwrite existing /{skill_name}?", default=True):
                info(f"/{skill_name} kept as-is")
                installed += 1
                continue

        try:
            shutil.copy2(src, dst)
            info(f"/{skill_name} installed")
            installed += 1
        except Exception as e:
            warn(f"Failed to copy {skill_name}: {e}")

    info(f"{installed}/{len(SKILL_NAMES)} skills installed to {GLOBAL_COMMANDS_DIR}")
    return True


def verify_installation():

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

    # Check skills
    commands_dir = DOCUFLOW_DIR / ".claude" / "commands"
    if commands_dir.exists():
        skills = sorted(p.stem for p in commands_dir.glob("*.md"))
        if skills:
            info(f"Project skills: {', '.join('/' + s for s in skills)}")
        else:
            warn("No skills found in .claude/commands/")
    else:
        warn("Project skills directory missing")

    # Check global skills
    if GLOBAL_COMMANDS_DIR.exists():
        global_skills = sorted(
            p.stem for p in GLOBAL_COMMANDS_DIR.glob("*.md")
            if p.stem in SKILL_NAMES
        )
        if global_skills:
            info(f"Global skills: {', '.join('/' + s for s in global_skills)}")
        else:
            info("No DocuFlow skills installed globally")
    else:
        info("Global commands directory not present")

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
    [skill]        .claude/commands/ — 6 project-level skills
    [skill]        ~/.claude/commands/ — 6 global skills (if chosen)
    [permissions]  .claude/settings.local.json — all tools auto-allowed

  Skills (type in Claude Code):
    /ppt-slide-generator   Generate HTML-to-PPTX presentations
    /report-generator      Create professional Word reports
    /excel-dashboard       Build data dashboards with charts
    /pdf-toolkit           Merge, split, encrypt, extract PDFs
    /doc-convert           Convert between 40+ document formats
    /ocr-extract           OCR text from images and scanned PDFs

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
            has_mcp = "mcpServers" in data and "docuflow" in data["mcpServers"]
            has_perms = any(
                p.startswith("mcp__docuflow__")
                for p in data.get("permissions", {}).get("allow", [])
            )
            if has_mcp or has_perms:
                if ask_yes_no(f"Also remove from global settings?", default=True):
                    if has_mcp:
                        del data["mcpServers"]["docuflow"]
                    if has_perms:
                        data["permissions"]["allow"] = [
                            p for p in data["permissions"]["allow"]
                            if not p.startswith("mcp__docuflow__")
                        ]
                    with open(global_settings, "w", encoding="utf-8") as f:
                        json.dump(data, f, indent=2, ensure_ascii=False)
                        f.write("\n")
                    info("Removed MCP server and permissions from global settings")
        except Exception:
            pass

    # Clean settings.local.json permissions
    settings_local = DOCUFLOW_DIR / ".claude" / "settings.local.json"
    if settings_local.exists():
        if ask_yes_no("Remove .claude/settings.local.json?", default=False):
            settings_local.unlink()
            info("Removed settings.local.json")

    # Clean CLAUDE.md
    claude_md = DOCUFLOW_DIR / "CLAUDE.md"
    if claude_md.exists():
        if ask_yes_no("Remove CLAUDE.md (project agent instructions)?", default=False):
            claude_md.unlink()
            info("Removed CLAUDE.md")

    # Clean project-level skills
    project_skills = sorted(PROJECT_COMMANDS_DIR.glob("*.md")) if PROJECT_COMMANDS_DIR.exists() else []
    if project_skills:
        if ask_yes_no(
            f"Remove {len(project_skills)} project skills from .claude/commands/?",
            default=False,
        ):
            for skill_file in project_skills:
                try:
                    skill_file.unlink()
                except Exception:
                    pass
            info(f"Removed {len(project_skills)} project skills")

    # Clean global skills
    global_skills_found = [
        GLOBAL_COMMANDS_DIR / f"{name}.md"
        for name in SKILL_NAMES
        if (GLOBAL_COMMANDS_DIR / f"{name}.md").exists()
    ]
    if global_skills_found:
        if ask_yes_no(
            f"Remove {len(global_skills_found)} DocuFlow skills from {GLOBAL_COMMANDS_DIR}?",
            default=True,
        ):
            for skill_file in global_skills_found:
                try:
                    skill_file.unlink()
                except Exception:
                    pass
            info(f"Removed {len(global_skills_found)} global skills")

    print(f"\n  {Color.GREEN}DocuFlow uninstalled.{Color.RESET}\n")


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    if "--uninstall" in sys.argv:
        uninstall()
        return

    print_banner()

    steps = [
        ("Check Python",                   check_python_version),
        ("Install DocuFlow Package",        install_package),
        ("Check Optional Tools",            check_optional_tools),
        ("Configure MCP Server",            configure_mcp_server),
        ("Install Project Skills & Perms",  install_agent),
        ("Install Global Skills",           install_global_skills),
        ("Verify Installation",             verify_installation),
    ]

    total = len(steps)
    for i, (label, step_fn) in enumerate(steps, 1):
        header(f"Step {i}/{total}  {label}")
        if not step_fn():
            fail("Installation aborted.")
            sys.exit(1)

    print_done()


if __name__ == "__main__":
    main()
