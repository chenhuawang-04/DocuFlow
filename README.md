# DocuFlow MCP

For Chinese documentation, see [README.zh-CN.md](README.zh-CN.md).

DocuFlow is an MCP (Model Context Protocol) server for document processing. It exposes **149 tools** across Word, Excel, PowerPoint, PDF, format conversion, OCR, HTML-to-PPTX, and AI image generation workflows.

The project is designed to provide one consistent MCP surface for common document tasks, so Claude Code, Codex, and other MCP-compatible clients can operate on documents directly.

## What It Covers

DocuFlow groups its tools into the following areas:

| Module | Tools | Capabilities |
| --- | ---: | --- |
| Word (.docx) | 49 | Documents, paragraphs, headings, tables, images, lists, pages, headers/footers, comments, styles, templates, export |
| Excel (.xlsx) | 33 | Workbooks, sheets, cells, formulas, statistics, charts, pivot tables, conditional formatting, data processing |
| PowerPoint (.pptx) | 30 | Slides, text boxes, shapes, charts, animations, transitions, masters, placeholders |
| PDF | 23 | Extract, merge, split, rotate, watermark, forms, redact, convert |
| Conversion | 4 | 40+ format conversions via pandoc |
| OCR | 4 | Image and PDF OCR via Tesseract or OpenAI-compatible completion APIs |
| HTML -> PPTX | 3 | Convert slide-like HTML into PowerPoint |
| AI Image Generation | 3 | Text-to-image and PPT illustration generation |

## Installation

### Option 1: Installer Script

```bash
cd DocuFlow
python install.py
```

### Option 2: Editable Install

```bash
pip install -e .
```

After installation, the MCP entry point is available as:

```bash
docuflow-mcp
```

## MCP Client Setup

Add DocuFlow to your MCP client configuration:

```json
{
  "mcpServers": {
    "docuflow": {
      "command": "docuflow-mcp"
    }
  }
}
```

Restart the client after updating the configuration.

Typical prompts include:

- "Create a new Word document named report.docx with the title Monthly Sales Report."
- "Extract tables from invoice.pdf into Excel."
- "Convert report.docx to PDF."
- "Run OCR on scan.png and return only the extracted text."

## Usage Modes

DocuFlow supports two distinct usage modes.

### 1. MCP Client Usage

This is the default and recommended mode. A client such as Claude Code or Codex connects to `docuflow-mcp` and invokes tools over MCP.

### 2. Direct Python API Usage

The internal Python modules can also be imported directly, for example:

```python
from docuflow_mcp.extensions.ocr import OCROperations
```

This mode is intended for development, testing, and custom integrations. It is not the same as MCP client usage.

## OCR Architecture

Remote OCR is unified behind an **OpenAI-compatible `chat/completions` interface**. The OCR path does not depend on the Anthropic or OpenAI Python SDKs.

OCR tools:

- `ocr_image`
- `ocr_pdf`
- `ocr_to_docx`
- `ocr_status`

### OCR Engines

- `tesseract`
  Local OCR. No remote API is required.
- `completion`
  Remote OCR via an OpenAI-compatible `chat/completions` endpoint.
- `claude`
  Compatibility alias. Internally mapped to `completion`.
- `auto`
  Prefer Tesseract first, then fall back to `completion` when needed.

## OCR Configuration

Remote OCR reads configuration from `ocr_config.json` in the project root:

```json
{
  "api_url": "https://your-api.example.com/v1/chat/completions",
  "model": "grok-4.1-thinking",
  "timeout": 120,
  "api_key": "your-api-key"
}
```

Field definitions:

- `api_url`: Full OpenAI-compatible completion endpoint. In most deployments this must end with `/v1/chat/completions`.
- `model`: Default remote model.
- `timeout`: Request timeout in seconds.
- `api_key`: Remote API key.

### Parameter Precedence

For the `completion` OCR path, effective values are resolved in this order:

1. Explicit tool call arguments
2. `ocr_config.json`
3. Built-in defaults

Examples:

- `ocr_image(image_path="scan.png", engine="completion")`
  Uses `api_url`, `model`, `timeout`, and `api_key` from `ocr_config.json`.
- `ocr_image(image_path="scan.png", engine="completion", model="grok-4")`
  Uses the explicit `model="grok-4"` while the other values continue to come from `ocr_config.json`.

### Default OCR Prompt Policy

The default completion OCR prompt is strict by design. It aims to:

- return only visible text from the image
- avoid explanations, summaries, and answers to questions
- avoid Markdown headings, lists, and code fences
- preserve line breaks and reading order as closely as possible
- avoid duplicate output unless the image itself contains repeated lines

If a specific layout policy is required, you can override the default with the `prompt` parameter in `ocr_image`, `ocr_pdf`, or `ocr_to_docx`.

## OCR Dependencies and Prerequisites

### Image OCR

- Local Tesseract OCR requires `tesseract` to be installed.
- Remote completion OCR requires a valid `ocr_config.json`.

### Scanned PDF OCR

When `ocr_pdf` or `ocr_to_docx` processes scanned PDFs, the pipeline depends on:

- `pdf2image`
- `Pillow`
- On Windows, **Poppler** is commonly required so PDF pages can be rendered to images

Recommended install:

```bash
pip install pdf2image Pillow
```

If you are on Windows, install Poppler separately and ensure its binaries are available to the system. Without it, scanned PDF OCR may fail during PDF-to-image conversion.

## OCR Examples

### Natural Language Examples for MCP Clients

- "Run OCR on scan.png and return only the extracted text."
- "Run OCR on the first three pages of scan.pdf and export the result to Word."
- "Show the current OCR model and endpoint configuration."

### Direct Python API Examples

#### Single Image OCR

```python
from docuflow_mcp.extensions.ocr import OCROperations

result = OCROperations.ocr_image(
    image_path="scan.png",
    engine="completion",
)
```

#### PDF OCR

```python
result = OCROperations.ocr_pdf(
    pdf_path="scan.pdf",
    engine="completion",
    pages=[1, 2, 3],
)
```

#### OCR to Word

```python
result = OCROperations.ocr_to_docx(
    source="scan.pdf",
    output_path="scan_ocr.docx",
    engine="completion",
)
```

#### OCR Status

```python
status = OCROperations.get_status()
```

`ocr_status` reports:

- the resolved `ocr_config.json` path
- the effective `api_url`, `model`, and `timeout`
- Tesseract and completion availability
- the `claude -> completion` compatibility alias

## Project Layout

```text
DocuFlow/
|-- src/docuflow_mcp/
|   |-- server.py                # MCP server entry point
|   |-- document.py              # Word document operations
|   |-- core/
|   |   |-- registry.py          # Tool registration and dispatch
|   |   `-- middleware.py        # Middleware
|   |-- extensions/
|   |   |-- excel.py             # Excel operations
|   |   |-- pdf.py               # PDF operations
|   |   |-- ppt.py               # PowerPoint operations
|   |   |-- converter.py         # Format conversion
|   |   |-- ocr.py               # OCR
|   |   |-- image_gen.py         # AI image generation
|   |   |-- html_to_pptx.py      # HTML to PPTX
|   |   |-- styles.py            # Style management
|   |   |-- templates.py         # Template management
|   |   `-- validator.py         # Validation and repair
|   `-- utils/
|       |-- deps.py              # Dependency checks
|       `-- paths.py             # Path validation
|-- tests/                       # Test suite
|-- scripts/                     # Maintenance and batch-fix scripts
|-- install.py                   # Installer
|-- install_codex.py             # Codex installer
|-- pyproject.toml               # Project metadata
|-- LICENSE                      # Apache 2.0 license
|-- README.md                    # English documentation
`-- README.zh-CN.md              # Chinese documentation
```

## Development and Testing

Install development dependencies:

```bash
pip install -e .[dev]
```

Run the full test suite:

```bash
pytest -q
```

Run OCR tests only:

```bash
pytest -q tests/test_ocr.py
```

## Troubleshooting

### 1. `ocr_image` runs but returns extra explanatory content

Check the following first:

- whether the selected model follows OCR instructions tightly enough
- whether you overrode the default `prompt`
- whether the upstream service maps your requested model to a different effective model

### 2. `ocr_pdf` fails with missing dependency errors

Check the following first:

- whether `pdf2image` is installed
- whether Poppler is installed on Windows
- whether `ocr_status` reports `pdf2image` and `PIL` as available

### 3. Remote OCR fails or times out

Check the following first:

- whether `ocr_config.json` uses a full `/v1/chat/completions` endpoint
- whether `api_key` and `model` are valid
- whether the remote service can process multimodal requests reliably
- whether a proxy or upstream service is failing due to timeout or TLS issues

## Security Notes

- `ocr_config.json` is ignored by Git by default.
- Do not commit real API keys into README examples, tests, or commit history.
- If remote OCR is used in CI, use a separate environment and a restricted key.

## License

Apache License 2.0. See `LICENSE` for the full text.