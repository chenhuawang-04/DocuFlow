Extract text from images or scanned documents using OCR: $ARGUMENTS

## Architecture

This skill uses **Claude Code's native multimodal vision** to recognize text — no Tesseract or API key required. For PDFs, DocuFlow tools extract page images first, then you read them directly.

## Workflow

1. **Identify input** — image file(s) or scanned PDF
2. **Prepare images** — if PDF, use DocuFlow to extract page images
3. **Read image** — use the `Read` tool to view each image (Claude's multimodal vision)
4. **Extract text** — transcribe all visible text, preserving structure
5. **Post-process** — save to desired format via DocuFlow tools if needed
6. **Report** — extracted text, output file paths

## Strategy by Input Type

### Single Image (png/jpg/tiff/bmp)
```
1. Read the image file directly (Read tool — multimodal)
2. Transcribe all text you see, preserving layout
3. Return text or save to file
```

### Scanned PDF
```
1. mcp__docuflow__pdf_info → get page count
2. mcp__docuflow__pdf_extract_text → try direct text extraction first
3. If text is empty/garbled (scanned), use mcp__docuflow__pdf_extract_images → export pages as images
4. Read each image (Read tool — multimodal)
5. Transcribe text from each page
6. Combine results
```

### PDF with Selectable Text (not scanned)
```
1. mcp__docuflow__pdf_extract_text → returns good text directly
2. No OCR needed — return the extracted text
```

## Text Extraction Guidelines

When reading an image, extract text following these rules:
- Transcribe **all** visible text: titles, body, headers, footers, captions, watermarks
- Preserve paragraph structure and hierarchy
- Render tables in Markdown table format
- Preserve list formatting (numbered/bulleted)
- Maintain reading order (top-to-bottom, left-to-right)
- For mixed Chinese/English, preserve both languages as-is
- Output **only** the recognized text — no commentary or explanations

## Post-Processing Options

After extraction, offer to:
- Save as plain text file (write to .txt)
- Create a formatted Word document (`mcp__docuflow__doc_create` + `mcp__docuflow__paragraph_add`)
- Save as Markdown
- Extract table data into Excel (`mcp__docuflow__excel_create` + `mcp__docuflow__cell_write`)

## Fallback: Tesseract / Claude API

If the user explicitly requests Tesseract or the Claude API engine:
- Check with `mcp__docuflow__ocr_status`
- Use `mcp__docuflow__ocr_image` (single image) or `mcp__docuflow__ocr_pdf` (PDF)
- These require Tesseract installed or ANTHROPIC_API_KEY set

## Error Recovery

- **Empty image**: file may be corrupted — inform user
- **Unreadable text**: image too low resolution or blurry — suggest higher quality scan
- **PDF not scanned**: `pdf_extract_text` returns good text — use that instead of OCR
- **Large PDF (many pages)**: process in batches of 5-10 pages to avoid context overflow
