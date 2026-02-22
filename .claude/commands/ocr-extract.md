Extract text from images or scanned documents using OCR: $ARGUMENTS

## Workflow

1. **Identify input** — image file(s) or scanned PDF
2. **Detect language** — ask user or infer from context (Chinese, English, Japanese, etc.)
3. **Check OCR availability** — use `ocr_status` to verify Tesseract is installed
4. **Run OCR** — choose the right tool based on input type
5. **Post-process** — clean up output, convert to desired format if needed
6. **Report** — extracted text, confidence notes, output file paths

## Tool Selection

| Input Type | Tool | Output |
|---|---|---|
| Single image (png/jpg/tiff/bmp) | `mcp__docuflow__ocr_image` | Extracted text string |
| Multi-page scanned PDF | `mcp__docuflow__ocr_pdf` | Text per page |
| Scanned PDF → editable Word | `mcp__docuflow__ocr_to_docx` | .docx file |
| Check Tesseract status | `mcp__docuflow__ocr_status` | Installation info |

## Language Codes

Common language parameters for the `language` field:

| Language | Code | Notes |
|---|---|---|
| English | `eng` | Default |
| Simplified Chinese | `chi_sim` | Most common for Chinese docs |
| Traditional Chinese | `chi_tra` | Taiwan/HK documents |
| Japanese | `jpn` | |
| Korean | `kor` | |
| French | `fra` | |
| German | `deu` | |
| Spanish | `spa` | |
| Mixed Chinese+English | `chi_sim+eng` | Use `+` to combine languages |

## Common Workflows

### Image → Text
```
1. ocr_status → verify Tesseract is available
2. ocr_image(path, language) → extracted text
3. Return text to user or save to file
```

### Scanned PDF → Searchable Text
```
1. ocr_status → check availability
2. pdf_info → get page count
3. ocr_pdf(path, language, pages) → text per page
4. Combine and format the output
```

### Scanned PDF → Editable Word
```
1. ocr_status → check availability
2. ocr_to_docx(path, language, output_path) → .docx file
3. doc_info → verify the output
4. Optionally apply template styles
```

### Scanned PDF → Excel (tables)
```
1. ocr_pdf → extract text
2. Parse table structure from text
3. excel_create + cell_write → structured spreadsheet
```

### Batch Image OCR
```
For each image:
  1. ocr_image(image_path, language) → text
  2. Collect results
Combine into single document or return per-image
```

## OCR Quality Tips

- **Resolution**: 300 DPI is optimal; below 150 DPI degrades accuracy
- **Contrast**: high contrast (dark text on white background) works best
- **Skew**: straighten rotated images before OCR if possible
- **Language**: always specify the correct language code for best accuracy
- **Mixed languages**: use combined codes like `chi_sim+eng`
- **Noise**: scanned documents with noise/artifacts may need preprocessing

## Error Recovery

- **Tesseract not installed**: inform user with install instructions:
  - Windows: `winget install UB-Mannheim.TesseractOCR`
  - macOS: `brew install tesseract tesseract-lang`
  - Linux: `sudo apt install tesseract-ocr tesseract-ocr-chi-sim`
- **Empty result**: image may be too low resolution, wrong language, or not contain text
- **Garbled output**: likely wrong language code — ask user to confirm document language
- **PDF not scanned**: if `pdf_extract_text` returns good text, OCR is unnecessary — use direct extraction instead

## Post-Processing Options

After OCR extraction, offer to:
- Save as plain text file
- Create a formatted Word document (`doc_create` + `paragraph_add`)
- Convert to Markdown for further editing
- Extract tables into Excel
- Search and replace specific patterns (`search_replace` if in docx)
