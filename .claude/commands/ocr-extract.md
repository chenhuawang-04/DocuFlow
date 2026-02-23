Extract text from images or scanned documents using OCR: $ARGUMENTS

## Goal

Extract text reliably with a hybrid strategy: **native image reading for images, direct PDF extraction first, then PDF-to-images plus native reading when PDF quality is poor**.

## Unified Workflow

1. **Identify input type** — image file(s) or PDF.
2. **For images, use native multimodal reading first** when the image is attached in chat.
3. **If PDF, try direct extraction first** — run `mcp__docuflow__pdf_extract_text`.
4. **If PDF text is empty/garbled/scanned, fallback to native reading via images**:
   - Extract page images with `mcp__docuflow__pdf_extract_images`
   - Read extracted page images with native multimodal vision
   - Merge page-level text in order
5. **Post-process and deliver** — structure text, save outputs if requested.

## Tool Selection

- Check OCR availability: `mcp__docuflow__ocr_status`
- Extract selectable PDF text: `mcp__docuflow__pdf_extract_text`
- Extract PDF page images: `mcp__docuflow__pdf_extract_images`
- OCR image: `mcp__docuflow__ocr_image`
- Native multimodal image reading: use for attached images/page screenshots

## Language Codes

- English: `eng`
- Simplified Chinese: `chi_sim`
- Traditional Chinese: `chi_tra`
- Japanese: `jpn`
- Korean: `kor`
- Mixed Chinese + English: `chi_sim+eng`

## Common Patterns

### Image to text
1. Native multimodal reading on attached image
2. If deterministic OCR output is explicitly required: `mcp__docuflow__ocr_status`
3. `mcp__docuflow__ocr_image` with language (optional)

### PDF to text
1. `mcp__docuflow__pdf_extract_text`
2. If low quality, run `mcp__docuflow__pdf_extract_images`
3. Read extracted page images with native multimodal vision and merge by page order

### Scanned PDF to editable Word
1. `mcp__docuflow__pdf_extract_images`
2. Native multimodal reading for each page image
3. Assemble extracted text, then create structured output as needed

## Error Recovery

- **Tesseract not installed**: continue with native reading path; OCR tools are optional.
- **Low-quality scans**: recommend higher DPI (around 300), better contrast, and less blur.
- **Large PDFs**: process pages in batches.
- **No attachment for native reading**: use `mcp__docuflow__pdf_extract_images` first, then read generated images.
