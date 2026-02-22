Convert documents between formats: $ARGUMENTS

## Workflow

1. **Identify source** — confirm input file path and format
2. **Determine target** — clarify desired output format
3. **Check compatibility** — use `convert_formats` to list supported conversions
4. **Convert** — use `convert` (single file) or `convert_batch` (multiple files)
5. **Verify** — check output file exists and is valid
6. **Report** — output path, file size, any warnings

## Supported Format Categories

| Category | Formats |
|---|---|
| Office | docx, xlsx, pptx, odt, ods, odp, rtf |
| PDF | pdf (via pandoc + LaTeX or wkhtmltopdf) |
| Markup | md (Markdown), html, rst (reStructuredText), asciidoc |
| Academic | latex, tex, bibtex |
| Ebook | epub, fb2 |
| Plain | txt, csv, tsv |
| Wiki | mediawiki, dokuwiki, jira |
| Presentation | pptx, revealjs, beamer (LaTeX slides) |
| Data | json, yaml |

Use `mcp__docuflow__convert_formats` to get the full live list.

## Common Conversions

### Office ↔ Open Formats
```
docx → odt, rtf, pdf, md, html, txt, epub, latex
odt → docx, pdf, md, html
xlsx → csv (via excel_save_as or convert)
pptx → pdf, odp
```

### Markdown as Hub Format
```
md → docx, pdf, html, epub, latex, pptx, rst
html → md → any target (two-step via Markdown)
```

### Academic / Publishing
```
latex → pdf, docx, html, epub
md → latex → pdf (for best typesetting)
docx → latex (for journal submission)
```

### Web Content
```
html → docx, pdf, md, epub
md → html (static site content)
rst → html (Sphinx documentation)
```

## Tool Selection

| Scenario | Tool |
|---|---|
| Single file conversion | `mcp__docuflow__convert` |
| Multiple files at once | `mcp__docuflow__convert_batch` |
| List available formats | `mcp__docuflow__convert_formats` |
| Apply template during conversion | `mcp__docuflow__convert_with_template` |
| PDF → editable document | `mcp__docuflow__pdf_to_editable` (better than pandoc for PDFs) |
| Scanned PDF → text | `mcp__docuflow__ocr_pdf` + convert (OCR first) |

## Template-Based Conversion

Use `mcp__docuflow__convert_with_template` when:
- Converting Markdown to styled docx (with custom fonts, headers, margins)
- Generating branded PDF output
- Producing consistent document styles across batch conversions

Template workflow:
1. Create a reference docx with desired styles using `template_create_from_preset`
2. Pass it as the template parameter during conversion

## Batch Conversion

For multiple files:
```
mcp__docuflow__convert_batch with:
  input_files: ["file1.md", "file2.md", "file3.md"]
  output_format: "docx"
  output_dir: "./output/"
```

## Edge Cases & Tips

- **PDF output** requires pandoc + LaTeX (or wkhtmltopdf); if unavailable, create docx first then convert to PDF via the OS
- **Encoding**: source files should be UTF-8; specify encoding if non-UTF-8
- **Images in Markdown**: use absolute paths or ensure images are in the same directory
- **Large files**: convert in chunks if conversion times out
- **Character loss**: if special characters are lost, try an intermediate format (e.g., source → html → target)

## Quality Checklist

Before delivering:
- Output file exists and is non-empty
- Formatting is preserved (headings, bold, tables)
- Images are included (not broken links)
- Character encoding is correct (no garbled text)
- Page layout is reasonable for the target format
