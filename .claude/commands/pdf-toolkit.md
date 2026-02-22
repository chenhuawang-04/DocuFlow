Perform PDF operations on: $ARGUMENTS

## Workflow

1. **Inspect** — always start with `pdf_info` to understand the document (page count, metadata, encryption status)
2. **Determine operation** — identify which operations the user needs from the list below
3. **Execute** — call the appropriate tools
4. **Verify** — use `pdf_info` on the output to confirm success
5. **Report** — summarize what was done, output file paths, page counts

## Available Operations

### Extract Content
| Need | Tool | Notes |
|---|---|---|
| Get all text | `pdf_extract_text` | Specify page range for large PDFs |
| Get tables | `pdf_extract_tables` | Returns structured table data |
| Get images | `pdf_extract_images` | Extracts embedded images to files |
| Get outline | `pdf_get_outline` | Returns bookmark/TOC structure |

### Manipulate Pages
| Need | Tool | Notes |
|---|---|---|
| Merge multiple PDFs | `pdf_merge` | Provide list of file paths |
| Split into parts | `pdf_split` | Split by page ranges |
| Extract specific pages | `pdf_extract_pages` | e.g., pages 1,3,5-10 |
| Rotate pages | `pdf_rotate` | 90/180/270 degrees |
| Delete pages | `pdf_delete_pages` | Remove unwanted pages |

### Annotate & Edit
| Need | Tool | Notes |
|---|---|---|
| Add watermark | `pdf_add_watermark` | Text or image watermark |
| Find & replace text | `pdf_text_replace` | Simple text substitution |
| Redact sensitive info | `pdf_redact` | Permanently removes content |
| Add text annotation | `pdf_annotate_text` | Add notes/comments |

### Security
| Need | Tool | Notes |
|---|---|---|
| Encrypt (add password) | `pdf_encrypt` | Set user and/or owner password |
| Decrypt (remove password) | `pdf_decrypt` | Requires current password |

### Forms
| Need | Tool | Notes |
|---|---|---|
| Read form fields | `pdf_form_get_fields` | List all fillable fields |
| Fill form | `pdf_form_fill` | Provide field name → value mapping |

### Convert
| Need | Tool | Notes |
|---|---|---|
| PDF → editable Word/Markdown | `pdf_to_editable` | Best for re-editing content |
| PDF tables → Word | `pdf_tables_to_word` | Preserves table structure |
| PDF tables → Excel | `pdf_tables_to_excel` | For data analysis |
| PDF → plain text | `pdf_to_text` | Simple text extraction |

## Common Multi-Step Workflows

### Merge & Secure
```
1. pdf_info on each input file
2. pdf_merge → combined.pdf
3. pdf_encrypt → combined_secured.pdf
4. pdf_info to verify
```

### Extract & Convert
```
1. pdf_info → check page count
2. pdf_extract_text → review content
3. pdf_extract_tables → get structured data
4. pdf_tables_to_excel → spreadsheet output
```

### Redact & Watermark
```
1. pdf_info → understand structure
2. pdf_redact → remove sensitive content
3. pdf_add_watermark → add "CONFIDENTIAL" stamp
4. pdf_encrypt → password protect
```

### Split & Reorganize
```
1. pdf_info → check total pages
2. pdf_extract_pages → pull relevant pages
3. pdf_rotate → fix orientation if needed
4. pdf_merge → reassemble in new order
```

## Tool Call Notes

- Always use absolute paths for all file parameters
- Always run `pdf_info` first to understand the input
- Encrypted PDFs require `password` parameter for most operations
- `pdf_redact` is irreversible — confirm with user before executing
- For large PDFs, process page ranges instead of entire document
- Output paths should differ from input paths to avoid overwriting

## Error Recovery

- If a PDF is encrypted: ask for password, use `pdf_decrypt` first
- If text extraction returns empty: PDF may be scanned images, suggest OCR (`ocr_pdf`)
- If merge fails: check that all input files exist and are valid PDFs
- If form fill fails: use `pdf_form_get_fields` to verify exact field names
