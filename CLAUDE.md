# DocuFlow — All-in-One Document Processing

You have access to the **DocuFlow** MCP server with 149 tools for document processing.
Use these tools to help users create, edit, convert, and analyze documents.

## Available Modules

### Word (.docx) — 49 tools
Create and edit Word documents with full formatting support.
- **Document**: `doc_create`, `doc_read`, `doc_info`, `doc_set_properties`, `doc_merge`
- **Content**: `paragraph_add/modify/delete/get`, `heading_add`, `heading_get_outline`
- **Tables**: `table_add`, `table_get`, `table_set_cell`, `table_add_row/column`, `table_delete_row`, `table_merge_cells`, `table_set_column_width`, `table_delete`
- **Images**: `image_add`, `image_add_to_paragraph`
- **Lists**: `list_add_bullet`, `list_add_numbered`
- **Page Setup**: `page_set_margins`, `page_set_size`, `page_add_break`, `page_add_section_break`
- **Headers/Footers**: `header_set`, `footer_set`, `page_number_add`
- **Search**: `search_find`, `search_replace`
- **Special**: `hyperlink_add`, `toc_add`, `line_break_add`, `horizontal_line_add`
- **Styles**: `style_create`, `style_modify`, `style_export`, `style_import`, `doc_get_styles`
- **Templates**: `template_list_presets`, `template_create_from_preset`, `template_apply_styles`
- **Comments**: `comment_add`, `comment_list`
- **Export**: `export_to_text`, `export_to_markdown`

### Excel (.xlsx) — 33 tools
Create spreadsheets with formulas, charts, and data analysis.
- **Workbook**: `excel_create`, `excel_read`, `excel_info`, `excel_save_as`, `excel_status`
- **Sheets**: `sheet_list`, `sheet_add`, `sheet_delete`, `sheet_rename`, `sheet_copy`
- **Cells**: `cell_read`, `cell_write`, `cell_format`, `cell_merge`, `cell_formula`
- **Rows/Cols**: `row_insert`, `row_delete`, `col_insert`, `col_delete`
- **Formulas**: `formula_batch`, `formula_quick`
- **Data**: `data_sort`, `data_filter`, `data_validate`, `data_deduplicate`, `data_fill`
- **Analysis**: `stats_summary`, `conditional_format`, `named_range`, `pivot_create`
- **Charts**: `chart_create`, `excel_chart_modify`
- **Integration**: `excel_to_word`

### PowerPoint (.pptx) — 30 tools
Create presentations with shapes, charts, animations, and transitions.
- **Presentation**: `ppt_create`, `ppt_read`, `ppt_info`, `ppt_set_properties`, `ppt_merge`, `ppt_status`
- **Slides**: `slide_add`, `slide_delete`, `slide_duplicate`, `slide_get_layouts`
- **Shapes**: `shape_add_text`, `shape_add_image`, `shape_add_table`, `shape_add_shape`
- **Placeholders**: `placeholder_list`, `placeholder_set`
- **Backgrounds**: `slide_set_background`
- **Notes**: `slide_add_notes`
- **Animations**: `animation_add`, `animation_list`, `animation_remove`
- **Transitions**: `slide_set_transition`, `slide_remove_transition`
- **Charts**: `chart_add`, `chart_get_data`, `chart_list`, `chart_delete`, `ppt_chart_modify`
- **Masters**: `master_list`, `master_get_info`

### PDF — 23 tools
Extract, manipulate, secure, and convert PDF documents.
- **Info**: `pdf_info`, `pdf_status`, `pdf_get_outline`
- **Extract**: `pdf_extract_text`, `pdf_extract_tables`, `pdf_extract_images`
- **Manipulate**: `pdf_merge`, `pdf_split`, `pdf_extract_pages`, `pdf_rotate`, `pdf_delete_pages`
- **Annotate**: `pdf_add_watermark`, `pdf_text_replace`, `pdf_redact`, `pdf_annotate_text`
- **Security**: `pdf_encrypt`, `pdf_decrypt`
- **Forms**: `pdf_form_get_fields`, `pdf_form_fill`
- **Convert**: `pdf_tables_to_word`, `pdf_tables_to_excel`, `pdf_to_text`, `pdf_to_editable`

### Format Conversion — 4 tools
Convert between 40+ formats via pandoc (docx/pdf/md/html/latex/epub...).
- `convert`, `convert_batch`, `convert_formats`, `convert_with_template`

### OCR — 4 tools
Extract text from images and scanned PDFs.
- `ocr_image`, `ocr_pdf`, `ocr_to_docx`, `ocr_status`

### HTML to PPTX — 3 tools
Convert HTML slides to PowerPoint format.
- `html_to_pptx_convert`, `html_to_pptx_convert_multi`, `html_to_pptx_status`

### AI Image Generation — 3 tools
Generate images from text descriptions.
- `image_gen_status`, `image_generate`, `image_generate_for_ppt`

## Best Practices

1. **Always use absolute paths** for file operations
2. **Check file existence** with `doc_info`/`excel_info`/`ppt_info`/`pdf_info` before editing
3. **Use templates** (`template_create_from_preset`) for professional documents
4. **Batch operations**: use `formula_batch` instead of repeated `cell_formula` calls
5. **Export chain**: create in one format, then `convert` to another if needed

## Common Workflows

### Create a Professional Report
```
1. template_create_from_preset → base document
2. heading_add → sections
3. paragraph_add → content
4. table_add + table_set_cell → data tables
5. image_add → figures
6. toc_add → table of contents
7. convert → export to PDF
```

### Build a Data Dashboard (Excel)
```
1. excel_create → new workbook
2. cell_write → input data
3. formula_batch → calculations
4. chart_create → visualizations
5. conditional_format → highlights
6. pivot_create → summary table
```

### Create a Presentation
```
1. ppt_create → new presentation
2. slide_add → add slides
3. shape_add_text/image/table → content
4. chart_add → data charts
5. animation_add → entrance effects
6. slide_set_transition → slide transitions
```

### Process a PDF
```
1. pdf_info → check structure
2. pdf_extract_text → get content
3. pdf_extract_tables → get tabular data
4. pdf_to_editable → convert to Word/Markdown
5. pdf_encrypt → secure the document
```
