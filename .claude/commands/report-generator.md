Create a professional Word document/report based on: $ARGUMENTS

## Workflow

1. **Clarify requirements** — topic, audience, structure, style, language, output path
2. **Select template** — use `template_list_presets` to show available presets, then `template_create_from_preset` to start with a professional base
3. **Build structure** — add headings with `heading_add` for each section
4. **Fill content** — use `paragraph_add` for body text, `list_add_bullet`/`list_add_numbered` for lists
5. **Insert data** — `table_add` + `table_set_cell` for tables, `image_add` for figures
6. **Add navigation** — `toc_add` for table of contents, `page_number_add` for pagination
7. **Set metadata** — `doc_set_properties` for title/author/subject
8. **Export if needed** — `convert` to PDF or other formats

## Template Selection Guide

Use `mcp__docuflow__template_list_presets` first, then apply one:
- **professional** — formal business documents, reports
- **academic** — papers, theses, research documents
- **minimal** — clean, simple documents
- **creative** — marketing materials, newsletters

## Document Structure Patterns

### Business Report
```
1. Title page (heading level 0 + subtitle paragraph)
2. Table of Contents
3. Executive Summary (heading 1)
4. Background / Introduction (heading 1)
5. Findings / Analysis (heading 1, with sub-headings)
6. Data tables and charts
7. Recommendations (heading 1)
8. Appendix (heading 1)
```

### Technical Document
```
1. Title + version info
2. Table of Contents
3. Overview (heading 1)
4. Architecture / Design (heading 1)
5. Implementation Details (heading 1, with sub-headings)
6. Code snippets in tables
7. Testing / Validation (heading 1)
8. References (heading 1)
```

### Meeting Minutes
```
1. Meeting title + date/time/location
2. Attendees (table)
3. Agenda Items (numbered headings)
4. Discussion summaries (paragraphs)
5. Action Items (table: task / owner / deadline)
6. Next meeting date
```

## Formatting Conventions

- Use `heading_add` with level 1 for main sections, level 2 for sub-sections
- Keep paragraph text concise; prefer bullet lists for 3+ related points
- Use tables for structured data; set column widths with `table_set_column_width`
- Add `page_add_break` before major new sections
- Use `header_set` / `footer_set` for running headers and page numbers
- Apply `style_modify` to customize fonts/colors if the template defaults don't match user preference

## Page Setup Defaults

- Margins: 2.54 cm all sides (standard) — adjust with `page_set_margins` if requested
- Size: A4 by default — use `page_set_size` for Letter or custom
- Orientation: portrait unless user requests landscape

## Tool Calls

- Always use absolute paths for `path` parameters
- Check existing document with `doc_info` before modifying
- Use `paragraph_modify` to edit existing content (not delete + re-add)
- Use `search_replace` for bulk text changes
- Use `comment_add` to leave review notes if requested

## Quality Checklist

Before delivering:
- Heading hierarchy is consistent (no skipped levels)
- Tables have clear headers and aligned data
- Page breaks separate major sections
- TOC is present for documents > 3 pages
- Metadata (title, author) is set
- Document opens without errors
