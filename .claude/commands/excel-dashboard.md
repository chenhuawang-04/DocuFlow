Build an Excel data dashboard or analysis workbook based on: $ARGUMENTS

## Workflow

1. **Clarify requirements** — data source, metrics, chart types, audience, output path
2. **Create workbook** — `excel_create` with named sheets for raw data, calculations, and dashboard
3. **Input data** — `cell_write` for headers and values; use `data_fill` for series/patterns
4. **Add formulas** — `formula_batch` for bulk calculations (SUM, AVERAGE, VLOOKUP, IF, etc.)
5. **Visualize** — `chart_create` for bar/line/pie/scatter charts
6. **Format** — `conditional_format` for highlights, `cell_format` for number/date/currency styles
7. **Analyze** — `stats_summary` for descriptive stats, `pivot_create` for group-by summaries
8. **Finalize** — `sheet_rename` for clear tab names, freeze panes via `cell_format`

## Sheet Organization Pattern

```
Sheet 1: "Raw Data"     — source data with headers in row 1
Sheet 2: "Calculations"  — formulas referencing Raw Data
Sheet 3: "Dashboard"     — charts + KPI summary cells
Sheet 4: "Pivot"         — pivot table summaries (optional)
```

## Data Input Best Practices

- Always write headers first (row 1), then data starting row 2
- Use `cell_write` with explicit types: `type: "number"` for numeric data, `type: "date"` for dates
- For large datasets, write column by column or use range notation
- Use `data_validate` to add dropdown lists or input constraints

## Formula Patterns

Use `mcp__docuflow__formula_batch` for efficiency. Common patterns:

```json
{
  "formulas": [
    {"cell": "E2", "formula": "=SUM(B2:D2)"},
    {"cell": "E3", "formula": "=AVERAGE(B2:B100)"},
    {"cell": "E4", "formula": "=COUNTIF(C2:C100,\">0\")"},
    {"cell": "E5", "formula": "=VLOOKUP(A2,Sheet2!A:B,2,FALSE)"}
  ]
}
```

Use `mcp__docuflow__formula_quick` for single common operations (sum, average, count, max, min).

## Chart Types Guide

| Data Pattern | Recommended Chart | Notes |
|---|---|---|
| Trend over time | `line` | Use for time series |
| Category comparison | `bar` or `column` | Horizontal bar for long labels |
| Part of whole | `pie` or `doughnut` | Max 6-8 slices |
| Correlation | `scatter` | Add trendline if needed |
| Distribution | `bar` (histogram-style) | Sort by frequency |

## Conditional Formatting Patterns

- **Heat map**: color scale from green (high) to red (low)
- **Data bars**: proportional bars inside cells
- **Icon sets**: arrows or traffic lights for status
- **Threshold**: highlight cells above/below target values

## Analysis Workflow

For statistical analysis:
1. `stats_summary` — get count, mean, median, std dev, min, max
2. `pivot_create` — group by category, aggregate by sum/average/count
3. `chart_create` — visualize the summary data
4. `conditional_format` — highlight outliers or targets

## Tool Call Notes

- Always use absolute paths
- Check existing workbook with `excel_info` before modifying
- Use `sheet_list` to verify sheet names before referencing
- Use `cell_read` to verify written data
- Use `named_range` for frequently referenced ranges

## Quality Checklist

Before delivering:
- All formulas compute correctly (no #REF!, #VALUE!, #N/A)
- Charts have titles, axis labels, and legends
- Number formats are appropriate (currency, percentage, dates)
- Sheet tabs have descriptive names
- Headers are bold/colored for readability
- Data is sorted or filtered as appropriate
