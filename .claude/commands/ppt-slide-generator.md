Generate a professional PowerPoint presentation based on: $ARGUMENTS

## Workflow

1. **Clarify requirements** — topic, audience, slide count, style (dark/light/tech), language
2. **Plan slide structure** — title slide, content slides, closing slide
3. **Generate HTML** per slide following the contract below
4. **Save HTML** to `./output/` as `slide_1.html`, `slide_2.html`, ...
5. **Convert to PPTX** via DocuFlow tools
6. **Report results** — file paths, slide count, offer revisions

## HTML Contract

Canvas: `1920 x 1080` (16:9). Use `1080 x 1080` only when user requests square slides.

Hard constraints:
- Only `<div>` and `<p>` tags
- All styles inline (`style="..."`)
- All elements `position: absolute`
- No animation, hover, transition, dynamic behavior
- No span, h1-h6, ul, li, or other tags
- No external CSS, class selectors, or IDs

Supported CSS: `position`, `left/top/right/bottom`, `width/height`, `background`, `background-color`, `border-radius`, `font-size`, `font-weight`, `font-family`, `color`, `text-align`

Color formats: hex (`#FF5733`), rgb, rgba

## Base HTML Template

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Slide Title</title>
</head>
<body style="margin: 0; padding: 0;">
    <div style="position: relative; width: 1920px; height: 1080px; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);">
        <!-- content elements here -->
    </div>
</body>
</html>
```

## Design Defaults

Typography:
- Main title: 72-120px
- Subtitle: 36-48px
- Body text: 28-36px
- Caption/footnote: 20-24px

Layout:
- Margins: min 100px horizontal, 80px vertical
- Element spacing: 40-60px
- Generous white space, never overcrowd

Color palettes:

**Dark (default):** background `linear-gradient(135deg, #1a1a2e, #16213e)`, accent `#667eea` / `#764ba2`, text `#ffffff` / `rgba(255,255,255,0.7)`

**Light:** background `#f5f5f5`, accent `#2563eb` / `#7c3aed`, text `#1f2937` / `#6b7280`

**Tech:** background `linear-gradient(135deg, #0f0c29, #302b63, #24243e)`, accent `#00d4ff` / `#7b2ff7`, text `#ffffff` / `rgba(255,255,255,0.6)`

## Slide Types

1. **Title** — centered title + subtitle + accent line
2. **Content (split)** — image placeholder left, bullet points right
3. **List** — numbered items with icon indicators
4. **Card grid** — 3-column cards for features/comparison
5. **Stats** — big numbers with supporting text
6. **Quote** — large quotation with attribution
7. **Closing** — thank you / contact info

## Tool Calls

- Single slide: `mcp__docuflow__html_to_pptx_convert`
- Multi-slide batch: `mcp__docuflow__html_to_pptx_convert_multi`

## Quality Checklist

Before converting, verify each slide:
- All elements use `position: absolute`
- All styles are inline
- Only `div` and `p` tags used
- Canvas is exactly 1920x1080
- Text contrast is sufficient for readability
- Layout is balanced with proper margins

## Error Recovery

If conversion fails:
1. Check HTML structure for malformed tags
2. Remove unsupported CSS properties or tags
3. Re-run conversion with simplified slide
4. Report error and regenerated files to user
