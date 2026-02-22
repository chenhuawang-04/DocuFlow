"""
DocuFlow MCP - Formatting utilities

Common formatting functions to reduce code duplication
"""

from typing import Optional
from docx.shared import Pt, Inches, Cm, Mm, RGBColor
from docx.oxml.ns import qn


def apply_font_format(run, font_name=None, font_size=None, font_color=None, bold=None, italic=None, underline=None):
    """Apply font formatting to a run"""
    if font_name:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    if font_size:
        from docuflow_mcp.document import parse_size
        run.font.size = parse_size(font_size)
    if font_color:
        from docuflow_mcp.document import parse_color
        color = parse_color(font_color)
        if color:
            run.font.color.rgb = color
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if underline is not None:
        run.underline = underline


def apply_paragraph_format(para, alignment=None, line_spacing=None, space_before=None, space_after=None, first_line_indent=None):
    """Apply paragraph formatting"""
    from docuflow_mcp.document import get_alignment, parse_size

    if alignment:
        para.alignment = get_alignment(alignment)

    pf = para.paragraph_format
    if line_spacing:
        pf.line_spacing = line_spacing
    if space_before:
        pf.space_before = parse_size(space_before)
    if space_after:
        pf.space_after = parse_size(space_after)
    if first_line_indent:
        pf.first_line_indent = parse_size(first_line_indent)


def apply_style_to_paragraph(para, style_dict):
    """Apply a complete style dictionary to a paragraph"""
    if not style_dict:
        return

    # Apply paragraph formatting
    if 'paragraph' in style_dict:
        p = style_dict['paragraph']
        apply_paragraph_format(
            para,
            alignment=p.get('alignment'),
            line_spacing=p.get('line_spacing'),
            space_before=p.get('space_before'),
            space_after=p.get('space_after'),
            first_line_indent=p.get('first_line_indent')
        )

    # Apply font formatting to all runs
    if 'font' in style_dict:
        f = style_dict['font']
        for run in para.runs:
            apply_font_format(
                run,
                font_name=f.get('name'),
                font_size=f.get('size'),
                font_color=f.get('color'),
                bold=f.get('bold'),
                italic=f.get('italic'),
                underline=f.get('underline')
            )
