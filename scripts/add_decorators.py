"""
自动为document.py中的所有方法添加@register_tool装饰器
"""
import re

# 工具名称映射（从server.py的dispatch_tool函数推断）
TOOL_MAPPINGS = [
    # DocumentOperations
    ("doc_create", "DocumentOperations", "create", ["path"], ["title", "template", "preset_template"]),
    ("doc_read", "DocumentOperations", "read", ["path"], ["include_formatting"]),
    ("doc_info", "DocumentOperations", "get_info", ["path"], []),
    ("doc_set_properties", "DocumentOperations", "set_properties", ["path", "properties"], []),
    ("doc_merge", "DocumentOperations", "merge", ["paths", "output_path"], ["add_page_break"]),
    ("doc_get_styles", "DocumentOperations", "get_styles", ["path"], []),

    # ParagraphOperations
    ("paragraph_add", "ParagraphOperations", "add", ["path", "text"], ["style", "alignment", "font_name", "font_size", "font_color", "bold", "italic", "underline", "line_spacing", "space_before", "space_after", "first_line_indent"]),
    ("paragraph_modify", "ParagraphOperations", "modify", ["path", "index"], ["text", "style", "alignment", "font_name", "font_size", "font_color", "bold", "italic", "underline"]),
    ("paragraph_delete", "ParagraphOperations", "delete", ["path", "index"], []),
    ("paragraph_get", "ParagraphOperations", "get", ["path", "index"], []),

    # HeadingOperations
    ("heading_add", "HeadingOperations", "add", ["path", "text", "level"], ["alignment"]),
    ("heading_get_outline", "HeadingOperations", "get_outline", ["path"], []),

    # TableOperations
    ("table_add", "TableOperations", "add", ["path", "rows", "cols"], ["data", "style", "header_row"]),
    ("table_get", "TableOperations", "get", ["path", "table_index"], []),
    ("table_set_cell", "TableOperations", "set_cell", ["path", "table_index", "row", "col", "text"], ["bold", "alignment", "vertical_alignment", "background_color"]),
    ("table_add_row", "TableOperations", "add_row", ["path", "table_index"], ["data"]),
    ("table_add_column", "TableOperations", "add_column", ["path", "table_index"], ["data"]),
    ("table_delete_row", "TableOperations", "delete_row", ["path", "table_index", "row_index"], []),
    ("table_merge_cells", "TableOperations", "merge_cells", ["path", "table_index", "start_row", "start_col", "end_row", "end_col"], []),
    ("table_set_column_width", "TableOperations", "set_column_width", ["path", "table_index", "col_index", "width"], []),
    ("table_delete", "TableOperations", "delete", ["path", "table_index"], []),

    # ImageOperations
    ("image_add", "ImageOperations", "add", ["path", "image_path"], ["width", "height", "alignment"]),
    ("image_add_to_paragraph", "ImageOperations", "add_to_paragraph", ["path", "paragraph_index", "image_path"], ["width", "height"]),

    # ListOperations
    ("list_add_bullet", "ListOperations", "add_bullet_list", ["path", "items"], ["level"]),
    ("list_add_numbered", "ListOperations", "add_numbered_list", ["path", "items"], ["level"]),

    # PageOperations
    ("page_set_margins", "PageOperations", "set_margins", ["path"], ["top", "bottom", "left", "right", "section_index"]),
    ("page_set_size", "PageOperations", "set_page_size", ["path"], ["width", "height", "orientation", "section_index"]),
    ("page_add_break", "PageOperations", "add_page_break", ["path"], []),
    ("page_add_section_break", "PageOperations", "add_section_break", ["path"], ["break_type"]),

    # HeaderFooterOperations
    ("header_set", "HeaderFooterOperations", "set_header", ["path", "text"], ["alignment", "section_index"]),
    ("footer_set", "HeaderFooterOperations", "set_footer", ["path", "text"], ["alignment", "section_index"]),
    ("page_number_add", "HeaderFooterOperations", "add_page_number", ["path"], ["position", "alignment"]),

    # SearchOperations
    ("search_find", "SearchOperations", "find", ["path", "text"], ["case_sensitive"]),
    ("search_replace", "SearchOperations", "replace", ["path", "old_text", "new_text"], ["case_sensitive", "replace_all"]),

    # SpecialOperations
    ("hyperlink_add", "SpecialOperations", "add_hyperlink", ["path", "text", "url"], ["paragraph_index"]),
    ("toc_add", "SpecialOperations", "add_table_of_contents", ["path"], []),
    ("line_break_add", "SpecialOperations", "add_line_break", ["path", "paragraph_index"], []),
    ("horizontal_line_add", "SpecialOperations", "add_horizontal_line", ["path"], []),

    # ExportOperations
    ("export_to_text", "ExportOperations", "to_text", ["path"], ["output_path"]),
    ("export_to_markdown", "ExportOperations", "to_markdown", ["path"], ["output_path"]),
]


def add_decorators_to_file():
    """为document.py添加装饰器"""

    # 读取文件
    with open("E:/Project/DocuFlow/src/docuflow_mcp/document.py", "r", encoding="utf-8") as f:
        lines = f.readlines()

    # 查找每个类和方法，添加装饰器
    new_lines = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # 检查是否是@staticmethod行
        if line.strip() == "@staticmethod":
            # 查看下一行的方法定义
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                method_match = re.match(r'\s+def\s+(\w+)\s*\(', next_line)

                if method_match:
                    method_name = method_match.group(1)

                    # 查找对应的工具映射
                    tool_info = None
                    for tool_name, class_name, func_name, required, optional in TOOL_MAPPINGS:
                        if func_name == method_name:
                            tool_info = (tool_name, required, optional)
                            break

                    if tool_info:
                        tool_name, required, optional = tool_info
                        indent = len(line) - len(line.lstrip())

                        # 添加@register_tool装饰器
                        decorator_line = ' ' * indent + f'@register_tool("{tool_name}", required_params={required}, optional_params={optional})\n'
                        new_lines.append(decorator_line)

        new_lines.append(line)
        i += 1

    # 写回文件
    with open("E:/Project/DocuFlow/src/docuflow_mcp/document.py", "w", encoding="utf-8") as f:
        f.writelines(new_lines)

    print(f"✅ 已为document.py添加装饰器")
    print(f"   共添加 {len(TOOL_MAPPINGS)} 个装饰器")


if __name__ == "__main__":
    add_decorators_to_file()
