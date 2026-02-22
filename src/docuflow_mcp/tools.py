"""
DocuFlow MCP - 工具定义模块

定义所有 MCP 工具的 schema
"""

from mcp.types import Tool


def get_all_tools():
    """返回所有工具定义"""
    tools = [
        # ============================================================
        # 文档级操作
        # ============================================================
        Tool(
            name="doc_create",
            description="创建新的 Word 文档 (.docx)",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档保存路径，必须以 .docx 结尾"
                    },
                    "title": {
                        "type": "string",
                        "description": "可选的文档标题"
                    },
                    "template": {
                        "type": "string",
                        "description": "可选的模板文件路径"
                    },
                    "preset_template": {
                        "type": "string",
                        "description": "可选的预设模板名称（如 'mba_thesis', 'business_report'）"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="doc_read",
            description="读取 Word 文档的全部内容，返回段落和表格信息",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "include_formatting": {
                        "type": "boolean",
                        "description": "是否包含详细的格式信息",
                        "default": False
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="doc_info",
            description="获取文档的基本信息，包括属性、统计数据和页面设置",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="doc_set_properties",
            description="设置文档属性，如标题、作者、主题等",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "properties": {
                        "type": "object",
                        "description": "要设置的属性",
                        "properties": {
                            "title": {"type": "string", "description": "文档标题"},
                            "author": {"type": "string", "description": "作者"},
                            "subject": {"type": "string", "description": "主题"},
                            "keywords": {"type": "string", "description": "关键词"},
                            "comments": {"type": "string", "description": "备注"},
                            "category": {"type": "string", "description": "类别"}
                        }
                    }
                },
                "required": ["path", "properties"]
            }
        ),

        Tool(
            name="doc_merge",
            description="合并多个 Word 文档为一个",
            inputSchema={
                "type": "object",
                "properties": {
                    "paths": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "要合并的文档路径列表"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文档路径"
                    },
                    "add_page_break": {
                        "type": "boolean",
                        "description": "是否在每个文档之间添加分页符",
                        "default": True
                    }
                },
                "required": ["paths", "output_path"]
            }
        ),

        Tool(
            name="doc_get_styles",
            description="获取文档中可用的所有样式列表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    }
                },
                "required": ["path"]
            }
        ),

        # ============================================================
        # 段落操作
        # ============================================================
        Tool(
            name="paragraph_add",
            description="向文档添加新段落，支持丰富的格式设置",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "text": {
                        "type": "string",
                        "description": "段落文本内容"
                    },
                    "style": {
                        "type": "string",
                        "description": "段落样式名称，如 'Normal', 'Body Text'"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right", "justify"],
                        "description": "段落对齐方式"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "字体名称，如 '微软雅黑', 'Arial'"
                    },
                    "font_size": {
                        "type": "string",
                        "description": "字号，如 '12pt', '14', '0.5in'"
                    },
                    "font_color": {
                        "type": "string",
                        "description": "字体颜色，如 '#FF0000', 'red', 'rgb(255,0,0)'"
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "是否加粗"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "是否斜体"
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "是否下划线"
                    },
                    "line_spacing": {
                        "type": "number",
                        "description": "行距倍数，如 1.5, 2.0"
                    },
                    "space_before": {
                        "type": "string",
                        "description": "段前间距，如 '12pt'"
                    },
                    "space_after": {
                        "type": "string",
                        "description": "段后间距，如 '12pt'"
                    },
                    "first_line_indent": {
                        "type": "string",
                        "description": "首行缩进，如 '2em', '0.5in'"
                    }
                },
                "required": ["path", "text"]
            }
        ),

        Tool(
            name="paragraph_modify",
            description="修改指定段落的内容和格式",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "index": {
                        "type": "integer",
                        "description": "段落索引（从 0 开始）"
                    },
                    "text": {
                        "type": "string",
                        "description": "新的段落文本"
                    },
                    "style": {
                        "type": "string",
                        "description": "段落样式"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right", "justify"],
                        "description": "对齐方式"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "字体名称"
                    },
                    "font_size": {
                        "type": "string",
                        "description": "字号"
                    },
                    "font_color": {
                        "type": "string",
                        "description": "字体颜色"
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "是否加粗"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "是否斜体"
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "是否下划线"
                    }
                },
                "required": ["path", "index"]
            }
        ),

        Tool(
            name="paragraph_delete",
            description="删除指定索引的段落",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "index": {
                        "type": "integer",
                        "description": "要删除的段落索引"
                    }
                },
                "required": ["path", "index"]
            }
        ),

        Tool(
            name="paragraph_get",
            description="获取指定段落的详细信息",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "index": {
                        "type": "integer",
                        "description": "段落索引"
                    }
                },
                "required": ["path", "index"]
            }
        ),

        # ============================================================
        # 标题操作
        # ============================================================
        Tool(
            name="heading_add",
            description="向文档添加标题",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "text": {
                        "type": "string",
                        "description": "标题文本"
                    },
                    "level": {
                        "type": "integer",
                        "minimum": 0,
                        "maximum": 9,
                        "description": "标题级别，0 为 Title，1-9 为 Heading 1-9"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right"],
                        "description": "对齐方式"
                    }
                },
                "required": ["path", "text", "level"]
            }
        ),

        Tool(
            name="heading_get_outline",
            description="获取文档的大纲结构（所有标题的层级关系）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    }
                },
                "required": ["path"]
            }
        ),

        # ============================================================
        # 表格操作
        # ============================================================
        Tool(
            name="table_add",
            description="向文档添加表格",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "rows": {
                        "type": "integer",
                        "minimum": 1,
                        "description": "表格行数"
                    },
                    "cols": {
                        "type": "integer",
                        "minimum": 1,
                        "description": "表格列数"
                    },
                    "data": {
                        "type": "array",
                        "items": {
                            "type": "array",
                            "items": {"type": "string"}
                        },
                        "description": "表格数据，二维数组"
                    },
                    "style": {
                        "type": "string",
                        "description": "表格样式，如 'Table Grid', 'Light Shading'"
                    },
                    "header_row": {
                        "type": "boolean",
                        "description": "是否将第一行作为标题行",
                        "default": True
                    }
                },
                "required": ["path", "rows", "cols"]
            }
        ),

        Tool(
            name="table_get",
            description="获取指定表格的内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "表格索引（从 0 开始）"
                    }
                },
                "required": ["path", "table_index"]
            }
        ),

        Tool(
            name="table_set_cell",
            description="设置表格单元格的内容和格式",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "表格索引"
                    },
                    "row": {
                        "type": "integer",
                        "description": "行索引"
                    },
                    "col": {
                        "type": "integer",
                        "description": "列索引"
                    },
                    "text": {
                        "type": "string",
                        "description": "单元格文本"
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "是否加粗"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right"],
                        "description": "水平对齐"
                    },
                    "vertical_alignment": {
                        "type": "string",
                        "enum": ["top", "center", "bottom"],
                        "description": "垂直对齐"
                    },
                    "background_color": {
                        "type": "string",
                        "description": "背景颜色"
                    }
                },
                "required": ["path", "table_index", "row", "col", "text"]
            }
        ),

        Tool(
            name="table_add_row",
            description="向表格添加新行",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "表格索引"
                    },
                    "data": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "行数据"
                    }
                },
                "required": ["path", "table_index"]
            }
        ),

        Tool(
            name="table_add_column",
            description="向表格添加新列",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "表格索引"
                    },
                    "data": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "列数据"
                    }
                },
                "required": ["path", "table_index"]
            }
        ),

        Tool(
            name="table_delete_row",
            description="删除表格中的指定行",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "表格索引"
                    },
                    "row_index": {
                        "type": "integer",
                        "description": "要删除的行索引"
                    }
                },
                "required": ["path", "table_index", "row_index"]
            }
        ),

        Tool(
            name="table_merge_cells",
            description="合并表格中的单元格",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "表格索引"
                    },
                    "start_row": {
                        "type": "integer",
                        "description": "起始行索引"
                    },
                    "start_col": {
                        "type": "integer",
                        "description": "起始列索引"
                    },
                    "end_row": {
                        "type": "integer",
                        "description": "结束行索引"
                    },
                    "end_col": {
                        "type": "integer",
                        "description": "结束列索引"
                    }
                },
                "required": ["path", "table_index", "start_row", "start_col", "end_row", "end_col"]
            }
        ),

        Tool(
            name="table_set_column_width",
            description="设置表格列宽",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "表格索引"
                    },
                    "col_index": {
                        "type": "integer",
                        "description": "列索引"
                    },
                    "width": {
                        "type": "string",
                        "description": "列宽，如 '2in', '5cm'"
                    }
                },
                "required": ["path", "table_index", "col_index", "width"]
            }
        ),

        Tool(
            name="table_delete",
            description="删除整个表格",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "要删除的表格索引"
                    }
                },
                "required": ["path", "table_index"]
            }
        ),

        # ============================================================
        # 图片操作
        # ============================================================
        Tool(
            name="image_add",
            description="向文档插入图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "image_path": {
                        "type": "string",
                        "description": "图片文件路径"
                    },
                    "width": {
                        "type": "string",
                        "description": "图片宽度，如 '3in', '8cm'"
                    },
                    "height": {
                        "type": "string",
                        "description": "图片高度"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right"],
                        "description": "图片对齐方式"
                    }
                },
                "required": ["path", "image_path"]
            }
        ),

        Tool(
            name="image_add_to_paragraph",
            description="在指定段落中插入图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "paragraph_index": {
                        "type": "integer",
                        "description": "段落索引"
                    },
                    "image_path": {
                        "type": "string",
                        "description": "图片文件路径"
                    },
                    "width": {
                        "type": "string",
                        "description": "图片宽度"
                    },
                    "height": {
                        "type": "string",
                        "description": "图片高度"
                    }
                },
                "required": ["path", "paragraph_index", "image_path"]
            }
        ),

        # ============================================================
        # 列表操作
        # ============================================================
        Tool(
            name="list_add_bullet",
            description="添加无序列表（项目符号列表）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "items": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "列表项内容"
                    },
                    "level": {
                        "type": "integer",
                        "minimum": 0,
                        "description": "缩进级别",
                        "default": 0
                    }
                },
                "required": ["path", "items"]
            }
        ),

        Tool(
            name="list_add_numbered",
            description="添加有序列表（编号列表）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "items": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "列表项内容"
                    },
                    "level": {
                        "type": "integer",
                        "minimum": 0,
                        "description": "缩进级别",
                        "default": 0
                    }
                },
                "required": ["path", "items"]
            }
        ),

        # ============================================================
        # 页面设置
        # ============================================================
        Tool(
            name="page_set_margins",
            description="设置页边距",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "top": {
                        "type": "string",
                        "description": "上边距，如 '1in', '2.54cm'"
                    },
                    "bottom": {
                        "type": "string",
                        "description": "下边距"
                    },
                    "left": {
                        "type": "string",
                        "description": "左边距"
                    },
                    "right": {
                        "type": "string",
                        "description": "右边距"
                    },
                    "section_index": {
                        "type": "integer",
                        "description": "节索引，默认为 0",
                        "default": 0
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="page_set_size",
            description="设置页面大小和方向",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "width": {
                        "type": "string",
                        "description": "页面宽度"
                    },
                    "height": {
                        "type": "string",
                        "description": "页面高度"
                    },
                    "orientation": {
                        "type": "string",
                        "enum": ["portrait", "landscape"],
                        "description": "页面方向：portrait（纵向）或 landscape（横向）"
                    },
                    "section_index": {
                        "type": "integer",
                        "description": "节索引",
                        "default": 0
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="page_add_break",
            description="添加分页符",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="page_add_section_break",
            description="添加分节符",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "break_type": {
                        "type": "string",
                        "enum": ["next_page", "continuous", "even_page", "odd_page"],
                        "description": "分节符类型",
                        "default": "next_page"
                    }
                },
                "required": ["path"]
            }
        ),

        # ============================================================
        # 页眉页脚
        # ============================================================
        Tool(
            name="header_set",
            description="设置页眉",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "text": {
                        "type": "string",
                        "description": "页眉文本"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right"],
                        "description": "对齐方式"
                    },
                    "section_index": {
                        "type": "integer",
                        "description": "节索引",
                        "default": 0
                    }
                },
                "required": ["path", "text"]
            }
        ),

        Tool(
            name="footer_set",
            description="设置页脚",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "text": {
                        "type": "string",
                        "description": "页脚文本"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right"],
                        "description": "对齐方式"
                    },
                    "section_index": {
                        "type": "integer",
                        "description": "节索引",
                        "default": 0
                    }
                },
                "required": ["path", "text"]
            }
        ),

        Tool(
            name="page_number_add",
            description="添加页码",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "position": {
                        "type": "string",
                        "enum": ["header", "footer"],
                        "description": "页码位置",
                        "default": "footer"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right"],
                        "description": "页码对齐方式",
                        "default": "center"
                    }
                },
                "required": ["path"]
            }
        ),

        # ============================================================
        # 搜索替换
        # ============================================================
        Tool(
            name="search_find",
            description="在文档中查找文本",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "text": {
                        "type": "string",
                        "description": "要查找的文本"
                    },
                    "case_sensitive": {
                        "type": "boolean",
                        "description": "是否区分大小写",
                        "default": False
                    }
                },
                "required": ["path", "text"]
            }
        ),

        Tool(
            name="search_replace",
            description="在文档中替换文本",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "old_text": {
                        "type": "string",
                        "description": "要替换的文本"
                    },
                    "new_text": {
                        "type": "string",
                        "description": "替换后的文本"
                    },
                    "case_sensitive": {
                        "type": "boolean",
                        "description": "是否区分大小写",
                        "default": False
                    },
                    "replace_all": {
                        "type": "boolean",
                        "description": "是否替换所有匹配项",
                        "default": True
                    }
                },
                "required": ["path", "old_text", "new_text"]
            }
        ),

        # ============================================================
        # 特殊内容
        # ============================================================
        Tool(
            name="hyperlink_add",
            description="添加超链接",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "text": {
                        "type": "string",
                        "description": "链接显示文本"
                    },
                    "url": {
                        "type": "string",
                        "description": "链接 URL"
                    },
                    "paragraph_index": {
                        "type": "integer",
                        "description": "添加到的段落索引（可选，不指定则新建段落）"
                    }
                },
                "required": ["path", "text", "url"]
            }
        ),

        Tool(
            name="toc_add",
            description="添加目录（需要在 Word 中按 F9 更新）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="line_break_add",
            description="在段落中添加换行符",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "paragraph_index": {
                        "type": "integer",
                        "description": "段落索引"
                    }
                },
                "required": ["path", "paragraph_index"]
            }
        ),

        Tool(
            name="horizontal_line_add",
            description="添加水平分隔线",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    }
                },
                "required": ["path"]
            }
        ),

        # ============================================================
        # 模板管理
        # ============================================================
        Tool(
            name="template_list_presets",
            description="列出所有可用的预设模板",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        Tool(
            name="template_create_from_preset",
            description="从预设模板创建文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "preset_name": {
                        "type": "string",
                        "description": "预设模板名称（如 'mba_thesis', 'business_report', 'academic_paper'）"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文档路径"
                    },
                    "title": {
                        "type": "string",
                        "description": "可选的文档标题"
                    }
                },
                "required": ["preset_name", "output_path"]
            }
        ),

        Tool(
            name="template_apply_styles",
            description="将预设模板的样式应用到现有文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "preset_name": {
                        "type": "string",
                        "description": "预设模板名称"
                    }
                },
                "required": ["path", "preset_name"]
            }
        ),

        # ============================================================
        # 样式管理
        # ============================================================
        Tool(
            name="style_create",
            description="创建自定义样式",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "style_name": {
                        "type": "string",
                        "description": "样式名称"
                    },
                    "style_type": {
                        "type": "string",
                        "enum": ["paragraph", "character", "table"],
                        "description": "样式类型",
                        "default": "paragraph"
                    },
                    "base_style": {
                        "type": "string",
                        "description": "基础样式名称"
                    },
                    "font_config": {
                        "type": "object",
                        "description": "字体配置",
                        "properties": {
                            "name": {"type": "string"},
                            "size": {"type": "string"},
                            "bold": {"type": "boolean"},
                            "italic": {"type": "boolean"},
                            "color": {"type": "string"}
                        }
                    },
                    "paragraph_config": {
                        "type": "object",
                        "description": "段落配置",
                        "properties": {
                            "alignment": {"type": "string"},
                            "line_spacing": {"type": "number"},
                            "space_before": {"type": "string"},
                            "space_after": {"type": "string"},
                            "first_line_indent": {"type": "string"}
                        }
                    }
                },
                "required": ["path", "style_name"]
            }
        ),

        Tool(
            name="style_modify",
            description="修改现有样式",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "style_name": {
                        "type": "string",
                        "description": "样式名称"
                    },
                    "font_config": {
                        "type": "object",
                        "description": "字体配置"
                    },
                    "paragraph_config": {
                        "type": "object",
                        "description": "段落配置"
                    }
                },
                "required": ["path", "style_name"]
            }
        ),

        Tool(
            name="style_export",
            description="导出样式为JSON",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出JSON文件路径（可选）"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="style_import",
            description="从JSON导入样式",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "styles_json": {
                        "type": "string",
                        "description": "样式JSON字符串"
                    }
                },
                "required": ["path", "styles_json"]
            }
        ),

        # ============================================================
        # 导出功能
        # ============================================================
        Tool(
            name="export_to_text",
            description="将文档导出为纯文本",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文件路径（可选，不指定则返回文本内容）"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="export_to_markdown",
            description="将文档导出为 Markdown 格式",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文件路径（可选，不指定则返回 Markdown 内容）"
                    }
                },
                "required": ["path"]
            }
        ),

        # ============================================================
        # Word批注操作
        # ============================================================
        Tool(
            name="comment_add",
            description="为Word文档的指定段落添加批注/评论",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    },
                    "paragraph_index": {
                        "type": "integer",
                        "description": "段落索引（从0开始）"
                    },
                    "text": {
                        "type": "string",
                        "description": "批注文本"
                    },
                    "author": {
                        "type": "string",
                        "description": "作者名称（默认'DocuFlow'）"
                    },
                    "date": {
                        "type": "string",
                        "description": "日期（ISO格式，默认当前时间）"
                    }
                },
                "required": ["path", "paragraph_index", "text"]
            }
        ),

        Tool(
            name="comment_list",
            description="列出Word文档中的所有批注/评论",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文档路径"
                    }
                },
                "required": ["path"]
            }
        ),

        # ============================================================
        # 格式转换（基于pandoc）
        # ============================================================
        Tool(
            name="convert",
            description="通用文档格式转换，支持40+格式互转（docx/pdf/md/html/latex/epub等）",
            inputSchema={
                "type": "object",
                "properties": {
                    "source": {
                        "type": "string",
                        "description": "源文件路径"
                    },
                    "target": {
                        "type": "string",
                        "description": "目标文件路径（可选，不指定则自动生成）"
                    },
                    "source_format": {
                        "type": "string",
                        "description": "源格式（可选，自动检测）"
                    },
                    "target_format": {
                        "type": "string",
                        "description": "目标格式（可选，从target路径推断）"
                    },
                    "extra_args": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "pandoc额外参数，如['--toc', '--standalone']"
                    }
                },
                "required": ["source"]
            }
        ),

        Tool(
            name="convert_batch",
            description="批量转换多个文件到指定格式",
            inputSchema={
                "type": "object",
                "properties": {
                    "sources": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "源文件路径列表"
                    },
                    "target_format": {
                        "type": "string",
                        "description": "目标格式（如 pdf, docx, html, markdown）"
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "输出目录（可选，默认与源文件同目录）"
                    },
                    "extra_args": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "pandoc额外参数"
                    }
                },
                "required": ["sources", "target_format"]
            }
        ),

        Tool(
            name="convert_formats",
            description="获取支持的格式列表和常用转换组合",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        Tool(
            name="convert_with_template",
            description="带模板/样式的文档转换（支持自定义CSS、pandoc模板、参考文档）",
            inputSchema={
                "type": "object",
                "properties": {
                    "source": {
                        "type": "string",
                        "description": "源文件路径"
                    },
                    "target": {
                        "type": "string",
                        "description": "目标文件路径"
                    },
                    "template": {
                        "type": "string",
                        "description": "pandoc模板文件路径（用于HTML/LaTeX输出）"
                    },
                    "css": {
                        "type": "string",
                        "description": "CSS样式文件路径（用于HTML输出）"
                    },
                    "reference_doc": {
                        "type": "string",
                        "description": "参考文档路径（用于docx/pptx，继承样式）"
                    },
                    "extra_args": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "其他pandoc参数"
                    }
                },
                "required": ["source", "target"]
            }
        ),

        # ============================================================
        # OCR识别（支持扫描件PDF和图片）
        # ============================================================
        Tool(
            name="ocr_image",
            description="OCR识别单张图片中的文字（支持中英日韩等多语言）",
            inputSchema={
                "type": "object",
                "properties": {
                    "image_path": {
                        "type": "string",
                        "description": "图片文件路径（支持png/jpg/tiff/bmp等）"
                    },
                    "lang": {
                        "type": "string",
                        "description": "识别语言（auto/chinese/english/japanese等），默认auto",
                        "default": "auto"
                    },
                    "engine": {
                        "type": "string",
                        "enum": ["auto", "tesseract", "claude"],
                        "description": "OCR引擎（auto自动选择，tesseract本地免费，claude AI增强）",
                        "default": "auto"
                    },
                    "api_key": {
                        "type": "string",
                        "description": "Claude API密钥（可选，也可用ANTHROPIC_API_KEY环境变量）"
                    },
                    "prompt": {
                        "type": "string",
                        "description": "自定义Claude识别提示词（仅claude引擎有效）"
                    }
                },
                "required": ["image_path"]
            }
        ),

        Tool(
            name="ocr_pdf",
            description="OCR识别PDF文档（支持扫描件/图像PDF）",
            inputSchema={
                "type": "object",
                "properties": {
                    "pdf_path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "要识别的页码列表（从1开始，不指定则识别全部）"
                    },
                    "lang": {
                        "type": "string",
                        "description": "识别语言",
                        "default": "auto"
                    },
                    "engine": {
                        "type": "string",
                        "enum": ["auto", "tesseract", "claude"],
                        "description": "OCR引擎",
                        "default": "auto"
                    },
                    "dpi": {
                        "type": "integer",
                        "description": "PDF转图片的DPI（越高越清晰但越慢，默认200）",
                        "default": 200
                    },
                    "api_key": {
                        "type": "string",
                        "description": "Claude API密钥"
                    },
                    "prompt": {
                        "type": "string",
                        "description": "自定义识别提示词"
                    }
                },
                "required": ["pdf_path"]
            }
        ),

        Tool(
            name="ocr_to_docx",
            description="OCR识别后直接生成Word文档（一步到位）",
            inputSchema={
                "type": "object",
                "properties": {
                    "source": {
                        "type": "string",
                        "description": "源文件路径（PDF或图片）"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出Word文档路径（.docx）"
                    },
                    "lang": {
                        "type": "string",
                        "description": "识别语言",
                        "default": "auto"
                    },
                    "engine": {
                        "type": "string",
                        "enum": ["auto", "tesseract", "claude"],
                        "description": "OCR引擎",
                        "default": "auto"
                    },
                    "dpi": {
                        "type": "integer",
                        "description": "PDF转换DPI",
                        "default": 200
                    },
                    "api_key": {
                        "type": "string",
                        "description": "Claude API密钥"
                    },
                    "prompt": {
                        "type": "string",
                        "description": "自定义识别提示词"
                    },
                    "title": {
                        "type": "string",
                        "description": "生成文档的标题"
                    }
                },
                "required": ["source", "output_path"]
            }
        ),

        Tool(
            name="ocr_status",
            description="获取OCR模块状态（可用引擎、支持的语言和格式）",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        # ============================================================
        # Excel表格操作
        # ============================================================
        Tool(
            name="excel_create",
            description="创建新的Excel文件（.xlsx）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文件保存路径，必须以.xlsx结尾"
                    },
                    "sheets": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "工作表名称列表，默认['Sheet1']"
                    },
                    "title": {
                        "type": "string",
                        "description": "文档标题（元数据）"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="excel_read",
            description="读取Excel文件内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（默认活动表）"
                    },
                    "range": {
                        "type": "string",
                        "description": "读取范围如'A1:D10'（默认全部）"
                    },
                    "include_formatting": {
                        "type": "boolean",
                        "description": "是否包含格式信息",
                        "default": False
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="excel_info",
            description="获取Excel工作簿信息（工作表列表、属性、统计）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="excel_save_as",
            description="Excel文件另存为（支持xlsx/csv格式）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "源文件路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "目标文件路径"
                    },
                    "format": {
                        "type": "string",
                        "enum": ["xlsx", "csv"],
                        "description": "目标格式（可选，从路径推断）"
                    }
                },
                "required": ["path", "output_path"]
            }
        ),

        Tool(
            name="sheet_list",
            description="列出Excel文件中所有工作表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="sheet_add",
            description="向Excel文件添加新工作表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "name": {
                        "type": "string",
                        "description": "新工作表名称"
                    },
                    "position": {
                        "type": "integer",
                        "description": "插入位置（可选，默认末尾）"
                    }
                },
                "required": ["path", "name"]
            }
        ),

        Tool(
            name="sheet_delete",
            description="删除Excel工作表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "name": {
                        "type": "string",
                        "description": "要删除的工作表名称"
                    }
                },
                "required": ["path", "name"]
            }
        ),

        Tool(
            name="sheet_rename",
            description="重命名Excel工作表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "old_name": {
                        "type": "string",
                        "description": "原工作表名称"
                    },
                    "new_name": {
                        "type": "string",
                        "description": "新工作表名称"
                    }
                },
                "required": ["path", "old_name", "new_name"]
            }
        ),

        Tool(
            name="sheet_copy",
            description="复制Excel工作表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "source_name": {
                        "type": "string",
                        "description": "源工作表名称"
                    },
                    "target_name": {
                        "type": "string",
                        "description": "目标工作表名称"
                    }
                },
                "required": ["path", "source_name", "target_name"]
            }
        ),

        Tool(
            name="cell_read",
            description="读取Excel单元格或区域的值",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "cell": {
                        "type": "string",
                        "description": "单个单元格如'A1'"
                    },
                    "range": {
                        "type": "string",
                        "description": "区域如'A1:D10'"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="cell_write",
            description="写入Excel单元格或区域",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "cell": {
                        "type": "string",
                        "description": "单个单元格如'A1'"
                    },
                    "value": {
                        "description": "单个值（与cell配合）"
                    },
                    "range": {
                        "type": "string",
                        "description": "区域起点如'A1'（与data配合）"
                    },
                    "data": {
                        "type": "array",
                        "items": {"type": "array"},
                        "description": "二维数组数据"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="cell_format",
            description="设置Excel单元格格式（字体、颜色、边框、对齐等）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "格式化范围如'A1:D10'或'A1'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "字体名称"
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "字号"
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "是否加粗"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "是否斜体"
                    },
                    "font_color": {
                        "type": "string",
                        "description": "字体颜色（十六进制如'FF0000'）"
                    },
                    "bg_color": {
                        "type": "string",
                        "description": "背景颜色"
                    },
                    "border": {
                        "type": "string",
                        "enum": ["thin", "medium", "thick"],
                        "description": "边框样式"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right"],
                        "description": "对齐方式"
                    },
                    "number_format": {
                        "type": "string",
                        "description": "数字格式如'0.00%', '#,##0'"
                    }
                },
                "required": ["path", "range"]
            }
        ),

        Tool(
            name="cell_merge",
            description="合并或取消合并Excel单元格",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "合并范围如'A1:D1'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "unmerge": {
                        "type": "boolean",
                        "description": "是否取消合并（默认false为合并）",
                        "default": False
                    }
                },
                "required": ["path", "range"]
            }
        ),

        Tool(
            name="cell_formula",
            description="设置Excel单元格公式",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "cell": {
                        "type": "string",
                        "description": "单元格位置如'E1'"
                    },
                    "formula": {
                        "type": "string",
                        "description": "公式如'=SUM(A1:D1)'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    }
                },
                "required": ["path", "cell", "formula"]
            }
        ),

        Tool(
            name="row_insert",
            description="在Excel中插入行",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "row": {
                        "type": "integer",
                        "description": "插入位置（行号，从1开始）"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "count": {
                        "type": "integer",
                        "description": "插入行数（默认1）",
                        "default": 1
                    }
                },
                "required": ["path", "row"]
            }
        ),

        Tool(
            name="row_delete",
            description="删除Excel中的行",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "row": {
                        "type": "integer",
                        "description": "删除起始行号"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "count": {
                        "type": "integer",
                        "description": "删除行数（默认1）",
                        "default": 1
                    }
                },
                "required": ["path", "row"]
            }
        ),

        Tool(
            name="col_insert",
            description="在Excel中插入列",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "col": {
                        "description": "插入位置（列号1或列字母'A'）"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "count": {
                        "type": "integer",
                        "description": "插入列数（默认1）",
                        "default": 1
                    }
                },
                "required": ["path", "col"]
            }
        ),

        Tool(
            name="col_delete",
            description="删除Excel中的列",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "col": {
                        "description": "删除起始列（列号1或列字母'A'）"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "count": {
                        "type": "integer",
                        "description": "删除列数（默认1）",
                        "default": 1
                    }
                },
                "required": ["path", "col"]
            }
        ),

        Tool(
            name="excel_to_word",
            description="将Excel表格插入Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "excel_path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "word_path": {
                        "type": "string",
                        "description": "Word文档路径（已存在则追加，不存在则创建）"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "range": {
                        "type": "string",
                        "description": "数据范围如'A1:D10'（可选，默认全部）"
                    },
                    "style": {
                        "type": "string",
                        "description": "Word表格样式（如'Table Grid'）"
                    }
                },
                "required": ["excel_path", "word_path"]
            }
        ),

        Tool(
            name="excel_status",
            description="获取Excel模块状态（openpyxl可用性、版本、功能）",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        # ============================================================
        # Excel高级功能：公式增强
        # ============================================================
        Tool(
            name="formula_batch",
            description="批量设置公式，支持行号占位符{row}自动填充多行",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "目标范围如'E2:E100'"
                    },
                    "formula": {
                        "type": "string",
                        "description": "公式模板，支持{row}占位符，如'=SUM(A{row}:D{row})'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    }
                },
                "required": ["path", "range", "formula"]
            }
        ),

        Tool(
            name="formula_quick",
            description="快捷函数生成，一键生成SUM/AVERAGE/MAX/MIN/COUNT等统计公式",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "data_range": {
                        "type": "string",
                        "description": "数据范围如'A1:A100'"
                    },
                    "function": {
                        "type": "string",
                        "enum": ["sum", "average", "avg", "max", "min", "count", "counta", "stdev", "var", "median"],
                        "description": "函数类型"
                    },
                    "output_cell": {
                        "type": "string",
                        "description": "输出单元格如'B1'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    }
                },
                "required": ["path", "data_range", "function", "output_cell"]
            }
        ),

        # ============================================================
        # Excel高级功能：数据操作
        # ============================================================
        Tool(
            name="data_sort",
            description="数据排序，支持多列、升序/降序",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "排序范围如'A1:D100'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "sort_by": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "col": {"type": "string", "description": "列字母如'C'"},
                                "order": {"type": "string", "enum": ["asc", "desc"], "description": "排序方向"}
                            }
                        },
                        "description": "排序规则列表，如[{\"col\": \"C\", \"order\": \"desc\"}]"
                    },
                    "has_header": {
                        "type": "boolean",
                        "description": "是否有标题行",
                        "default": True
                    }
                },
                "required": ["path", "range"]
            }
        ),

        Tool(
            name="data_filter",
            description="自动筛选，设置或清除筛选",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "筛选范围如'A1:D100'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "clear": {
                        "type": "boolean",
                        "description": "是否清除筛选（默认false为设置筛选）",
                        "default": False
                    }
                },
                "required": ["path", "range"]
            }
        ),

        Tool(
            name="data_validate",
            description="数据验证，支持下拉列表、数值范围、日期、文本长度等",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "验证范围如'B2:B100'"
                    },
                    "type": {
                        "type": "string",
                        "enum": ["list", "whole", "decimal", "date", "text_length", "custom"],
                        "description": "验证类型"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "values": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "下拉列表值（type=list时使用）"
                    },
                    "min_val": {
                        "type": "number",
                        "description": "最小值"
                    },
                    "max_val": {
                        "type": "number",
                        "description": "最大值"
                    },
                    "formula": {
                        "type": "string",
                        "description": "自定义公式（type=custom时使用）"
                    },
                    "error_message": {
                        "type": "string",
                        "description": "错误提示信息"
                    }
                },
                "required": ["path", "range", "type"]
            }
        ),

        Tool(
            name="data_deduplicate",
            description="去除重复行，基于指定列去重",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "数据范围如'A1:D100'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "columns": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "用于判断重复的列，如['A', 'B']，不指定则全部列"
                    },
                    "keep": {
                        "type": "string",
                        "enum": ["first", "last"],
                        "description": "保留策略",
                        "default": "first"
                    }
                },
                "required": ["path", "range"]
            }
        ),

        Tool(
            name="data_fill",
            description="序列填充，支持等差、等比、日期序列",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "填充范围如'A1:A10'"
                    },
                    "type": {
                        "type": "string",
                        "enum": ["linear", "growth", "date"],
                        "description": "填充类型：linear等差、growth等比、date日期"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "start": {
                        "type": "number",
                        "description": "起始值（默认1）"
                    },
                    "step": {
                        "type": "number",
                        "description": "步长（等差）或比率（等比），默认1"
                    }
                },
                "required": ["path", "range", "type"]
            }
        ),

        # ============================================================
        # Excel高级功能：统计与格式
        # ============================================================
        Tool(
            name="stats_summary",
            description="统计摘要，一键生成和/均值/最大/最小/计数/标准差等",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "data_range": {
                        "type": "string",
                        "description": "数据范围"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "output_cell": {
                        "type": "string",
                        "description": "输出起始单元格（可选，不指定则只返回结果）"
                    },
                    "metrics": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "统计指标列表：sum/average/max/min/count/stdev/var/median"
                    }
                },
                "required": ["path", "data_range"]
            }
        ),

        Tool(
            name="conditional_format",
            description="条件格式，支持高亮规则、色阶、数据条",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "range": {
                        "type": "string",
                        "description": "格式化范围"
                    },
                    "rule": {
                        "type": "string",
                        "enum": ["greater_than", "less_than", "equal", "between", "not_between", "color_scale", "data_bar"],
                        "description": "规则类型"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "value": {
                        "description": "比较值"
                    },
                    "value2": {
                        "description": "第二个比较值（between时使用）"
                    },
                    "format": {
                        "type": "object",
                        "description": "格式设置：{\"bg_color\": \"FF0000\", \"font_color\": \"FFFFFF\", \"bold\": true}"
                    },
                    "color_scale": {
                        "type": "object",
                        "description": "色阶设置：{\"min_color\": \"F8696B\", \"max_color\": \"63BE7B\"}"
                    },
                    "data_bar": {
                        "type": "object",
                        "description": "数据条设置：{\"color\": \"638EC6\"}"
                    }
                },
                "required": ["path", "range", "rule"]
            }
        ),

        Tool(
            name="named_range",
            description="命名范围操作，创建/列出/删除命名范围",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "action": {
                        "type": "string",
                        "enum": ["create", "list", "delete"],
                        "description": "操作类型"
                    },
                    "name": {
                        "type": "string",
                        "description": "范围名称（create/delete时需要）"
                    },
                    "range": {
                        "type": "string",
                        "description": "单元格范围（create时需要）"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    }
                },
                "required": ["path", "action"]
            }
        ),

        # ============================================================
        # Excel高级功能：图表
        # ============================================================
        Tool(
            name="chart_create",
            description="创建图表，支持柱状图/折线图/饼图/散点图等",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "type": {
                        "type": "string",
                        "enum": ["bar", "column", "line", "pie", "scatter", "area", "doughnut", "radar"],
                        "description": "图表类型"
                    },
                    "data_range": {
                        "type": "string",
                        "description": "数据范围如'A1:B10'"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "position": {
                        "type": "string",
                        "description": "图表位置如'E1'"
                    },
                    "title": {
                        "type": "string",
                        "description": "图表标题"
                    },
                    "x_title": {
                        "type": "string",
                        "description": "X轴标题"
                    },
                    "y_title": {
                        "type": "string",
                        "description": "Y轴标题"
                    },
                    "style": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 48,
                        "description": "图表样式编号(1-48)"
                    }
                },
                "required": ["path", "type", "data_range"]
            }
        ),

        Tool(
            name="excel_chart_modify",
            description="修改Excel图表属性，如标题、样式、尺寸",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "chart_index": {
                        "type": "integer",
                        "description": "图表索引（从0开始）"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "工作表名称（可选）"
                    },
                    "title": {
                        "type": "string",
                        "description": "新标题"
                    },
                    "x_title": {
                        "type": "string",
                        "description": "X轴标题"
                    },
                    "y_title": {
                        "type": "string",
                        "description": "Y轴标题"
                    },
                    "style": {
                        "type": "integer",
                        "description": "新样式编号"
                    },
                    "width": {
                        "type": "number",
                        "description": "图表宽度（厘米）"
                    },
                    "height": {
                        "type": "number",
                        "description": "图表高度（厘米）"
                    }
                },
                "required": ["path", "chart_index"]
            }
        ),

        # Excel数据透视汇总
        Tool(
            name="pivot_create",
            description="创建数据透视汇总表，按指定字段分组聚合（支持sum/average/count/max/min）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel文件路径"
                    },
                    "source_range": {
                        "type": "string",
                        "description": "数据源范围如'A1:D100'（第一行为表头）"
                    },
                    "target_cell": {
                        "type": "string",
                        "description": "输出起始单元格如'F1'"
                    },
                    "rows": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "分组字段名列表（表头名称）"
                    },
                    "values": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "聚合值字段名列表（表头名称）"
                    },
                    "sheet": {
                        "type": "string",
                        "description": "源数据工作表名称（可选）"
                    },
                    "agg_func": {
                        "type": "string",
                        "enum": ["sum", "average", "count", "max", "min"],
                        "default": "sum",
                        "description": "聚合函数"
                    },
                    "target_sheet": {
                        "type": "string",
                        "description": "输出工作表名称（默认同源）"
                    },
                    "include_totals": {
                        "type": "boolean",
                        "default": True,
                        "description": "是否包含总计行"
                    }
                },
                "required": ["path", "source_range", "target_cell", "rows", "values"]
            }
        ),

        # ============================================================
        # PDF操作：信息与提取
        # ============================================================
        Tool(
            name="pdf_info",
            description="获取PDF文件的基本信息，包括页数、元数据、文件大小等",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="pdf_extract_text",
            description="提取PDF文档的文本内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码列表（从1开始），不指定则提取全部"
                    },
                    "layout": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否保留原始布局"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="pdf_extract_tables",
            description="提取PDF中的表格数据",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码列表"
                    },
                    "format": {
                        "type": "string",
                        "enum": ["json", "csv", "list"],
                        "default": "json",
                        "description": "输出格式"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="pdf_extract_images",
            description="提取PDF中的图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码列表"
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "输出目录（不指定则只返回信息不保存）"
                    },
                    "format": {
                        "type": "string",
                        "enum": ["png", "jpg"],
                        "default": "png",
                        "description": "输出图片格式"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="pdf_get_outline",
            description="获取PDF大纲/书签结构",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    }
                },
                "required": ["path"]
            }
        ),

        # ============================================================
        # PDF操作：文件操作
        # ============================================================
        Tool(
            name="pdf_merge",
            description="合并多个PDF文件为一个",
            inputSchema={
                "type": "object",
                "properties": {
                    "paths": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "要合并的PDF文件路径列表"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文件路径"
                    },
                    "add_outline": {
                        "type": "boolean",
                        "default": True,
                        "description": "是否为每个源文件添加书签"
                    }
                },
                "required": ["paths", "output_path"]
            }
        ),

        Tool(
            name="pdf_split",
            description="拆分PDF为多个文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "输出目录"
                    },
                    "mode": {
                        "type": "string",
                        "enum": ["single", "range"],
                        "default": "single",
                        "description": "拆分模式：single(每页一个文件), range(按页数范围)"
                    },
                    "pages_per_file": {
                        "type": "integer",
                        "minimum": 1,
                        "default": 1,
                        "description": "每个文件的页数（mode=range时）"
                    }
                },
                "required": ["path", "output_dir"]
            }
        ),

        Tool(
            name="pdf_extract_pages",
            description="提取PDF指定页面到新文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "要提取的页码列表（从1开始）"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文件路径"
                    }
                },
                "required": ["path", "pages", "output_path"]
            }
        ),

        Tool(
            name="pdf_rotate",
            description="旋转PDF页面",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "angle": {
                        "type": "integer",
                        "enum": [90, 180, 270],
                        "description": "旋转角度（顺时针）"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码，不指定则旋转全部页面"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    }
                },
                "required": ["path", "angle"]
            }
        ),

        Tool(
            name="pdf_delete_pages",
            description="删除PDF指定页面",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "要删除的页码列表（从1开始）"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    }
                },
                "required": ["path", "pages"]
            }
        ),

        Tool(
            name="pdf_add_watermark",
            description="为PDF添加文字水印",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "watermark": {
                        "type": "string",
                        "description": "水印文字"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码，不指定则添加到全部页面"
                    },
                    "position": {
                        "type": "string",
                        "enum": ["center", "diagonal"],
                        "default": "center",
                        "description": "水印位置"
                    },
                    "opacity": {
                        "type": "number",
                        "minimum": 0,
                        "maximum": 1,
                        "default": 0.3,
                        "description": "透明度(0-1)"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    }
                },
                "required": ["path", "watermark"]
            }
        ),

        # ============================================================
        # PDF操作：转换与集成
        # ============================================================
        Tool(
            name="pdf_tables_to_word",
            description="将PDF中的表格提取并转换为Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "pdf_path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "word_path": {
                        "type": "string",
                        "description": "输出Word文档路径"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码"
                    },
                    "table_style": {
                        "type": "string",
                        "default": "Table Grid",
                        "description": "Word表格样式"
                    }
                },
                "required": ["pdf_path", "word_path"]
            }
        ),

        Tool(
            name="pdf_tables_to_excel",
            description="将PDF中的表格提取并转换为Excel文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "pdf_path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "excel_path": {
                        "type": "string",
                        "description": "输出Excel文件路径"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码"
                    },
                    "sheet_per_table": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否每个表格一个工作表"
                    }
                },
                "required": ["pdf_path", "excel_path"]
            }
        ),

        Tool(
            name="pdf_to_text",
            description="将PDF转换为纯文本文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文本文件路径（可选，不指定则只返回文本内容）"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="pdf_status",
            description="获取PDF模块状态（依赖库可用性、版本、功能列表）",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        # ============================================================
        # PDF编辑操作
        # ============================================================
        Tool(
            name="pdf_to_editable",
            description="将PDF转换为可编辑格式（Word或Markdown），提取文本和表格内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文件路径（可选，自动生成）"
                    },
                    "format": {
                        "type": "string",
                        "enum": ["docx", "markdown", "md"],
                        "default": "docx",
                        "description": "输出格式"
                    },
                    "include_tables": {
                        "type": "boolean",
                        "default": True,
                        "description": "是否包含表格"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="pdf_text_replace",
            description="在PDF中查找并替换文字（通过覆盖绘制方式，适合简单场景）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "old_text": {
                        "type": "string",
                        "description": "要替换的原文字"
                    },
                    "new_text": {
                        "type": "string",
                        "description": "替换后的新文字"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码列表（从1开始），不指定则处理全部"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    },
                    "font_name": {
                        "type": "string",
                        "default": "Helvetica",
                        "description": "字体名称"
                    },
                    "font_size": {
                        "type": "number",
                        "default": 12,
                        "description": "字号"
                    }
                },
                "required": ["path", "old_text", "new_text"]
            }
        ),

        Tool(
            name="pdf_redact",
            description="涂黑/删除PDF中的指定文字（用于敏感信息脱敏）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "text": {
                        "type": "string",
                        "description": "要涂黑的文字"
                    },
                    "pages": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "指定页码列表，不指定则处理全部"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    },
                    "redact_color": {
                        "type": "string",
                        "enum": ["black", "white", "gray"],
                        "default": "black",
                        "description": "涂黑颜色"
                    }
                },
                "required": ["path", "text"]
            }
        ),

        Tool(
            name="pdf_annotate_text",
            description="在PDF指定位置添加文字注释或标注",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "text": {
                        "type": "string",
                        "description": "要添加的文字"
                    },
                    "x": {
                        "type": "number",
                        "description": "X坐标（从左边开始，单位：点）"
                    },
                    "y": {
                        "type": "number",
                        "description": "Y坐标（从底部开始，单位：点）"
                    },
                    "page": {
                        "type": "integer",
                        "default": 1,
                        "description": "页码（从1开始）"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    },
                    "font_name": {
                        "type": "string",
                        "default": "Helvetica",
                        "description": "字体名称"
                    },
                    "font_size": {
                        "type": "number",
                        "default": 12,
                        "description": "字号"
                    },
                    "font_color": {
                        "type": "string",
                        "enum": ["black", "red", "blue", "green", "gray", "white"],
                        "default": "black",
                        "description": "字体颜色"
                    }
                },
                "required": ["path", "text", "x", "y"]
            }
        ),

        Tool(
            name="pdf_encrypt",
            description="加密PDF文件，设置用户密码和所有者密码",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "user_password": {
                        "type": "string",
                        "description": "用户密码（打开文档需要）"
                    },
                    "owner_password": {
                        "type": "string",
                        "description": "所有者密码（修改权限需要，默认与用户密码相同）"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    },
                    "algorithm": {
                        "type": "string",
                        "enum": ["AES-256", "AES-128", "RC4-128", "RC4-40"],
                        "default": "AES-256",
                        "description": "加密算法"
                    }
                },
                "required": ["path", "user_password"]
            }
        ),

        Tool(
            name="pdf_decrypt",
            description="解密PDF文件，移除密码保护",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "加密的PDF文件路径"
                    },
                    "password": {
                        "type": "string",
                        "description": "密码"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    }
                },
                "required": ["path", "password"]
            }
        ),

        Tool(
            name="pdf_form_get_fields",
            description="获取PDF表单字段信息（仅支持AcroForm，不支持XFA）",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "password": {
                        "type": "string",
                        "description": "密码（如果PDF加密）"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="pdf_form_fill",
            description="填写PDF表单字段",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "fields": {
                        "type": "object",
                        "description": "字段值字典，如 {\"name\": \"张三\", \"age\": \"30\"}"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径，不指定则覆盖原文件"
                    },
                    "password": {
                        "type": "string",
                        "description": "密码（如果PDF加密）"
                    },
                    "flatten": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否扁平化表单（填写后不可编辑）"
                    }
                },
                "required": ["path", "fields"]
            }
        ),

        # ============================================================
        # PPT操作
        # ============================================================
        Tool(
            name="ppt_create",
            description="创建新的PowerPoint文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "文件保存路径,必须以.pptx结尾"
                    },
                    "title": {
                        "type": "string",
                        "description": "可选的文档标题(元数据)"
                    },
                    "width": {
                        "type": "string",
                        "description": "幻灯片宽度,如 '10in', '25.4cm'"
                    },
                    "height": {
                        "type": "string",
                        "description": "幻灯片高度,如 '7.5in', '19.05cm'"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="ppt_read",
            description="读取PowerPoint文档内容,返回所有幻灯片的文本和结构信息",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "include_notes": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否包含演讲者备注"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="ppt_info",
            description="获取PowerPoint文档基本信息(幻灯片数、尺寸、属性等)",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="ppt_set_properties",
            description="设置PowerPoint文档属性(标题、作者、主题等)",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "title": {
                        "type": "string",
                        "description": "文档标题"
                    },
                    "author": {
                        "type": "string",
                        "description": "作者"
                    },
                    "subject": {
                        "type": "string",
                        "description": "主题"
                    },
                    "keywords": {
                        "type": "string",
                        "description": "关键词"
                    },
                    "comments": {
                        "type": "string",
                        "description": "备注"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="ppt_merge",
            description="合并多个PowerPoint文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "paths": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "要合并的PPT文件路径列表"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出文件路径"
                    }
                },
                "required": ["paths", "output_path"]
            }
        ),

        Tool(
            name="slide_add",
            description="添加新幻灯片",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "layout": {
                        "type": "string",
                        "description": "布局名称,如 'Title Slide', 'Title and Content', 'Blank' 等"
                    },
                    "index": {
                        "type": "integer",
                        "description": "插入位置(从1开始),不指定则添加到末尾"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="slide_delete",
            description="删除指定幻灯片",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "index": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    }
                },
                "required": ["path", "index"]
            }
        ),

        Tool(
            name="slide_duplicate",
            description="复制指定幻灯片",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "index": {
                        "type": "integer",
                        "description": "要复制的幻灯片索引(从1开始)"
                    }
                },
                "required": ["path", "index"]
            }
        ),

        Tool(
            name="slide_get_layouts",
            description="获取PPT中可用的幻灯片布局列表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="shape_add_text",
            description="在幻灯片中添加文本框",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "text": {
                        "type": "string",
                        "description": "文本内容"
                    },
                    "left": {
                        "type": "string",
                        "default": "1in",
                        "description": "左边距,如 '1in', '2.54cm'"
                    },
                    "top": {
                        "type": "string",
                        "default": "1in",
                        "description": "上边距"
                    },
                    "width": {
                        "type": "string",
                        "default": "8in",
                        "description": "文本框宽度"
                    },
                    "height": {
                        "type": "string",
                        "default": "1in",
                        "description": "文本框高度"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "字体名称"
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "字号(磅)"
                    },
                    "bold": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否加粗"
                    },
                    "italic": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否斜体"
                    },
                    "color": {
                        "type": "string",
                        "description": "字体颜色(十六进制,如 'FF0000')"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right"],
                        "description": "对齐方式"
                    }
                },
                "required": ["path", "slide", "text"]
            }
        ),

        Tool(
            name="shape_add_image",
            description="在幻灯片中添加图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "image_path": {
                        "type": "string",
                        "description": "图片文件路径"
                    },
                    "left": {
                        "type": "string",
                        "default": "1in",
                        "description": "左边距"
                    },
                    "top": {
                        "type": "string",
                        "default": "1in",
                        "description": "上边距"
                    },
                    "width": {
                        "type": "string",
                        "description": "图片宽度(可选,保持比例)"
                    },
                    "height": {
                        "type": "string",
                        "description": "图片高度(可选)"
                    }
                },
                "required": ["path", "slide", "image_path"]
            }
        ),

        Tool(
            name="shape_add_table",
            description="在幻灯片中添加表格",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "rows": {
                        "type": "integer",
                        "description": "行数"
                    },
                    "cols": {
                        "type": "integer",
                        "description": "列数"
                    },
                    "left": {
                        "type": "string",
                        "default": "1in",
                        "description": "左边距"
                    },
                    "top": {
                        "type": "string",
                        "default": "2in",
                        "description": "上边距"
                    },
                    "width": {
                        "type": "string",
                        "default": "8in",
                        "description": "表格宽度"
                    },
                    "height": {
                        "type": "string",
                        "default": "3in",
                        "description": "表格高度"
                    },
                    "data": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "string"}},
                        "description": "表格数据(二维数组)"
                    }
                },
                "required": ["path", "slide", "rows", "cols"]
            }
        ),

        Tool(
            name="shape_add_shape",
            description="在幻灯片中添加形状(矩形、圆形、箭头等)",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "shape_type": {
                        "type": "string",
                        "enum": ["rectangle", "rounded_rectangle", "oval", "triangle", "diamond", "pentagon", "hexagon", "arrow_right", "arrow_left", "arrow_up", "arrow_down", "star", "heart", "cloud"],
                        "description": "形状类型"
                    },
                    "left": {
                        "type": "string",
                        "default": "2in",
                        "description": "左边距"
                    },
                    "top": {
                        "type": "string",
                        "default": "2in",
                        "description": "上边距"
                    },
                    "width": {
                        "type": "string",
                        "default": "2in",
                        "description": "形状宽度"
                    },
                    "height": {
                        "type": "string",
                        "default": "2in",
                        "description": "形状高度"
                    },
                    "fill_color": {
                        "type": "string",
                        "description": "填充颜色(十六进制)"
                    },
                    "line_color": {
                        "type": "string",
                        "description": "边框颜色(十六进制)"
                    },
                    "text": {
                        "type": "string",
                        "description": "形状内的文字"
                    }
                },
                "required": ["path", "slide", "shape_type"]
            }
        ),

        Tool(
            name="slide_set_background",
            description="设置幻灯片背景颜色",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "color": {
                        "type": "string",
                        "description": "背景颜色(十六进制,如 'FFFFFF')"
                    },
                    "image_path": {
                        "type": "string",
                        "description": "背景图片路径(暂不支持)"
                    }
                },
                "required": ["path", "slide"]
            }
        ),

        Tool(
            name="slide_add_notes",
            description="为幻灯片添加演讲者备注",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "notes": {
                        "type": "string",
                        "description": "备注内容"
                    }
                },
                "required": ["path", "slide", "notes"]
            }
        ),

        Tool(
            name="ppt_status",
            description="获取PPT模块状态(依赖库可用性、版本、功能列表)",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        # ============================================================
        # PPT母版操作
        # ============================================================
        Tool(
            name="master_list",
            description="列出所有母版和布局",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="master_get_info",
            description="获取母版详细信息(形状、占位符、布局)",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "master_index": {
                        "type": "integer",
                        "default": 0,
                        "description": "母版索引(默认0)"
                    }
                },
                "required": ["path"]
            }
        ),

        Tool(
            name="placeholder_list",
            description="列出幻灯片中的占位符",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    }
                },
                "required": ["path", "slide"]
            }
        ),

        Tool(
            name="placeholder_set",
            description="设置占位符内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "idx": {
                        "type": "integer",
                        "description": "占位符索引"
                    },
                    "text": {
                        "type": "string",
                        "description": "文本内容"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "字体名称"
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "字号"
                    },
                    "bold": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否加粗"
                    },
                    "italic": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否斜体"
                    },
                    "color": {
                        "type": "string",
                        "description": "字体颜色(十六进制)"
                    }
                },
                "required": ["path", "slide", "idx"]
            }
        ),

        # ============================================================
        # PPT动画操作
        # ============================================================
        Tool(
            name="animation_add",
            description="为形状添加动画效果(进入/强调/退出)",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "形状索引(从0开始)"
                    },
                    "effect": {
                        "type": "string",
                        "enum": ["appear", "fade", "fly_in", "float_in", "zoom", "wipe", "split", "wheel",
                                 "pulse", "spin", "grow_shrink",
                                 "disappear", "fade_out", "fly_out", "zoom_out"],
                        "description": "动画效果类型"
                    },
                    "trigger": {
                        "type": "string",
                        "enum": ["on_click", "with_previous", "after_previous"],
                        "default": "on_click",
                        "description": "触发方式"
                    },
                    "duration": {
                        "type": "number",
                        "default": 0.5,
                        "description": "持续时间(秒)"
                    },
                    "delay": {
                        "type": "number",
                        "default": 0.0,
                        "description": "延迟时间(秒)"
                    },
                    "direction": {
                        "type": "string",
                        "enum": ["left", "right", "top", "bottom", "up", "down"],
                        "description": "方向(仅fly_in/fly_out等支持)"
                    }
                },
                "required": ["path", "slide", "shape_index", "effect"]
            }
        ),

        Tool(
            name="animation_list",
            description="列出幻灯片上的动画",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    }
                },
                "required": ["path", "slide"]
            }
        ),

        Tool(
            name="animation_remove",
            description="删除动画",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "形状索引(删除该形状的动画)"
                    },
                    "remove_all": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否删除所有动画"
                    }
                },
                "required": ["path", "slide"]
            }
        ),

        # PPT幻灯片切换效果
        Tool(
            name="slide_set_transition",
            description="设置幻灯片切换效果(fade/push/wipe/split/dissolve等)",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "effect": {
                        "type": "string",
                        "enum": ["fade", "push", "wipe", "split", "cut", "dissolve", "cover", "uncover", "randomBars", "blinds", "wheel", "comb", "checker", "random", "strips", "plus", "circle", "diamond", "wedge", "zoom"],
                        "description": "切换效果"
                    },
                    "speed": {
                        "type": "string",
                        "enum": ["slow", "medium", "fast"],
                        "default": "medium",
                        "description": "切换速度"
                    },
                    "advance_click": {
                        "type": "boolean",
                        "default": True,
                        "description": "是否点击切换"
                    },
                    "advance_time": {
                        "type": "integer",
                        "description": "自动切换时间(毫秒)"
                    },
                    "duration": {
                        "type": "integer",
                        "description": "切换持续时间(毫秒)"
                    }
                },
                "required": ["path", "slide", "effect"]
            }
        ),

        Tool(
            name="slide_remove_transition",
            description="移除幻灯片切换效果",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    }
                },
                "required": ["path", "slide"]
            }
        ),

        # PPT图表工具
        Tool(
            name="chart_add",
            description="向幻灯片添加图表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "chart_type": {
                        "type": "string",
                        "description": "图表类型(column/bar/line/pie/area/scatter/bubble/doughnut/radar)",
                        "enum": ["column", "column_stacked", "bar", "bar_stacked", "line", "line_markers", "pie", "pie_exploded", "doughnut", "area", "area_stacked", "scatter", "scatter_lines", "bubble", "radar", "radar_filled"]
                    },
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "分类标签列表，如['Q1', 'Q2', 'Q3']"
                    },
                    "series": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": {"type": "string"},
                                "values": {"type": "array"}
                            }
                        },
                        "description": "系列数据列表，如[{\"name\": \"销售额\", \"values\": [100, 200, 150]}]"
                    },
                    "x": {
                        "type": "string",
                        "description": "图表左边距，如'1in', '2.5cm'"
                    },
                    "y": {
                        "type": "string",
                        "description": "图表上边距"
                    },
                    "width": {
                        "type": "string",
                        "description": "图表宽度"
                    },
                    "height": {
                        "type": "string",
                        "description": "图表高度"
                    },
                    "title": {
                        "type": "string",
                        "description": "图表标题"
                    },
                    "has_legend": {
                        "type": "boolean",
                        "default": True,
                        "description": "是否显示图例"
                    },
                    "legend_position": {
                        "type": "string",
                        "enum": ["right", "left", "top", "bottom", "corner"],
                        "description": "图例位置"
                    },
                    "has_data_labels": {
                        "type": "boolean",
                        "default": False,
                        "description": "是否显示数据标签"
                    },
                    "data_label_position": {
                        "type": "string",
                        "enum": ["center", "inside_end", "inside_base", "outside_end", "best_fit"],
                        "description": "数据标签位置"
                    }
                },
                "required": ["path", "slide", "chart_type", "categories", "series"]
            }
        ),

        Tool(
            name="ppt_chart_modify",
            description="修改幻灯片中的图表属性",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "chart_index": {
                        "type": "integer",
                        "description": "图表索引(从0开始)"
                    },
                    "title": {
                        "type": "string",
                        "description": "新的图表标题"
                    },
                    "has_legend": {
                        "type": "boolean",
                        "description": "是否显示图例"
                    },
                    "legend_position": {
                        "type": "string",
                        "enum": ["right", "left", "top", "bottom", "corner"],
                        "description": "图例位置"
                    },
                    "has_data_labels": {
                        "type": "boolean",
                        "description": "是否显示数据标签"
                    },
                    "data_label_position": {
                        "type": "string",
                        "enum": ["center", "inside_end", "inside_base", "outside_end", "best_fit"],
                        "description": "数据标签位置"
                    },
                    "style": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 48,
                        "description": "图表样式编号(1-48)"
                    }
                },
                "required": ["path", "slide", "chart_index"]
            }
        ),

        Tool(
            name="chart_get_data",
            description="获取幻灯片中图表的数据",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "chart_index": {
                        "type": "integer",
                        "description": "图表索引(从0开始)"
                    }
                },
                "required": ["path", "slide", "chart_index"]
            }
        ),

        Tool(
            name="chart_list",
            description="列出幻灯片中的所有图表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    }
                },
                "required": ["path", "slide"]
            }
        ),

        Tool(
            name="chart_delete",
            description="删除幻灯片中的图表",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引(从1开始)"
                    },
                    "chart_index": {
                        "type": "integer",
                        "description": "图表索引(从0开始)"
                    }
                },
                "required": ["path", "slide", "chart_index"]
            }
        ),

        # ============================================================
        # AI图片生成
        # ============================================================
        Tool(
            name="image_gen_status",
            description="获取图片生成模块状态（配置、API可用性）",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        Tool(
            name="image_generate",
            description="使用AI生成图片，根据文字描述生成图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "prompt": {
                        "type": "string",
                        "description": "图片描述提示词，描述你想要生成的图片内容"
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "输出目录（可选，默认 generated_images）"
                    },
                    "filename": {
                        "type": "string",
                        "description": "文件名（可选，自动生成）"
                    },
                    "timeout": {
                        "type": "integer",
                        "description": "超时时间（秒），默认120"
                    },
                    "model": {
                        "type": "string",
                        "description": "模型名称（可选）"
                    },
                    "api_url": {
                        "type": "string",
                        "description": "API地址（可选）"
                    }
                },
                "required": ["prompt"]
            }
        ),

        Tool(
            name="image_generate_for_ppt",
            description="生成AI图片并直接插入到PPT幻灯片中",
            inputSchema={
                "type": "object",
                "properties": {
                    "ppt_path": {
                        "type": "string",
                        "description": "PPT文件路径"
                    },
                    "slide": {
                        "type": "integer",
                        "description": "幻灯片索引（从1开始）"
                    },
                    "prompt": {
                        "type": "string",
                        "description": "图片描述提示词"
                    },
                    "left": {
                        "type": "string",
                        "default": "1in",
                        "description": "图片左边距，如 '1in', '2.54cm'"
                    },
                    "top": {
                        "type": "string",
                        "default": "1in",
                        "description": "图片上边距"
                    },
                    "width": {
                        "type": "string",
                        "description": "图片宽度（可选，保持比例）"
                    },
                    "height": {
                        "type": "string",
                        "description": "图片高度（可选）"
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "图片保存目录（可选）"
                    },
                    "timeout": {
                        "type": "integer",
                        "description": "超时时间（秒）"
                    },
                    "model": {
                        "type": "string",
                        "description": "模型名称（可选）"
                    }
                },
                "required": ["ppt_path", "slide", "prompt"]
            }
        ),

        # ============================================================
        # HTML转PPTX
        # ============================================================
        Tool(
            name="html_to_pptx_convert",
            description="将HTML网页(div+p标签,内联样式)转换为PPTX幻灯片",
            inputSchema={
                "type": "object",
                "properties": {
                    "html_source": {
                        "type": "string",
                        "description": "HTML文件路径或HTML内容字符串"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出PPTX文件路径"
                    },
                    "base_path": {
                        "type": "string",
                        "description": "基础路径(用于解析相对图片路径)"
                    }
                },
                "required": ["html_source", "output_path"]
            }
        ),

        Tool(
            name="html_to_pptx_status",
            description="获取HTML转PPTX模块状态(支持的CSS属性、元素等)",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        Tool(
            name="html_to_pptx_convert_multi",
            description="将多个HTML网页(div+p标签,内联样式)转换为一个多页PPTX幻灯片",
            inputSchema={
                "type": "object",
                "properties": {
                    "html_sources": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "HTML文件路径或HTML内容字符串的列表，每个元素生成一张幻灯片"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出PPTX文件路径"
                    }
                },
                "required": ["html_sources", "output_path"]
            }
        ),
    ]

    # ============================================================
    # Auto-generate schemas for registered tools without explicit schemas
    # Ensures all callable tools are also discoverable via list_tools
    # ============================================================
    try:
        from docuflow_mcp.core.registry import get_all_registered_tools, get_tool_info
        schema_names = {t.name for t in tools}
        registered_names = set(get_all_registered_tools())

        for name in sorted(registered_names - schema_names):
            info = get_tool_info(name)
            properties = {}
            for p in info.get("required_params", []):
                properties[p] = {"type": "string", "description": p}
            for p in info.get("optional_params", []):
                properties[p] = {"type": "string", "description": f"(optional) {p}"}
            tools.append(Tool(
                name=name,
                description=f"{name} (auto-schema)",
                inputSchema={
                    "type": "object",
                    "properties": properties,
                    "required": info.get("required_params", [])
                }
            ))
    except Exception:
        pass  # Fail silently — static list is still returned

    return tools
