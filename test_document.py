"""
DocuFlow MCP 功能测试脚本

测试文档创建和各种样式应用
"""

import sys
import os

# 添加源代码路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from docuflow_mcp.document import (
    DocumentOperations,
    ParagraphOperations,
    HeadingOperations,
    TableOperations,
    ListOperations,
    PageOperations,
    HeaderFooterOperations,
    SearchOperations,
    SpecialOperations,
    ExportOperations
)

def test_create_styled_document():
    """测试创建带有多种样式的文档"""

    doc_path = "E:/Project/DocuFlow/test_output/styled_document.docx"

    # 确保输出目录存在
    os.makedirs(os.path.dirname(doc_path), exist_ok=True)

    print("=" * 60)
    print("DocuFlow MCP 功能测试")
    print("=" * 60)

    # 1. 创建文档
    print("\n[1] 创建文档...")
    result = DocumentOperations.create(doc_path, title="DocuFlow 功能演示文档")
    print(f"    结果: {result}")

    # 2. 设置文档属性
    print("\n[2] 设置文档属性...")
    result = DocumentOperations.set_properties(doc_path, {
        "author": "DocuFlow MCP",
        "subject": "功能演示",
        "keywords": "MCP, Word, Python, AI"
    })
    print(f"    结果: {result}")

    # 3. 设置页面边距
    print("\n[3] 设置页边距...")
    result = PageOperations.set_margins(doc_path, top="2.54cm", bottom="2.54cm", left="3.17cm", right="3.17cm")
    print(f"    结果: {result}")

    # 4. 添加页眉
    print("\n[4] 设置页眉...")
    result = HeaderFooterOperations.set_header(doc_path, "DocuFlow MCP - Word 文档处理演示", alignment="center")
    print(f"    结果: {result}")

    # 5. 添加页脚页码
    print("\n[5] 添加页码...")
    result = HeaderFooterOperations.add_page_number(doc_path, position="footer", alignment="center")
    print(f"    结果: {result}")

    # 6. 添加一级标题
    print("\n[6] 添加一级标题...")
    result = HeadingOperations.add(doc_path, "第一章 文档样式演示", level=1)
    print(f"    结果: {result}")

    # 7. 添加普通段落
    print("\n[7] 添加普通段落...")
    result = ParagraphOperations.add(
        doc_path,
        "这是一个使用 DocuFlow MCP 创建的演示文档。DocuFlow MCP 是一个强大的 Word 文档处理工具，"
        "它通过 MCP (Model Context Protocol) 协议让 AI 能够直接操作 Word 文档。",
        font_name="微软雅黑",
        font_size="12pt",
        line_spacing=1.5,
        first_line_indent="24pt"
    )
    print(f"    结果: {result}")

    # 8. 添加二级标题
    print("\n[8] 添加二级标题...")
    result = HeadingOperations.add(doc_path, "1.1 文本格式演示", level=2)
    print(f"    结果: {result}")

    # 9. 添加加粗文本
    print("\n[9] 添加加粗红色文本...")
    result = ParagraphOperations.add(
        doc_path,
        "这是加粗的红色文本，用于强调重要内容。",
        font_name="微软雅黑",
        font_size="14pt",
        font_color="#FF0000",
        bold=True
    )
    print(f"    结果: {result}")

    # 10. 添加斜体文本
    print("\n[10] 添加斜体蓝色文本...")
    result = ParagraphOperations.add(
        doc_path,
        "这是斜体的蓝色文本，通常用于引用或注释。",
        font_name="微软雅黑",
        font_size="12pt",
        font_color="blue",
        italic=True
    )
    print(f"    结果: {result}")

    # 11. 添加居中文本
    print("\n[11] 添加居中文本...")
    result = ParagraphOperations.add(
        doc_path,
        "— 这是居中对齐的文本 —",
        font_name="微软雅黑",
        font_size="14pt",
        alignment="center",
        space_before="12pt",
        space_after="12pt"
    )
    print(f"    结果: {result}")

    # 12. 添加二级标题
    print("\n[12] 添加二级标题 - 列表演示...")
    result = HeadingOperations.add(doc_path, "1.2 列表演示", level=2)
    print(f"    结果: {result}")

    # 13. 添加无序列表
    print("\n[13] 添加无序列表...")
    result = ListOperations.add_bullet_list(doc_path, [
        "支持创建和读取文档",
        "支持段落、标题、表格操作",
        "支持图片插入",
        "支持页面设置和页眉页脚"
    ])
    print(f"    结果: {result}")

    # 14. 添加有序列表
    print("\n[14] 添加有序列表...")
    result = ListOperations.add_numbered_list(doc_path, [
        "第一步：安装依赖",
        "第二步：配置 MCP 服务器",
        "第三步：在 Claude Code 中使用"
    ])
    print(f"    结果: {result}")

    # 15. 添加二级标题
    print("\n[15] 添加二级标题 - 表格演示...")
    result = HeadingOperations.add(doc_path, "1.3 表格演示", level=2)
    print(f"    结果: {result}")

    # 16. 添加表格
    print("\n[16] 添加数据表格...")
    result = TableOperations.add(
        doc_path,
        rows=5,
        cols=4,
        data=[
            ["功能类别", "工具数量", "状态", "说明"],
            ["文档操作", "6", "✓ 完成", "创建、读取、合并等"],
            ["段落操作", "4", "✓ 完成", "添加、修改、删除"],
            ["表格操作", "9", "✓ 完成", "全面的表格处理"],
            ["导出功能", "2", "✓ 完成", "文本、Markdown"]
        ],
        style="Table Grid"
    )
    print(f"    结果: {result}")

    # 17. 设置表格标题行样式
    print("\n[17] 设置表格标题行样式...")
    for col in range(4):
        result = TableOperations.set_cell(
            doc_path,
            table_index=0,
            row=0,
            col=col,
            text=["功能类别", "工具数量", "状态", "说明"][col],
            bold=True,
            alignment="center",
            background_color="#4472C4"
        )
    print(f"    结果: 标题行样式已设置")

    # 18. 添加分页符
    print("\n[18] 添加分页符...")
    result = PageOperations.add_page_break(doc_path)
    print(f"    结果: {result}")

    # 19. 添加一级标题
    print("\n[19] 添加第二章标题...")
    result = HeadingOperations.add(doc_path, "第二章 高级功能", level=1)
    print(f"    结果: {result}")

    # 20. 添加超链接
    print("\n[20] 添加超链接...")
    result = SpecialOperations.add_hyperlink(
        doc_path,
        text="点击访问 GitHub 仓库",
        url="https://github.com"
    )
    print(f"    结果: {result}")

    # 21. 添加水平线
    print("\n[21] 添加水平分隔线...")
    result = SpecialOperations.add_horizontal_line(doc_path)
    print(f"    结果: {result}")

    # 22. 添加引用样式段落
    print("\n[22] 添加引用段落...")
    result = ParagraphOperations.add(
        doc_path,
        "\"工欲善其事，必先利其器。\" —— 《论语·卫灵公》",
        font_name="楷体",
        font_size="14pt",
        font_color="gray",
        italic=True,
        alignment="center",
        space_before="18pt",
        space_after="18pt"
    )
    print(f"    结果: {result}")

    # 23. 添加水平线
    result = SpecialOperations.add_horizontal_line(doc_path)

    # 24. 添加结束段落
    print("\n[24] 添加结束段落...")
    result = ParagraphOperations.add(
        doc_path,
        "本文档由 DocuFlow MCP 自动生成，演示了文档处理的多种功能。",
        font_name="微软雅黑",
        font_size="11pt",
        alignment="right",
        font_color="#666666"
    )
    print(f"    结果: {result}")

    # 25. 获取文档信息
    print("\n[25] 获取文档信息...")
    result = DocumentOperations.get_info(doc_path)
    print(f"    段落数: {result['statistics']['paragraph_count']}")
    print(f"    表格数: {result['statistics']['table_count']}")
    print(f"    字符数: {result['statistics']['character_count']}")

    # 26. 获取文档大纲
    print("\n[26] 获取文档大纲...")
    result = HeadingOperations.get_outline(doc_path)
    print("    大纲结构:")
    for item in result['outline']:
        indent = "    " * item['level']
        print(f"      {indent}{item['text']}")

    # 27. 导出为 Markdown
    print("\n[27] 导出为 Markdown...")
    md_path = "E:/Project/DocuFlow/test_output/styled_document.md"
    result = ExportOperations.to_markdown(doc_path, md_path)
    print(f"    结果: {result}")

    print("\n" + "=" * 60)
    print("测试完成!")
    print(f"文档已保存到: {doc_path}")
    print(f"Markdown 已导出到: {md_path}")
    print("=" * 60)

    return doc_path


if __name__ == "__main__":
    test_create_styled_document()
