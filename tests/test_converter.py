"""
DocuFlow Converter - 功能测试脚本

测试文档格式转换功能
"""
import os
import sys
import tempfile
import shutil

import pytest

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from docuflow_mcp.extensions.converter import ConverterOperations
from docuflow_mcp.core.registry import get_all_registered_tools

requires_pandoc = pytest.mark.skipif(
    shutil.which("pandoc") is None,
    reason="pandoc not installed",
)


def test_tool_registration():
    """测试工具是否正确注册"""
    print("=" * 60)
    print("测试1: 工具注册检查")
    print("=" * 60)

    tools = get_all_registered_tools()
    converter_tools = ['convert', 'convert_batch', 'convert_formats', 'convert_with_template']

    registered = []
    missing = []

    for tool in converter_tools:
        if tool in tools:
            registered.append(tool)
        else:
            missing.append(tool)

    print(f"已注册转换工具: {registered}")
    if missing:
        print(f"未注册工具: {missing}")
        assert False, "未注册工具: ..."

    print(f"✓ 所有4个转换工具已注册")
    print(f"  当前总工具数: {len(tools)}")

def test_get_formats():
    """测试获取支持格式"""
    print("\n" + "=" * 60)
    print("测试2: 获取支持格式")
    print("=" * 60)

    result = ConverterOperations.get_formats()
    assert isinstance(result, dict), "Expected dict result"

    print(f"Pandoc可用: {result.get('pandoc_available', False)}")
    print(f"输入格式数: {result.get('total_input', 0)}")
    print(f"输出格式数: {result.get('total_output', 0)}")

    if result.get('popular_conversions'):
        print("\n常用转换:")
        for conv in result['popular_conversions'][:5]:
            print(f"  - {conv['from']} -> {conv['to']}: {conv['desc']}")

    assert result.get('success', False), "获取格式列表失败"


@requires_pandoc
def test_markdown_to_html():
    """测试Markdown转HTML"""
    print("\n" + "=" * 60)
    print("测试3: Markdown -> HTML 转换")
    print("=" * 60)

    # 创建临时目录
    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试Markdown文件
        md_path = os.path.join(temp_dir, "test.md")
        html_path = os.path.join(temp_dir, "test.html")

        md_content = """# 测试文档

## 章节一

这是一段测试文本。

- 列表项1
- 列表项2
- 列表项3

## 章节二

更多内容在这里。
"""
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(md_content)

        print(f"创建测试文件: {md_path}")

        # 执行转换
        result = ConverterOperations.convert(
            source=md_path,
            target=html_path
        )

        print(f"转换结果: {result.get('success', False)}")

        if result.get('success'):
            # 验证输出文件
            if os.path.exists(html_path):
                with open(html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                print(f"HTML文件大小: {len(html_content)} 字符")
                print(f"包含<h1>: {'<h1>' in html_content or '<h1 ' in html_content}")
                print(f"✓ 转换成功!")
            else:
                print("✗ 输出文件不存在")
                assert False, "输出文件不存在"
        else:
            print(f"✗ 转换失败: {result.get('error', '未知错误')}")
            assert False, "转换失败: {result.get("

    finally:
        # 清理临时目录
        shutil.rmtree(temp_dir, ignore_errors=True)


@requires_pandoc
def test_markdown_to_docx():
    """测试Markdown转Word"""
    print("\n" + "=" * 60)
    print("测试4: Markdown -> Word 转换")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        md_path = os.path.join(temp_dir, "test.md")
        docx_path = os.path.join(temp_dir, "test.docx")

        md_content = """# 测试Word文档

## 第一章

这是中文内容测试。

1. 有序列表1
2. 有序列表2

## 第二章

**加粗文本** 和 *斜体文本*。
"""
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(md_content)

        print(f"创建测试文件: {md_path}")

        result = ConverterOperations.convert(
            source=md_path,
            target=docx_path
        )

        print(f"转换结果: {result.get('success', False)}")

        if result.get('success'):
            if os.path.exists(docx_path):
                file_size = os.path.getsize(docx_path)
                print(f"Word文件大小: {file_size} 字节")
                print(f"✓ 转换成功!")
            else:
                print("✗ 输出文件不存在")
                assert False, "输出文件不存在"
        else:
            print(f"✗ 转换失败: {result.get('error', '未知错误')}")
            assert False, "转换失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@requires_pandoc
def test_batch_conversion():
    """测试批量转换"""
    print("\n" + "=" * 60)
    print("测试5: 批量转换")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    output_dir = os.path.join(temp_dir, "output")

    try:
        # 创建多个Markdown文件
        sources = []
        for i in range(3):
            md_path = os.path.join(temp_dir, f"doc{i+1}.md")
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(f"# 文档 {i+1}\n\n这是第{i+1}个测试文档的内容。\n")
            sources.append(md_path)

        print(f"创建了 {len(sources)} 个测试文件")

        # 批量转换为HTML
        result = ConverterOperations.convert_batch(
            sources=sources,
            target_format='html',
            output_dir=output_dir
        )

        print(f"批量转换结果:")
        print(f"  - 总数: {result.get('total', 0)}")
        print(f"  - 成功: {result.get('converted', 0)}")
        print(f"  - 失败: {result.get('failed', 0)}")

        if result.get('success'):
            print("✓ 批量转换成功!")
        else:
            print("✗ 部分转换失败")
            for r in result.get('results', []):
                if not r.get('success'):
                    print(f"  失败: {r.get('error', '未知')}")
            assert False, "失败: {r.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_format_detection():
    """测试格式自动检测"""
    print("\n" + "=" * 60)
    print("测试6: 格式自动检测")
    print("=" * 60)

    test_cases = [
        ("test.md", "markdown"),
        ("test.markdown", "markdown"),
        ("document.docx", "docx"),
        ("report.pdf", "pdf"),
        ("page.html", "html"),
        ("paper.tex", "latex"),
        ("notes.txt", "plain"),
    ]

    all_pass = True
    for filename, expected in test_cases:
        detected = ConverterOperations._detect_format(filename)
        normalized = ConverterOperations._normalize_format(detected)
        status = "✓" if normalized == expected else "✗"
        print(f"  {status} {filename} -> {detected} (预期: {expected})")
        if normalized != expected:
            all_pass = False

    assert all_pass, "部分格式检测不匹配"


def main():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("DocuFlow Converter 功能测试")
    print("=" * 60)

    results = {
        "工具注册": test_tool_registration(),
        "获取格式": test_get_formats(),
        "MD->HTML": test_markdown_to_html(),
        "MD->Word": test_markdown_to_docx(),
        "批量转换": test_batch_conversion(),
        "格式检测": test_format_detection(),
    }

    print("\n" + "=" * 60)
    print("测试汇总")
    print("=" * 60)

    passed = sum(1 for v in results.values() if v)
    total = len(results)

    for name, result in results.items():
        status = "✓ 通过" if result else "✗ 失败"
        print(f"  {name}: {status}")

    print("-" * 60)
    print(f"总计: {passed}/{total} 通过")

    if passed == total:
        print("\n🎉 所有测试通过!")
    else:
        print(f"\n⚠️ {total - passed} 个测试失败")

    return passed == total


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
