"""
DocuFlow PDF - 功能测试脚本

测试PDF文档处理功能
"""
import os
import sys
import tempfile
import shutil
import io

# 设置stdout为utf-8编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from docuflow_mcp.extensions.pdf import PDFOperations
from docuflow_mcp.core.registry import get_all_registered_tools


def test_tool_registration():
    """测试工具是否正确注册"""
    print("=" * 60)
    print("测试1: PDF工具注册检查")
    print("=" * 60)

    tools = get_all_registered_tools()
    pdf_tools = [
        # 信息与提取 (5个)
        'pdf_info', 'pdf_extract_text', 'pdf_extract_tables',
        'pdf_extract_images', 'pdf_get_outline',
        # 文件操作 (6个)
        'pdf_merge', 'pdf_split', 'pdf_extract_pages',
        'pdf_rotate', 'pdf_delete_pages', 'pdf_add_watermark',
        # 转换与集成 (4个)
        'pdf_tables_to_word', 'pdf_tables_to_excel',
        'pdf_to_text', 'pdf_status'
    ]

    registered = []
    missing = []

    for tool in pdf_tools:
        if tool in tools:
            registered.append(tool)
        else:
            missing.append(tool)

    print(f"已注册PDF工具: {len(registered)}/{len(pdf_tools)}")
    if missing:
        print(f"未注册工具: {missing}")
        return False

    print(f"✓ 所有15个PDF工具已注册")
    print(f"  当前总工具数: {len(tools)}")
    return True


def test_pdf_status():
    """测试PDF状态检查"""
    print("\n" + "=" * 60)
    print("测试2: PDF模块状态检查")
    print("=" * 60)

    result = PDFOperations.get_status()

    print(f"pdfplumber可用: {result.get('pdfplumber_available', False)}")
    print(f"pypdf可用: {result.get('pypdf_available', False)}")

    if result.get('versions'):
        print("\n版本信息:")
        for lib, ver in result['versions'].items():
            print(f"  - {lib}: {ver}")

    if result.get('features'):
        print("\n支持的功能:")
        for feature in result['features']:
            print(f"  - {feature}")

    return result.get('success', False)


def create_test_pdf(temp_dir: str) -> str:
    """创建测试用PDF文件（使用pypdf）"""
    try:
        from pypdf import PdfWriter
        from pypdf.generic import (
            DictionaryObject, NameObject, TextStringObject,
            ArrayObject, NumberObject
        )

        pdf_path = os.path.join(temp_dir, "test.pdf")

        writer = PdfWriter()

        # 创建简单的PDF页面
        for i in range(3):
            page = writer.add_blank_page(width=612, height=792)

        writer.add_metadata({
            "/Title": "测试PDF文档",
            "/Author": "DocuFlow Test",
            "/Subject": "PDF功能测试"
        })

        with open(pdf_path, 'wb') as f:
            writer.write(f)

        return pdf_path
    except Exception as e:
        print(f"  (创建测试PDF失败: {e})")
        return None


def test_pdf_info():
    """测试获取PDF信息"""
    print("\n" + "=" * 60)
    print("测试3: 获取PDF信息")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        # 获取信息
        result = PDFOperations.get_info(pdf_path)

        if result.get('success'):
            print(f"✓ 获取PDF信息成功")
            print(f"  页数: {result.get('pages')}")
            print(f"  文件大小: {result.get('file_size_mb')} MB")
            if result.get('metadata'):
                print(f"  元数据: {result.get('metadata')}")
            return True
        else:
            print(f"✗ 获取信息失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_extract_text():
    """测试文本提取"""
    print("\n" + "=" * 60)
    print("测试4: 文本提取")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        # 提取文本
        result = PDFOperations.extract_text(pdf_path)

        if result.get('success'):
            print(f"✓ 文本提取成功")
            print(f"  提取页数: {result.get('page_count')}")
            print(f"  总字符数: {result.get('total_chars')}")
            return True
        else:
            print(f"✗ 文本提取失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_extract_tables():
    """测试表格提取"""
    print("\n" + "=" * 60)
    print("测试5: 表格提取")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        # 提取表格
        result = PDFOperations.extract_tables(pdf_path)

        if result.get('success'):
            print(f"✓ 表格提取成功")
            print(f"  找到表格数: {result.get('table_count')}")
            return True
        else:
            print(f"✗ 表格提取失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_get_outline():
    """测试获取大纲"""
    print("\n" + "=" * 60)
    print("测试6: 获取PDF大纲")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        # 获取大纲
        result = PDFOperations.get_outline(pdf_path)

        if result.get('success'):
            print(f"✓ 获取大纲成功")
            print(f"  书签数: {result.get('outline_count')}")
            return True
        else:
            print(f"✗ 获取大纲失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_merge():
    """测试PDF合并"""
    print("\n" + "=" * 60)
    print("测试7: PDF合并")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建多个测试PDF
        pdf1 = create_test_pdf(temp_dir)
        if pdf1:
            os.rename(pdf1, os.path.join(temp_dir, "doc1.pdf"))
            pdf1 = os.path.join(temp_dir, "doc1.pdf")

        pdf2 = create_test_pdf(temp_dir)
        if pdf2:
            os.rename(pdf2, os.path.join(temp_dir, "doc2.pdf"))
            pdf2 = os.path.join(temp_dir, "doc2.pdf")

        if not pdf1 or not pdf2:
            print("✗ 无法创建测试PDF")
            return False

        output_path = os.path.join(temp_dir, "merged.pdf")

        # 合并
        result = PDFOperations.merge(
            paths=[pdf1, pdf2],
            output_path=output_path,
            add_outline=True
        )

        if result.get('success'):
            print(f"✓ PDF合并成功")
            print(f"  合并文件数: {result.get('input_count')}")
            print(f"  总页数: {result.get('total_pages')}")
            return True
        else:
            print(f"✗ PDF合并失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_split():
    """测试PDF拆分"""
    print("\n" + "=" * 60)
    print("测试8: PDF拆分")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        output_dir = os.path.join(temp_dir, "split_output")

        # 拆分 - 每页一个文件
        result = PDFOperations.split(
            path=pdf_path,
            output_dir=output_dir,
            mode='single'
        )

        if result.get('success'):
            print(f"✓ PDF拆分成功")
            print(f"  生成文件数: {result.get('file_count')}")
            return True
        else:
            print(f"✗ PDF拆分失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_extract_pages():
    """测试页面提取"""
    print("\n" + "=" * 60)
    print("测试9: 页面提取")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        output_path = os.path.join(temp_dir, "extracted.pdf")

        # 提取第1和第3页
        result = PDFOperations.extract_pages(
            path=pdf_path,
            pages=[1, 3],
            output_path=output_path
        )

        if result.get('success'):
            print(f"✓ 页面提取成功")
            print(f"  提取页码: {result.get('extracted_pages')}")
            return True
        else:
            print(f"✗ 页面提取失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_rotate():
    """测试页面旋转"""
    print("\n" + "=" * 60)
    print("测试10: 页面旋转")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        output_path = os.path.join(temp_dir, "rotated.pdf")

        # 旋转90度
        result = PDFOperations.rotate(
            path=pdf_path,
            angle=90,
            pages=[1],
            output_path=output_path
        )

        if result.get('success'):
            print(f"✓ 页面旋转成功")
            print(f"  旋转页数: {result.get('rotated_pages')}")
            print(f"  旋转角度: {result.get('angle')}°")
            return True
        else:
            print(f"✗ 页面旋转失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_delete_pages():
    """测试页面删除"""
    print("\n" + "=" * 60)
    print("测试11: 页面删除")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        output_path = os.path.join(temp_dir, "deleted.pdf")

        # 删除第2页
        result = PDFOperations.delete_pages(
            path=pdf_path,
            pages=[2],
            output_path=output_path
        )

        if result.get('success'):
            print(f"✓ 页面删除成功")
            print(f"  删除页数: {len(result.get('deleted_pages', []))}")
            print(f"  剩余页数: {result.get('remaining_pages')}")
            return True
        else:
            print(f"✗ 页面删除失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_watermark():
    """测试添加水印"""
    print("\n" + "=" * 60)
    print("测试12: 添加水印")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        output_path = os.path.join(temp_dir, "watermarked.pdf")

        # 添加水印
        result = PDFOperations.add_watermark(
            path=pdf_path,
            watermark="CONFIDENTIAL",
            position="center",
            opacity=0.3,
            output_path=output_path
        )

        if result.get('success'):
            print(f"✓ 添加水印成功")
            print(f"  水印文字: {result.get('watermark')}")
            print(f"  水印页数: {result.get('watermarked_pages')}")
            return True
        else:
            print(f"✗ 添加水印失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_pdf_to_text():
    """测试PDF转文本"""
    print("\n" + "=" * 60)
    print("测试13: PDF转文本")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        output_path = os.path.join(temp_dir, "output.txt")

        # 转换
        result = PDFOperations.to_text(
            path=pdf_path,
            output_path=output_path
        )

        if result.get('success'):
            print(f"✓ PDF转文本成功")
            print(f"  字符数: {result.get('chars')}")
            if os.path.exists(output_path):
                print(f"  输出文件: {output_path}")
            return True
        else:
            print(f"✗ PDF转文本失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_to_editable():
    """测试PDF转可编辑文档"""
    print("\n" + "=" * 60)
    print("测试14: PDF转可编辑文档")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        docx_path = os.path.join(temp_dir, "editable.docx")

        # 转换为Word
        result = PDFOperations.to_editable(
            path=pdf_path,
            output_path=docx_path,
            format='docx'
        )

        if result.get('success'):
            print(f"✓ PDF转可编辑文档成功")
            print(f"  输出文件: {result.get('output_path')}")
            print(f"  页数: {result.get('pages')}")
            print(f"  段落数: {result.get('paragraphs')}")
            return True
        else:
            print(f"✗ PDF转可编辑文档失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_annotate_text():
    """测试添加文字注释"""
    print("\n" + "=" * 60)
    print("测试15: 添加文字注释")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        # 创建测试PDF
        pdf_path = create_test_pdf(temp_dir)
        if not pdf_path:
            print("✗ 无法创建测试PDF")
            return False

        output_path = os.path.join(temp_dir, "annotated.pdf")

        # 添加文字
        result = PDFOperations.annotate_text(
            path=pdf_path,
            text="Test Annotation",
            x=100,
            y=700,
            page=1,
            output_path=output_path,
            font_color='red'
        )

        if result.get('success'):
            print(f"✓ 添加文字注释成功")
            print(f"  文字: {result.get('text')}")
            print(f"  位置: {result.get('position')}")
            return True
        else:
            print(f"✗ 添加文字注释失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_tool_registration_edit():
    """测试PDF编辑工具是否正确注册"""
    print("\n" + "=" * 60)
    print("测试16: PDF编辑工具注册检查")
    print("=" * 60)

    tools = get_all_registered_tools()
    edit_tools = [
        'pdf_to_editable',
        'pdf_text_replace',
        'pdf_redact',
        'pdf_annotate_text'
    ]

    registered = []
    missing = []

    for tool in edit_tools:
        if tool in tools:
            registered.append(tool)
        else:
            missing.append(tool)

    print(f"已注册PDF编辑工具: {len(registered)}/{len(edit_tools)}")
    if missing:
        print(f"未注册工具: {missing}")
        return False

    print(f"✓ 所有4个PDF编辑工具已注册")
    print(f"  当前总工具数: {len(tools)}")
    return True


def main():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("DocuFlow PDF 功能测试")
    print("=" * 60)

    results = {
        # 基础测试
        "工具注册": test_tool_registration(),
        "状态检查": test_pdf_status(),
        # 信息与提取
        "获取信息": test_pdf_info(),
        "文本提取": test_extract_text(),
        "表格提取": test_extract_tables(),
        "获取大纲": test_get_outline(),
        # 文件操作
        "PDF合并": test_merge(),
        "PDF拆分": test_split(),
        "页面提取": test_extract_pages(),
        "页面旋转": test_rotate(),
        "页面删除": test_delete_pages(),
        "添加水印": test_watermark(),
        # 转换
        "PDF转文本": test_pdf_to_text(),
        # PDF编辑（新增）
        "PDF转可编辑": test_to_editable(),
        "添加文字注释": test_annotate_text(),
        "编辑工具注册": test_tool_registration_edit(),
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
        print("\n所有测试通过!")
    else:
        print(f"\n{total - passed} 个测试失败")

    # 显示安装提示
    status = PDFOperations.get_status()
    if not status.get('all_available'):
        print("\n" + "=" * 60)
        print("安装提示")
        print("=" * 60)
        if not status.get('pdfplumber_available'):
            print("  pip install pdfplumber  # PDF读取和表格提取")
        if not status.get('pypdf_available'):
            print("  pip install pypdf       # PDF文件操作")

    return passed == total


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
