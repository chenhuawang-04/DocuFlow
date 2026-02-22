"""
DocuFlow MCP 新功能测试脚本
"""

import sys
import os

# 添加路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from docuflow_mcp.extensions.templates import TemplateManager
from docuflow_mcp.extensions.styles import StyleManager

def test_template_list():
    """测试1: 列出预设模板"""
    print("=" * 50)
    print("测试1: 列出预设模板")
    print("=" * 50)
    result = TemplateManager.list_presets()
    print(f"成功: {result['success']}")
    print(f"模板数量: {result['count']}")
    for preset in result['presets']:
        print(f"  - {preset['id']}: {preset['name']} - {preset['description']}")
    print()

def test_create_from_preset():
    """测试2: 从预设创建文档"""
    print("=" * 50)
    print("测试2: 从预设创建MBA论文文档")
    print("=" * 50)
    output_path = "test_mba_thesis.docx"
    result = TemplateManager.create_from_preset(
        preset_name="mba_thesis",
        output_path=output_path,
        title="测试MBA论文"
    )
    print(f"成功: {result['success']}")
    print(f"消息: {result['message']}")
    print(f"文件路径: {result.get('path', 'N/A')}")

    # 检查文件是否存在
    if os.path.exists(output_path):
        print(f"[OK] 文件已创建: {output_path}")
        print(f"  文件大小: {os.path.getsize(output_path)} 字节")
    else:
        print(f"[ERROR] 文件未创建")
    print()

def test_style_export():
    """测试3: 导出样式"""
    print("=" * 50)
    print("测试3: 导出文档样式")
    print("=" * 50)

    # 先确保有文档存在
    if not os.path.exists("test_mba_thesis.docx"):
        print("跳过: 需要先运行测试2创建文档")
        return

    result = StyleManager.export_styles(
        path="test_mba_thesis.docx"
    )
    print(f"成功: {result['success']}")
    print(f"样式数量: {result.get('style_count', 0)}")
    if 'styles' in result:
        print("前5个样式:")
        for i, (name, _) in enumerate(list(result['styles'].items())[:5]):
            print(f"  {i+1}. {name}")
    print()

def test_business_report():
    """测试4: 创建商业报告"""
    print("=" * 50)
    print("测试4: 创建商业报告文档")
    print("=" * 50)
    output_path = "test_business_report.docx"
    result = TemplateManager.create_from_preset(
        preset_name="business_report",
        output_path=output_path,
        title="2024年度商业报告"
    )
    print(f"成功: {result['success']}")
    print(f"消息: {result['message']}")

    if os.path.exists(output_path):
        print(f"[OK] 文件已创建: {output_path}")
    print()

if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("DocuFlow MCP 新功能测试")
    print("=" * 50 + "\n")

    try:
        test_template_list()
        test_create_from_preset()
        test_style_export()
        test_business_report()

        print("=" * 50)
        print("[SUCCESS] 所有测试完成！")
        print("=" * 50)
        print("\n生成的测试文件:")
        print("  - test_mba_thesis.docx")
        print("  - test_business_report.docx")
        print("\n请打开这些文件检查格式是否正确应用。")

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
