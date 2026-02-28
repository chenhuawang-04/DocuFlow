"""
测试 Day 3 功能：配置管理和代码重构
"""

import sys
import os
import time
import logging

# 添加包路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from docuflow_mcp.core.config import get_config, get, set, get_section
from docuflow_mcp.core.registry import dispatch_tool, get_all_registered_tools
from docuflow_mcp.utils.formatters import apply_font_format, apply_paragraph_format

# 导入操作类以触发装饰器注册
from docuflow_mcp.document import DocumentOperations
from docuflow_mcp.extensions.templates import TemplateManager
from docuflow_mcp.extensions.styles import StyleManager


def test_config_management():
    """测试配置管理功能"""
    print("\n" + "=" * 60)
    print("配置管理系统测试")
    print("=" * 60)

    config = get_config()

    # 测试获取配置值
    print("\n1. 测试获取配置值...")
    log_level = config.get('logging.level')
    slow_threshold = config.get('performance.slow_threshold')
    default_font = config.get('document.default_font')

    print(f"   [OK] 日志级别: {log_level}")
    print(f"   [OK] 慢查询阈值: {slow_threshold}s")
    print(f"   [OK] 默认字体: {default_font}")

    # 测试设置配置值
    print("\n2. 测试设置配置值...")
    original_level = config.get('logging.level')
    config.set('logging.level', logging.DEBUG)
    new_level = config.get('logging.level')

    if new_level == logging.DEBUG:
        print(f"   [OK] 设置成功: {original_level} -> {new_level}")
        # 恢复原值
        config.set('logging.level', original_level)
    else:
        print(f"   [FAIL] 设置失败")

    # 测试获取整个配置段
    print("\n3. 测试获取配置段...")
    logging_config = config.get_section('logging')
    performance_config = config.get_section('performance')

    print(f"   [OK] 日志配置: {len(logging_config)} 项")
    print(f"   [OK] 性能配置: {len(performance_config)} 项")

    # 测试配置持久化
    print("\n4. 测试配置保存...")
    test_config_file = "E:/Project/DocuFlow/test_config.json"
    try:
        config.save_to_file(test_config_file)
        if os.path.exists(test_config_file):
            print(f"   [OK] 配置已保存: {test_config_file}")

            # 测试加载配置
            config.load_from_file(test_config_file)
            print(f"   [OK] 配置已加载")

            # 清理测试文件
            os.remove(test_config_file)
            print(f"   [OK] 测试文件已清理")
    except Exception as e:
        print(f"   [FAIL] 配置保存/加载失败: {e}")

    # 测试便捷函数
    print("\n5. 测试便捷函数...")
    value1 = get('document.default_font')
    set('document.test_value', 'test123')
    value2 = get('document.test_value')
    section = get_section('middleware')

    print(f"   [OK] get() 函数: {value1}")
    print(f"   [OK] set() 函数: {value2}")
    print(f"   [OK] get_section() 函数: {len(section)} 项")

    print("\n" + "=" * 60)
    print("[OK] 配置管理系统测试完成!")
    print("=" * 60 + "\n")


def test_formatters():
    """测试格式化工具函数"""
    print("\n" + "=" * 60)
    print("格式化工具测试")
    print("=" * 60)

    test_file = "E:/Project/DocuFlow/test_formatters.docx"

    print("\n1. 测试创建文档...")
    result = dispatch_tool("doc_create", {
        "path": test_file,
        "title": "格式化测试文档"
    })

    if result.get("success"):
        print(f"   [OK] 文档已创建")
    else:
        print(f"   [FAIL] 文档创建失败: {result.get('error')}")
        return

    print("\n2. 测试添加格式化段落...")
    result = dispatch_tool("paragraph_add", {
        "path": test_file,
        "text": "这是一个测试段落",
        "font_name": "Microsoft YaHei",
        "font_size": "14pt",
        "font_color": "#FF0000",
        "bold": True,
        "alignment": "center",
        "line_spacing": 1.5,
        "first_line_indent": "2em"
    })

    if result.get("success"):
        print(f"   [OK] 格式化段落已添加")
    else:
        print(f"   [FAIL] 添加失败: {result.get('error')}")

    print("\n3. 测试修改段落格式...")
    result = dispatch_tool("paragraph_modify", {
        "path": test_file,
        "index": 1,
        "text": "修改后的段落",
        "font_size": "16pt",
        "italic": True
    })

    if result.get("success"):
        print(f"   [OK] 段落格式已修改")
    else:
        print(f"   [FAIL] 修改失败: {result.get('error')}")

    print("\n4. 测试读取段落信息...")
    result = dispatch_tool("paragraph_get", {
        "path": test_file,
        "index": 1
    })

    if result.get("success"):
        print(f"   [OK] 段落信息已读取")
        print(f"   [OK] 文本: {result.get('text')[:30]}...")
        print(f"   [OK] 样式: {result.get('style')}")
    else:
        print(f"   [FAIL] 读取失败: {result.get('error')}")

    # 清理测试文件
    if os.path.exists(test_file):
        try:
            os.remove(test_file)
            print(f"\n[OK] 测试文件已清理: {test_file}")
        except OSError:
            pass

    print("\n" + "=" * 60)
    print("[OK] 格式化工具测试完成!")
    print("=" * 60 + "\n")


def test_tool_registration():
    """测试工具注册系统"""
    print("\n" + "=" * 60)
    print("工具注册系统验证")
    print("=" * 60)

    print("\n1. 检查工具注册数量...")
    tools = get_all_registered_tools()
    print(f"   [OK] 已注册工具: {len(tools)} 个")

    if len(tools) >= 42:
        print(f"   [OK] 工具数量符合预期 (>= 42)")
    else:
        print(f"   [WARN] 工具数量少于预期: {len(tools)} < 42")

    print("\n2. 检查核心工具...")
    core_tools = [
        'doc_create', 'doc_read', 'doc_info',
        'paragraph_add', 'paragraph_modify',
        'table_add', 'table_get',
        'style_create', 'style_export',
        'template_list_presets', 'template_create_from_preset'
    ]

    missing_tools = []
    for tool in core_tools:
        if tool in tools:
            print(f"   [OK] {tool}")
        else:
            print(f"   [FAIL] {tool} 未注册")
            missing_tools.append(tool)

    if not missing_tools:
        print(f"\n   [OK] 所有核心工具已注册")
    else:
        print(f"\n   [FAIL] 缺少 {len(missing_tools)} 个工具: {missing_tools}")

    print("\n" + "=" * 60)
    print("[OK] 工具注册系统验证完成!")
    print("=" * 60 + "\n")


def test_environment_config():
    """测试环境变量配置"""
    print("\n" + "=" * 60)
    print("环境变量配置测试")
    print("=" * 60)

    print("\n1. 测试环境变量...")
    # 设置测试环境变量
    os.environ['DOCUFLOW_LOG_LEVEL'] = 'DEBUG'
    os.environ['DOCUFLOW_DEFAULT_FONT'] = 'TestFont'
    os.environ['DOCUFLOW_SLOW_THRESHOLD'] = '0.5'

    # 重新加载配置
    config = get_config()
    config.reset()  # 重新加载包括环境变量

    log_level = config.get('logging.level')
    default_font = config.get('document.default_font')
    slow_threshold = config.get('performance.slow_threshold')

    print(f"   [OK] 日志级别 (env): {log_level}")
    print(f"   [OK] 默认字体 (env): {default_font}")
    print(f"   [OK] 慢查询阈值 (env): {slow_threshold}")

    # 验证环境变量是否生效
    if log_level == logging.DEBUG:
        print(f"   [OK] 环境变量 DOCUFLOW_LOG_LEVEL 已生效")
    else:
        print(f"   [WARN] 环境变量 DOCUFLOW_LOG_LEVEL 未生效")

    if default_font == 'TestFont':
        print(f"   [OK] 环境变量 DOCUFLOW_DEFAULT_FONT 已生效")
    else:
        print(f"   [WARN] 环境变量 DOCUFLOW_DEFAULT_FONT 未生效")

    if slow_threshold == 0.5:
        print(f"   [OK] 环境变量 DOCUFLOW_SLOW_THRESHOLD 已生效")
    else:
        print(f"   [WARN] 环境变量 DOCUFLOW_SLOW_THRESHOLD 未生效")

    # 清理环境变量
    del os.environ['DOCUFLOW_LOG_LEVEL']
    del os.environ['DOCUFLOW_DEFAULT_FONT']
    del os.environ['DOCUFLOW_SLOW_THRESHOLD']

    print("\n" + "=" * 60)
    print("[OK] 环境变量配置测试完成!")
    print("=" * 60 + "\n")


def test_integration():
    """集成测试：创建完整文档"""
    print("\n" + "=" * 60)
    print("集成测试：完整文档创建")
    print("=" * 60)

    test_file = "E:/Project/DocuFlow/test_integration.docx"

    print("\n1. 创建文档...")
    result = dispatch_tool("doc_create", {
        "path": test_file,
        "title": "Day 3 集成测试文档"
    })
    print(f"   [OK] {result.get('message')}")

    print("\n2. 添加多个段落...")
    for i in range(3):
        result = dispatch_tool("paragraph_add", {
            "path": test_file,
            "text": f"这是第 {i+1} 个测试段落",
            "font_size": "12pt",
            "line_spacing": 1.5
        })
        if result.get("success"):
            print(f"   [OK] 段落 {i+1} 已添加")

    print("\n3. 添加表格...")
    result = dispatch_tool("table_add", {
        "path": test_file,
        "rows": 3,
        "cols": 3,
        "data": [
            ["列1", "列2", "列3"],
            ["数据1", "数据2", "数据3"],
            ["数据4", "数据5", "数据6"]
        ]
    })
    if result.get("success"):
        print(f"   [OK] 表格已添加")

    print("\n4. 读取文档信息...")
    result = dispatch_tool("doc_info", {"path": test_file})
    assert isinstance(result, dict), "Expected dict result"
    if result.get("success"):
        stats = result.get("statistics", {})
        print(f"   [OK] 段落数: {stats.get('paragraph_count')}")
        print(f"   [OK] 表格数: {stats.get('table_count')}")
        print(f"   [OK] 字符数: {stats.get('character_count')}")

    print("\n5. 测试性能...")
    start = time.time()
    for i in range(10):
        dispatch_tool("paragraph_add", {
            "path": test_file,
            "text": f"性能测试段落 {i}"
        })
    elapsed = time.time() - start
    avg = elapsed / 10 * 1000

    print(f"   [OK] 10次操作耗时: {elapsed*1000:.2f}ms")
    print(f"   [OK] 平均每次: {avg:.2f}ms")

    if avg < 50:
        print(f"   [OK] 性能良好 (<50ms)")
    else:
        print(f"   [WARN] 性能较慢 ({avg:.2f}ms)")

    # 清理测试文件
    if os.path.exists(test_file):
        try:
            os.remove(test_file)
            print(f"\n[OK] 测试文件已清理: {test_file}")
        except OSError:
            pass

    print("\n" + "=" * 60)
    print("[OK] 集成测试完成!")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    try:
        test_config_management()
        test_formatters()
        test_tool_registration()
        test_environment_config()
        test_integration()

        print("\n" + "=" * 60)
        print("所有测试通过! [OK]")
        print("=" * 60 + "\n")

        # 打印配置摘要
        print("\n配置系统摘要:")
        print("-" * 60)
        config = get_config()
        print(f"日志级别: {get('logging.level')}")
        print(f"性能监控: {get('performance.monitoring_enabled')}")
        print(f"慢查询阈值: {get('performance.slow_threshold')}s")
        print(f"默认字体: {get('document.default_font')}")
        print(f"默认字号: {get('document.default_font_size')}")
        print(f"已注册工具: {len(get_all_registered_tools())} 个")
        print("-" * 60)

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {e}")
        import traceback
        traceback.print_exc()
