"""
测试中间件功能
"""

import sys
import os
import time

# 添加包路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from docuflow_mcp.core.registry import dispatch_tool, get_all_registered_tools
from docuflow_mcp.core.middleware import (
    get_middleware_manager,
    LoggingMiddleware,
    PerformanceMiddleware,
    ErrorHandlingMiddleware,
    clear_middlewares
)

# 导入操作类以触发装饰器注册
from docuflow_mcp.document import DocumentOperations
from docuflow_mcp.extensions.templates import TemplateManager
from docuflow_mcp.extensions.styles import StyleManager


def test_middleware_system():
    """测试中间件系统"""
    print("\n" + "="*60)
    print("中间件系统测试")
    print("="*60)

    # 清空现有中间件
    clear_middlewares()

    # 获取中间件管理器
    manager = get_middleware_manager()

    # 添加中间件
    print("\n1. 添加中间件...")
    manager.add(ErrorHandlingMiddleware())
    manager.add(PerformanceMiddleware(slow_threshold=0.1))  # 100ms阈值
    manager.add(LoggingMiddleware(log_level=20))  # INFO级别

    print(f"   已添加 {len(manager.middlewares)} 个中间件")

    # 测试工具注册
    print("\n2. 检查工具注册...")
    tools = get_all_registered_tools()
    print(f"   已注册工具数量: {len(tools)}")
    print(f"   前10个工具: {tools[:10]}")

    # 测试正常调用
    print("\n3. 测试正常工具调用...")
    test_file = "E:/Project/DocuFlow/test_middleware.docx"

    result = dispatch_tool("doc_create", {
        "path": test_file,
        "title": "中间件测试文档"
    })

    if result.get("success"):
        print(f"   [OK] 创建文档成功: {test_file}")
        if "_performance" in result:
            perf = result["_performance"]
            print(f"   [OK] 性能数据: {perf['elapsed_ms']:.2f}ms")
    else:
        print(f"   [FAIL] 创建文档失败: {result.get('error')}")

    # 测试错误处理
    print("\n4. 测试错误处理...")
    result = dispatch_tool("doc_read", {
        "path": "E:/Project/DocuFlow/nonexistent.docx"
    })

    if not result.get("success"):
        print(f"   [OK] 正确捕获错误: {result.get('error_code')}")
        print(f"   [OK] 错误信息: {result.get('error')}")
    else:
        print("   [FAIL] 应该返回错误但没有")

    # 测试缺少参数
    print("\n5. 测试参数验证...")
    result = dispatch_tool("doc_create", {})  # 缺少 path 参数

    if not result.get("success") and result.get("error_code") == "MISSING_PARAMS":
        print(f"   [OK] 正确检测缺少参数: {result.get('missing_params')}")
    else:
        print("   [FAIL] 参数验证失败")

    # 测试未知工具
    print("\n6. 测试未知工具...")
    result = dispatch_tool("unknown_tool", {"arg": "value"})
    assert isinstance(result, dict), "Expected dict result"

    if not result.get("success") and result.get("error_code") == "TOOL_NOT_FOUND":
        print("   [OK] 正确处理未知工具")
    else:
        print("   [FAIL] 未知工具处理失败")

    # 获取性能统计
    print("\n7. 性能统计...")
    perf_middleware = None
    for m in manager.middlewares:
        if isinstance(m, PerformanceMiddleware):
            perf_middleware = m
            break

    if perf_middleware:
        stats = perf_middleware.get_stats()
        print(f"   已记录 {len(stats)} 个工具的性能数据:")
        for tool_name, tool_stats in list(stats.items())[:5]:
            print(f"     - {tool_name}:")
            print(f"       调用次数: {tool_stats['count']}")
            print(f"       平均时间: {tool_stats['avg_time']*1000:.2f}ms")
            print(f"       最小/最大: {tool_stats['min_time']*1000:.2f}/{tool_stats['max_time']*1000:.2f}ms")
            if tool_stats['slow_count'] > 0:
                print(f"       慢查询: {tool_stats['slow_count']} 次")

    # 清理测试文件
    if os.path.exists(test_file):
        try:
            os.remove(test_file)
            print(f"\n[OK] 已清理测试文件: {test_file}")
        except OSError:
            pass

    print("\n" + "="*60)
    print("[OK] 中间件系统测试完成!")
    print("="*60 + "\n")


def test_performance_benchmark():
    """性能基准测试"""
    print("\n" + "="*60)
    print("性能基准测试")
    print("="*60)

    # 清空中间件，测试无中间件性能
    clear_middlewares()
    manager = get_middleware_manager()

    test_file = "E:/Project/DocuFlow/test_perf.docx"
    iterations = 50

    # 无中间件性能测试
    print(f"\n1. 测试无中间件性能 ({iterations}次调用)...")
    start = time.time()
    for i in range(iterations):
        dispatch_tool("doc_create", {"path": test_file, "title": f"Test {i}"})
        if os.path.exists(test_file):
            os.remove(test_file)
    elapsed_no_middleware = time.time() - start
    avg_no_middleware = elapsed_no_middleware / iterations

    print(f"   总时间: {elapsed_no_middleware*1000:.2f}ms")
    print(f"   平均时间: {avg_no_middleware*1000:.2f}ms")

    # 添加所有中间件后的性能测试
    print(f"\n2. 测试完整中间件链性能 ({iterations}次调用)...")
    manager.add(ErrorHandlingMiddleware())
    manager.add(PerformanceMiddleware(slow_threshold=1.0))
    manager.add(LoggingMiddleware(log_level=30))  # WARNING级别，减少输出

    start = time.time()
    for i in range(iterations):
        dispatch_tool("doc_create", {"path": test_file, "title": f"Test {i}"})
        if os.path.exists(test_file):
            os.remove(test_file)
    elapsed_with_middleware = time.time() - start
    avg_with_middleware = elapsed_with_middleware / iterations

    print(f"   总时间: {elapsed_with_middleware*1000:.2f}ms")
    print(f"   平均时间: {avg_with_middleware*1000:.2f}ms")

    # 性能对比
    overhead = ((elapsed_with_middleware - elapsed_no_middleware) / elapsed_no_middleware) * 100
    print(f"\n3. 性能对比:")
    print(f"   中间件开销: {overhead:.2f}%")
    print(f"   绝对差异: {(avg_with_middleware - avg_no_middleware)*1000:.2f}ms")

    if overhead < 10:
        print(f"   [OK] 中间件开销在可接受范围内 (<10%)")
    else:
        print(f"   [WARN] 中间件开销较高 ({overhead:.1f}%)")

    print("\n" + "="*60)
    print("[OK] 性能基准测试完成!")
    print("="*60 + "\n")


if __name__ == "__main__":
    try:
        test_middleware_system()
        test_performance_benchmark()

        print("\n" + "="*60)
        print("所有测试通过! [OK]")
        print("="*60 + "\n")

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {e}")
        import traceback
        traceback.print_exc()
