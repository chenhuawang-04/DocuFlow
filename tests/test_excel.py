"""
DocuFlow Excel - 功能测试脚本

测试Excel表格处理功能
"""
import os
import sys
import tempfile
import shutil

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from docuflow_mcp.extensions.excel import ExcelOperations
from docuflow_mcp.core.registry import get_all_registered_tools


def test_tool_registration():
    """测试工具是否正确注册"""
    print("=" * 60)
    print("测试1: Excel工具注册检查")
    print("=" * 60)

    tools = get_all_registered_tools()
    excel_tools = [
        # 基础工具 (20个)
        'excel_create', 'excel_read', 'excel_info', 'excel_save_as',
        'sheet_list', 'sheet_add', 'sheet_delete', 'sheet_rename', 'sheet_copy',
        'cell_read', 'cell_write', 'cell_format', 'cell_merge', 'cell_formula',
        'row_insert', 'row_delete', 'col_insert', 'col_delete',
        'excel_to_word', 'excel_status',
        # 高级功能 (12个)
        'formula_batch', 'formula_quick',  # 公式增强
        'data_sort', 'data_filter', 'data_validate', 'data_deduplicate', 'data_fill',  # 数据操作
        'stats_summary', 'conditional_format', 'named_range',  # 统计格式
        'chart_create', 'excel_chart_modify'  # 图表
    ]

    registered = []
    missing = []

    for tool in excel_tools:
        if tool in tools:
            registered.append(tool)
        else:
            missing.append(tool)

    print(f"已注册Excel工具: {len(registered)}/{len(excel_tools)}")
    if missing:
        print(f"未注册工具: {missing}")
        return False

    print(f"✓ 所有32个Excel工具已注册")
    print(f"  当前总工具数: {len(tools)}")
    return True


def test_excel_status():
    """测试Excel状态检查"""
    print("\n" + "=" * 60)
    print("测试2: Excel模块状态检查")
    print("=" * 60)

    result = ExcelOperations.get_status()
    assert isinstance(result, dict), "Expected dict result"

    print(f"openpyxl可用: {result.get('openpyxl_available', False)}")
    if result.get('version'):
        print(f"openpyxl版本: {result.get('version')}")

    if result.get('features'):
        print("\n支持的功能:")
        for feature in result['features']:
            print(f"  - {feature}")

    return result.get('success', False)


def test_create_and_read():
    """测试创建和读取Excel"""
    print("\n" + "=" * 60)
    print("测试3: 创建和读取Excel")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "test.xlsx")

    try:
        # 创建Excel
        result = ExcelOperations.create(
            path=xlsx_path,
            sheets=["数据", "汇总"],
            title="测试文档"
        )

        if not result.get('success'):
            print(f"✗ 创建失败: {result.get('error')}")
            return False

        print(f"✓ 创建成功: {xlsx_path}")
        print(f"  工作表: {result.get('sheets')}")

        # 读取信息
        info = ExcelOperations.get_info(xlsx_path)
        if info.get('success'):
            print(f"✓ 读取信息成功")
            print(f"  工作表数: {info.get('sheet_count')}")
        else:
            print(f"✗ 读取信息失败: {info.get('error')}")
            return False

        return True

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_cell_operations():
    """测试单元格读写"""
    print("\n" + "=" * 60)
    print("测试4: 单元格读写操作")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "cells.xlsx")

    try:
        # 创建文件
        ExcelOperations.create(path=xlsx_path)

        # 写入单个单元格
        result = ExcelOperations.write_cell(
            path=xlsx_path,
            cell="A1",
            value="Hello Excel"
        )
        if not result.get('success'):
            print(f"✗ 写入单元格失败: {result.get('error')}")
            return False
        print("✓ 写入单元格 A1 成功")

        # 写入数据区域
        data = [
            ["姓名", "部门", "销售额"],
            ["张三", "销售部", 50000],
            ["李四", "市场部", 45000],
            ["王五", "技术部", 60000]
        ]
        result = ExcelOperations.write_cell(
            path=xlsx_path,
            range="A3",
            data=data
        )
        if not result.get('success'):
            print(f"✗ 写入数据区域失败: {result.get('error')}")
            return False
        print(f"✓ 写入数据区域成功: {len(data)} 行")

        # 读取单元格
        result = ExcelOperations.read_cell(
            path=xlsx_path,
            cell="A1"
        )
        if result.get('success') and result.get('value') == "Hello Excel":
            print(f"✓ 读取单元格成功: {result.get('value')}")
        else:
            print(f"✗ 读取单元格失败")
            return False

        # 读取区域
        result = ExcelOperations.read_cell(
            path=xlsx_path,
            range="A3:C6"
        )
        if result.get('success'):
            print(f"✓ 读取区域成功: {result.get('rows')} 行")
        else:
            print(f"✗ 读取区域失败")
            return False

        return True

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_formatting():
    """测试格式化"""
    print("\n" + "=" * 60)
    print("测试5: 单元格格式化")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "format.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[["标题行", "数据1", "数据2", "数据3"]]
        )

        # 格式化标题行
        result = ExcelOperations.format_cell(
            path=xlsx_path,
            range="A1:D1",
            bold=True,
            bg_color="4472C4",
            font_color="FFFFFF",
            alignment="center"
        )

        if result.get('success'):
            print("✓ 格式化标题行成功")
        else:
            print(f"✗ 格式化失败: {result.get('error')}")
            return False

        # 合并单元格
        result = ExcelOperations.merge_cell(
            path=xlsx_path,
            range="A3:D3"
        )
        if result.get('success'):
            print("✓ 合并单元格成功")
        else:
            print(f"✗ 合并单元格失败: {result.get('error')}")
            return False

        return True

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_formula():
    """测试公式"""
    print("\n" + "=" * 60)
    print("测试6: 公式设置")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "formula.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["Q1", "Q2", "Q3", "Q4", "合计"],
                [100, 200, 300, 400, None]
            ]
        )

        # 设置公式
        result = ExcelOperations.set_formula(
            path=xlsx_path,
            cell="E2",
            formula="=SUM(A2:D2)"
        )

        if result.get('success'):
            print(f"✓ 公式设置成功: {result.get('formula')}")
        else:
            print(f"✗ 公式设置失败: {result.get('error')}")
            return False

        return True

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_sheet_operations():
    """测试工作表操作"""
    print("\n" + "=" * 60)
    print("测试7: 工作表操作")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "sheets.xlsx")

    try:
        # 创建文件
        ExcelOperations.create(path=xlsx_path)

        # 列出工作表
        result = ExcelOperations.list_sheets(xlsx_path)
        assert isinstance(result, dict), "Expected dict result"
        print(f"初始工作表: {result.get('sheets')}")

        # 添加工作表
        result = ExcelOperations.add_sheet(xlsx_path, "新工作表")
        assert isinstance(result, dict), "Expected dict result"
        if result.get('success'):
            print("✓ 添加工作表成功")
        else:
            print(f"✗ 添加失败: {result.get('error')}")
            return False

        # 重命名工作表
        result = ExcelOperations.rename_sheet(xlsx_path, "新工作表", "报表")
        assert isinstance(result, dict), "Expected dict result"
        if result.get('success'):
            print("✓ 重命名工作表成功")
        else:
            print(f"✗ 重命名失败: {result.get('error')}")
            return False

        # 复制工作表
        result = ExcelOperations.copy_sheet(xlsx_path, "报表", "报表副本")
        assert isinstance(result, dict), "Expected dict result"
        if result.get('success'):
            print("✓ 复制工作表成功")
        else:
            print(f"✗ 复制失败: {result.get('error')}")
            return False

        # 列出最终结果
        result = ExcelOperations.list_sheets(xlsx_path)
        assert isinstance(result, dict), "Expected dict result"
        print(f"最终工作表: {result.get('sheets')}")

        return True

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_row_col_operations():
    """测试行列操作"""
    print("\n" + "=" * 60)
    print("测试8: 行列操作")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "rowcol.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["A", "B", "C"],
                ["1", "2", "3"],
                ["4", "5", "6"]
            ]
        )

        # 插入行
        result = ExcelOperations.insert_row(xlsx_path, row=2)
        assert isinstance(result, dict), "Expected dict result"
        if result.get('success'):
            print("✓ 插入行成功")
        else:
            print(f"✗ 插入行失败: {result.get('error')}")
            return False

        # 插入列
        result = ExcelOperations.insert_col(xlsx_path, col="B")
        assert isinstance(result, dict), "Expected dict result"
        if result.get('success'):
            print("✓ 插入列成功")
        else:
            print(f"✗ 插入列失败: {result.get('error')}")
            return False

        # 删除行
        result = ExcelOperations.delete_row(xlsx_path, row=2)
        assert isinstance(result, dict), "Expected dict result"
        if result.get('success'):
            print("✓ 删除行成功")
        else:
            print(f"✗ 删除行失败: {result.get('error')}")
            return False

        return True

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_excel_to_word():
    """测试Excel转Word"""
    print("\n" + "=" * 60)
    print("测试9: Excel表格插入Word")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "data.xlsx")
    docx_path = os.path.join(temp_dir, "output.docx")

    try:
        # 创建Excel并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["产品", "单价", "数量", "金额"],
                ["产品A", 100, 5, 500],
                ["产品B", 200, 3, 600],
                ["产品C", 150, 4, 600]
            ]
        )

        # 插入到Word
        result = ExcelOperations.to_word(
            excel_path=xlsx_path,
            word_path=docx_path,
            range="A1:D4",
            style="Table Grid"
        )

        if result.get('success'):
            print(f"✓ Excel表格插入Word成功")
            print(f"  表格大小: {result.get('rows')}x{result.get('cols')}")
            print(f"  输出文件: {docx_path}")
            return True
        else:
            print(f"✗ 插入失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_formula_batch():
    """测试批量公式"""
    print("\n" + "=" * 60)
    print("测试10: 批量公式")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "formula_batch.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["A", "B", "C", "D", "合计"],
                [10, 20, 30, 40, None],
                [15, 25, 35, 45, None],
                [20, 30, 40, 50, None]
            ]
        )

        # 批量设置公式
        result = ExcelOperations.formula_batch(
            path=xlsx_path,
            range="E2:E4",
            formula="=SUM(A{row}:D{row})"
        )

        if result.get('success'):
            print(f"✓ 批量公式设置成功: {result.get('cells_updated')} 个单元格")
            return True
        else:
            print(f"✗ 批量公式失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_formula_quick():
    """测试快捷函数"""
    print("\n" + "=" * 60)
    print("测试11: 快捷函数")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "formula_quick.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["销售额"],
                [100], [200], [300], [400], [500]
            ]
        )

        # 快捷求和
        result = ExcelOperations.formula_quick(
            path=xlsx_path,
            data_range="A2:A6",
            function="sum",
            output_cell="B1"
        )

        if result.get('success'):
            print(f"✓ 快捷函数成功: {result.get('formula')}")
            return True
        else:
            print(f"✗ 快捷函数失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_data_sort():
    """测试数据排序"""
    print("\n" + "=" * 60)
    print("测试12: 数据排序")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "data_sort.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["姓名", "部门", "销售额"],
                ["张三", "销售部", 50000],
                ["李四", "市场部", 80000],
                ["王五", "技术部", 30000]
            ]
        )

        # 按销售额降序排序
        result = ExcelOperations.data_sort(
            path=xlsx_path,
            range="A1:C4",
            sort_by=[{"col": "C", "order": "desc"}],
            has_header=True
        )

        if result.get('success'):
            print(f"✓ 排序成功: {result.get('rows_sorted')} 行")
            return True
        else:
            print(f"✗ 排序失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_data_validate():
    """测试数据验证"""
    print("\n" + "=" * 60)
    print("测试13: 数据验证")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "data_validate.xlsx")

    try:
        # 创建文件
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[["部门"]]
        )

        # 设置下拉列表
        result = ExcelOperations.data_validate(
            path=xlsx_path,
            range="A2:A10",
            type="list",
            values=["销售部", "技术部", "市场部", "财务部"]
        )

        if result.get('success'):
            print("✓ 数据验证（下拉列表）设置成功")
            return True
        else:
            print(f"✗ 数据验证失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_stats_summary():
    """测试统计摘要"""
    print("\n" + "=" * 60)
    print("测试14: 统计摘要")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "stats.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["数值"],
                [10], [20], [30], [40], [50]
            ]
        )

        # 计算统计摘要
        result = ExcelOperations.stats_summary(
            path=xlsx_path,
            data_range="A2:A6",
            metrics=["sum", "average", "max", "min", "count"]
        )

        if result.get('success'):
            stats = result.get('statistics', {})
            print(f"✓ 统计摘要成功:")
            print(f"  总和: {stats.get('sum')}")
            print(f"  平均: {stats.get('average')}")
            print(f"  最大: {stats.get('max')}")
            print(f"  最小: {stats.get('min')}")
            print(f"  计数: {stats.get('count')}")
            return True
        else:
            print(f"✗ 统计摘要失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_conditional_format():
    """测试条件格式"""
    print("\n" + "=" * 60)
    print("测试15: 条件格式")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "conditional.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["销售额"],
                [30000], [50000], [80000], [20000], [60000]
            ]
        )

        # 设置色阶
        result = ExcelOperations.conditional_format(
            path=xlsx_path,
            range="A2:A6",
            rule="color_scale",
            color_scale={"min_color": "F8696B", "max_color": "63BE7B"}
        )

        if result.get('success'):
            print("✓ 条件格式（色阶）设置成功")
            return True
        else:
            print(f"✗ 条件格式失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_chart_create():
    """测试图表创建"""
    print("\n" + "=" * 60)
    print("测试16: 图表创建")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "chart.xlsx")

    try:
        # 创建并写入数据
        ExcelOperations.create(path=xlsx_path)
        ExcelOperations.write_cell(
            path=xlsx_path,
            range="A1",
            data=[
                ["月份", "销售额"],
                ["1月", 100],
                ["2月", 150],
                ["3月", 200],
                ["4月", 180]
            ]
        )

        # 创建柱状图
        result = ExcelOperations.chart_create(
            path=xlsx_path,
            type="column",
            data_range="A1:B5",
            position="D1",
            title="月度销售趋势"
        )

        if result.get('success'):
            print(f"✓ 图表创建成功: {result.get('type')}")
            return True
        else:
            print(f"✗ 图表创建失败: {result.get('error')}")
            return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("DocuFlow Excel 功能测试")
    print("=" * 60)

    results = {
        # 基础测试 (9个)
        "工具注册": test_tool_registration(),
        "状态检查": test_excel_status(),
        "创建读取": test_create_and_read(),
        "单元格操作": test_cell_operations(),
        "格式化": test_formatting(),
        "公式": test_formula(),
        "工作表操作": test_sheet_operations(),
        "行列操作": test_row_col_operations(),
        "Excel转Word": test_excel_to_word(),
        # 高级功能测试 (7个)
        "批量公式": test_formula_batch(),
        "快捷函数": test_formula_quick(),
        "数据排序": test_data_sort(),
        "数据验证": test_data_validate(),
        "统计摘要": test_stats_summary(),
        "条件格式": test_conditional_format(),
        "图表创建": test_chart_create(),
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

    # 显示使用建议
    status = ExcelOperations.get_status()
    if not status.get('openpyxl_available'):
        print("\n" + "=" * 60)
        print("使用建议")
        print("=" * 60)
        print("提示: 安装openpyxl以启用Excel功能")
        print("  pip install openpyxl")

    return passed == total


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
