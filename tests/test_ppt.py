#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DocuFlow PPT 功能测试脚本
"""
import os
import sys
import tempfile
import shutil

# 添加项目路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from docuflow_mcp.extensions.ppt import PPTOperations
from docuflow_mcp.core.registry import get_all_registered_tools


def test_tool_registration():
    """测试PPT工具是否正确注册"""
    print("\n" + "=" * 60)
    print("测试1: PPT工具注册检查")
    print("=" * 60)

    tools = get_all_registered_tools()
    ppt_tools = [
        'ppt_create', 'ppt_read', 'ppt_info', 'ppt_set_properties', 'ppt_merge',
        'slide_add', 'slide_delete', 'slide_duplicate', 'slide_get_layouts',
        'shape_add_text', 'shape_add_image', 'shape_add_table', 'shape_add_shape',
        'slide_set_background', 'slide_add_notes', 'ppt_status',
        # 新增母版和动画工具
        'master_list', 'master_get_info', 'placeholder_list', 'placeholder_set',
        'animation_add', 'animation_list', 'animation_remove',
        # 新增图表工具
        'chart_add', 'ppt_chart_modify', 'chart_get_data', 'chart_list', 'chart_delete'
    ]

    registered = []
    missing = []

    for tool in ppt_tools:
        if tool in tools:
            registered.append(tool)
        else:
            missing.append(tool)

    print(f"已注册PPT工具: {len(registered)}/{len(ppt_tools)}")
    if missing:
        print(f"未注册工具: {missing}")
        assert False, "未注册工具: ..."

    print(f"✓ 所有{len(ppt_tools)}个PPT工具已注册")
    print(f"  当前总工具数: {len(tools)}")

def test_ppt_status():
    """测试PPT模块状态"""
    print("\n" + "=" * 60)
    print("测试2: PPT模块状态检查")
    print("=" * 60)

    result = PPTOperations.get_status()
    assert isinstance(result, dict), "Expected dict result"

    if result.get('success'):
        print(f"python-pptx可用: {result.get('pptx_available')}")
        if result.get('version'):
            print(f"版本: {result.get('version')}")

        features = result.get('features', [])
        if features:
            print("\n支持的功能:")
            for f in features:
                print(f"  - {f}")
    else:
        print(f"✗ 状态检查失败: {result.get('error')}")
        assert False, "状态检查失败: {result.get("


def test_ppt_create():
    """测试创建PPT"""
    print("\n" + "=" * 60)
    print("测试3: 创建PPT文档")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        result = PPTOperations.create(
            path=ppt_path,
            title="测试PPT文档"
        )

        if result.get('success'):
            print(f"✓ 创建PPT成功: {ppt_path}")
        else:
            print(f"✗ 创建PPT失败: {result.get('error')}")
            assert False, "创建PPT失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_slide_add():
    """测试添加幻灯片"""
    print("\n" + "=" * 60)
    print("测试4: 添加幻灯片")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 先创建PPT
        PPTOperations.create(path=ppt_path)

        # 添加幻灯片
        result = PPTOperations.slide_add(
            path=ppt_path,
            layout="Blank"
        )

        if result.get('success'):
            print(f"✓ 添加幻灯片成功")
            print(f"  幻灯片索引: {result.get('slide_index')}")
            print(f"  布局: {result.get('layout')}")
        else:
            print(f"✗ 添加幻灯片失败: {result.get('error')}")
            assert False, "添加幻灯片失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_shape_add_text():
    """测试添加文本框"""
    print("\n" + "=" * 60)
    print("测试5: 添加文本框")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)

        # 添加文本框
        result = PPTOperations.shape_add_text(
            path=ppt_path,
            slide=1,
            text="Hello, PowerPoint!",
            left="1in",
            top="1in",
            width="8in",
            height="1in",
            font_size=24,
            bold=True
        )

        if result.get('success'):
            print(f"✓ 添加文本框成功")
            print(f"  文本: {result.get('text')}")
        else:
            print(f"✗ 添加文本框失败: {result.get('error')}")
            assert False, "添加文本框失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_shape_add_table():
    """测试添加表格"""
    print("\n" + "=" * 60)
    print("测试6: 添加表格")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)

        # 添加表格
        result = PPTOperations.shape_add_table(
            path=ppt_path,
            slide=1,
            rows=3,
            cols=3,
            data=[
                ["A", "B", "C"],
                ["1", "2", "3"],
                ["4", "5", "6"]
            ]
        )

        if result.get('success'):
            print(f"✓ 添加表格成功")
            print(f"  表格大小: {result.get('table_size')}")
        else:
            print(f"✗ 添加表格失败: {result.get('error')}")
            assert False, "添加表格失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_shape_add_shape():
    """测试添加形状"""
    print("\n" + "=" * 60)
    print("测试7: 添加形状")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)

        # 添加形状
        result = PPTOperations.shape_add_shape(
            path=ppt_path,
            slide=1,
            shape_type="rectangle",
            fill_color="0066CC",
            text="Shape Text"
        )

        if result.get('success'):
            print(f"✓ 添加形状成功")
            print(f"  形状类型: {result.get('shape_type')}")
        else:
            print(f"✗ 添加形状失败: {result.get('error')}")
            assert False, "添加形状失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_ppt_read():
    """测试读取PPT"""
    print("\n" + "=" * 60)
    print("测试8: 读取PPT内容")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加内容
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.shape_add_text(path=ppt_path, slide=1, text="Test Content")

        # 读取PPT
        result = PPTOperations.read(path=ppt_path)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 读取PPT成功")
            print(f"  幻灯片数: {result.get('total_slides')}")
        else:
            print(f"✗ 读取PPT失败: {result.get('error')}")
            assert False, "读取PPT失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_ppt_info():
    """测试获取PPT信息"""
    print("\n" + "=" * 60)
    print("测试9: 获取PPT信息")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT
        PPTOperations.create(path=ppt_path, title="Test PPT")
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)

        # 获取信息
        result = PPTOperations.info(path=ppt_path)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 获取PPT信息成功")
            print(f"  幻灯片数: {result.get('total_slides')}")
            print(f"  尺寸: {result.get('dimensions')}")
        else:
            print(f"✗ 获取PPT信息失败: {result.get('error')}")
            assert False, "获取PPT信息失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_slide_get_layouts():
    """测试获取布局列表"""
    print("\n" + "=" * 60)
    print("测试10: 获取布局列表")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT
        PPTOperations.create(path=ppt_path)

        # 获取布局列表
        result = PPTOperations.slide_get_layouts(path=ppt_path)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 获取布局列表成功")
            print(f"  可用布局数: {result.get('total_layouts')}")
            for layout in result.get('layouts', [])[:5]:
                print(f"    - {layout['name']}")
        else:
            print(f"✗ 获取布局列表失败: {result.get('error')}")
            assert False, "获取布局列表失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_slide_add_notes():
    """测试添加演讲者备注"""
    print("\n" + "=" * 60)
    print("测试11: 添加演讲者备注")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)

        # 添加备注
        result = PPTOperations.slide_add_notes(
            path=ppt_path,
            slide=1,
            notes="这是演讲者备注内容"
        )

        if result.get('success'):
            print(f"✓ 添加备注成功")
            print(f"  备注长度: {result.get('notes_length')} 字符")
        else:
            print(f"✗ 添加备注失败: {result.get('error')}")
            assert False, "添加备注失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_slide_delete():
    """测试删除幻灯片"""
    print("\n" + "=" * 60)
    print("测试12: 删除幻灯片")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加多个幻灯片
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)

        # 删除幻灯片
        result = PPTOperations.slide_delete(
            path=ppt_path,
            index=2
        )

        if result.get('success'):
            print(f"✓ 删除幻灯片成功")
            print(f"  删除索引: {result.get('deleted_index')}")
            print(f"  剩余幻灯片: {result.get('remaining_slides')}")
        else:
            print(f"✗ 删除幻灯片失败: {result.get('error')}")
            assert False, "删除幻灯片失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_master_list():
    """测试获取母版列表"""
    print("\n" + "=" * 60)
    print("测试13: 获取母版列表")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT
        PPTOperations.create(path=ppt_path)

        # 获取母版列表
        result = PPTOperations.master_list(path=ppt_path)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 获取母版列表成功")
            print(f"  母版数量: {result.get('master_count')}")
            masters = result.get('masters', [])
            for m in masters[:2]:  # 只显示前2个
                print(f"    - {m.get('name')}: {m.get('layout_count')} 个布局")
        else:
            print(f"✗ 获取母版列表失败: {result.get('error')}")
            assert False, "获取母版列表失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_master_get_info():
    """测试获取母版详细信息"""
    print("\n" + "=" * 60)
    print("测试14: 获取母版详细信息")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT
        PPTOperations.create(path=ppt_path)

        # 获取母版详细信息
        result = PPTOperations.master_get_info(path=ppt_path, master_index=0)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 获取母版信息成功")
            print(f"  母版名称: {result.get('name')}")
            print(f"  布局数量: {result.get('layout_count')}")
            print(f"  占位符数量: {result.get('placeholder_count')}")
        else:
            print(f"✗ 获取母版信息失败: {result.get('error')}")
            assert False, "获取母版信息失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_placeholder_list():
    """测试获取占位符列表"""
    print("\n" + "=" * 60)
    print("测试15: 获取占位符列表")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path, layout="Title Slide")

        # 获取占位符列表
        result = PPTOperations.placeholder_list(path=ppt_path, slide=1)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 获取占位符列表成功")
            print(f"  占位符数量: {result.get('placeholder_count')}")
            placeholders = result.get('placeholders', [])
            for p in placeholders[:3]:  # 只显示前3个
                print(f"    - idx={p.get('idx')}: {p.get('type')}")
        else:
            print(f"✗ 获取占位符列表失败: {result.get('error')}")
            assert False, "获取占位符列表失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_placeholder_set():
    """测试设置占位符内容"""
    print("\n" + "=" * 60)
    print("测试16: 设置占位符内容")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加带占位符的幻灯片
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path, layout="Title Slide")

        # 先获取占位符列表
        list_result = PPTOperations.placeholder_list(path=ppt_path, slide=1)
        placeholders = list_result.get('placeholders', [])

        if not placeholders:
            print("✗ 没有找到占位符")
            assert False, "没有找到占位符"

        # 设置第一个占位符的内容
        idx = placeholders[0].get('idx')
        result = PPTOperations.placeholder_set(
            path=ppt_path,
            slide=1,
            idx=idx,
            text="测试标题文本",
            font_size=32,
            bold=True
        )

        if result.get('success'):
            print(f"✓ 设置占位符成功")
            print(f"  占位符索引: {result.get('idx')}")
            print(f"  文本: {result.get('text')}")
        else:
            print(f"✗ 设置占位符失败: {result.get('error')}")
            assert False, "设置占位符失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_animation_add():
    """测试添加动画"""
    print("\n" + "=" * 60)
    print("测试17: 添加动画")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片和形状
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.shape_add_text(
            path=ppt_path,
            slide=1,
            text="动画测试文本",
            left="2in",
            top="2in"
        )

        # 添加淡入动画
        result = PPTOperations.animation_add(
            path=ppt_path,
            slide=1,
            shape_index=0,
            effect="fade",
            trigger="on_click",
            duration=0.5
        )

        if result.get('success'):
            print(f"✓ 添加动画成功")
            print(f"  动画效果: {result.get('effect')}")
            print(f"  触发方式: {result.get('trigger')}")
        else:
            print(f"✗ 添加动画失败: {result.get('error')}")
            assert False, "添加动画失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_animation_list():
    """测试列出动画"""
    print("\n" + "=" * 60)
    print("测试18: 列出动画")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片和形状
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.shape_add_text(path=ppt_path, slide=1, text="测试")

        # 添加动画
        PPTOperations.animation_add(
            path=ppt_path,
            slide=1,
            shape_index=0,
            effect="fly_in",
            direction="left"
        )

        # 列出动画
        result = PPTOperations.animation_list(path=ppt_path, slide=1)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 列出动画成功")
            print(f"  动画数量: {result.get('animation_count')}")
        else:
            print(f"✗ 列出动画失败: {result.get('error')}")
            assert False, "列出动画失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_animation_remove():
    """测试删除动画"""
    print("\n" + "=" * 60)
    print("测试19: 删除动画")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片和形状
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.shape_add_text(path=ppt_path, slide=1, text="测试")

        # 添加动画
        PPTOperations.animation_add(
            path=ppt_path,
            slide=1,
            shape_index=0,
            effect="zoom"
        )

        # 删除所有动画
        result = PPTOperations.animation_remove(
            path=ppt_path,
            slide=1,
            remove_all=True
        )

        if result.get('success'):
            print(f"✓ 删除动画成功")
            print(f"  删除数量: {result.get('removed_count', 'all')}")
        else:
            print(f"✗ 删除动画失败: {result.get('error')}")
            assert False, "删除动画失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_chart_add():
    """测试添加图表"""
    print("\n" + "=" * 60)
    print("测试20: 添加图表")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)

        # 添加柱状图
        result = PPTOperations.chart_add(
            path=ppt_path,
            slide=1,
            chart_type='column',
            categories=['Q1', 'Q2', 'Q3'],
            series=[{"name": "销售额", "values": [100, 200, 150]}],
            title="季度销售"
        )

        if result.get('success'):
            print(f"✓ 添加图表成功")
            print(f"  图表类型: {result.get('chart_type')}")
            print(f"  位置: {result.get('position')}")
        else:
            print(f"✗ 添加图表失败: {result.get('error')}")
            assert False, "添加图表失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_chart_list():
    """测试列出图表"""
    print("\n" + "=" * 60)
    print("测试21: 列出图表")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片和图表
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.chart_add(
            path=ppt_path, slide=1, chart_type='column',
            categories=['A', 'B'], series=[{"name": "S1", "values": [1, 2]}]
        )

        # 列出图表
        result = PPTOperations.chart_list(path=ppt_path, slide=1)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 列出图表成功")
            print(f"  图表数量: {result.get('chart_count')}")
            assert result.get('chart_count', 0) > 0, "图表数量应大于0"
        else:
            print(f"✗ 列出图表失败: {result.get('error')}")
            assert False, "列出图表失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_chart_get_data():
    """测试获取图表数据"""
    print("\n" + "=" * 60)
    print("测试22: 获取图表数据")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片和图表
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.chart_add(
            path=ppt_path, slide=1, chart_type='pie',
            categories=['A', 'B', 'C'], series=[{"name": "Data", "values": [30, 50, 20]}]
        )

        # 获取图表数据
        result = PPTOperations.chart_get_data(path=ppt_path, slide=1, chart_index=0)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 获取图表数据成功")
            print(f"  图表类型: {result.get('chart_type')}")
            print(f"  分类数: {len(result.get('categories', []))}")
        else:
            print(f"✗ 获取图表数据失败: {result.get('error')}")
            assert False, "获取图表数据失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_chart_modify():
    """测试修改图表"""
    print("\n" + "=" * 60)
    print("测试23: 修改图表")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片和图表
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.chart_add(
            path=ppt_path, slide=1, chart_type='line',
            categories=['1月', '2月'], series=[{"name": "趋势", "values": [10, 20]}]
        )

        # 修改图表
        result = PPTOperations.chart_modify(
            path=ppt_path, slide=1, chart_index=0,
            title="新标题", has_legend=True, legend_position="bottom"
        )

        if result.get('success'):
            print(f"✓ 修改图表成功")
            print(f"  修改内容: {result.get('modifications', [])}")
        else:
            print(f"✗ 修改图表失败: {result.get('error')}")
            assert False, "修改图表失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_chart_delete():
    """测试删除图表"""
    print("\n" + "=" * 60)
    print("测试24: 删除图表")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()

    try:
        ppt_path = os.path.join(temp_dir, "test.pptx")

        # 创建PPT并添加幻灯片和图表
        PPTOperations.create(path=ppt_path)
        PPTOperations.slide_add(path=ppt_path)
        PPTOperations.chart_add(
            path=ppt_path, slide=1, chart_type='bar',
            categories=['X', 'Y'], series=[{"name": "S", "values": [5, 10]}]
        )

        # 删除图表
        result = PPTOperations.chart_delete(path=ppt_path, slide=1, chart_index=0)
        assert isinstance(result, dict), "Expected dict result"

        if result.get('success'):
            print(f"✓ 删除图表成功")
            print(f"  剩余形状数: {result.get('remaining_shapes')}")
        else:
            print(f"✗ 删除图表失败: {result.get('error')}")
            assert False, "删除图表失败: {result.get("

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("DocuFlow PPT 功能测试")
    print("=" * 60)

    results = {
        "工具注册": test_tool_registration(),
        "状态检查": test_ppt_status(),
        "创建PPT": test_ppt_create(),
        "添加幻灯片": test_slide_add(),
        "添加文本框": test_shape_add_text(),
        "添加表格": test_shape_add_table(),
        "添加形状": test_shape_add_shape(),
        "读取PPT": test_ppt_read(),
        "获取信息": test_ppt_info(),
        "获取布局": test_slide_get_layouts(),
        "添加备注": test_slide_add_notes(),
        "删除幻灯片": test_slide_delete(),
        # 新增母版和动画测试
        "母版列表": test_master_list(),
        "母版详情": test_master_get_info(),
        "占位符列表": test_placeholder_list(),
        "设置占位符": test_placeholder_set(),
        "添加动画": test_animation_add(),
        "列出动画": test_animation_list(),
        "删除动画": test_animation_remove(),
        # 新增图表测试
        "添加图表": test_chart_add(),
        "列出图表": test_chart_list(),
        "获取图表数据": test_chart_get_data(),
        "修改图表": test_chart_modify(),
        "删除图表": test_chart_delete(),
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
        return 0
    else:
        print(f"\n{total - passed} 个测试失败")
        return 1


if __name__ == "__main__":
    sys.exit(main())
