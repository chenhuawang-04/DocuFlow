"""
DocuFlow MCP - Final Integration Test
Day 12 - Comprehensive integration testing

Tests all 68 tools across all modules to ensure:
- All tools are properly registered
- All tools work correctly
- Cross-module integration works
- Error handling is robust
- Complete workflows function properly
"""

import sys
import os
import time
from typing import Dict, Any, List

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from docuflow_mcp.core.registry import dispatch_tool, get_all_registered_tools
from docx import Document

# Ensure all tool modules are loaded so decorators register tools
import docuflow_mcp.server  # noqa: F401


def print_section(title: str):
    """Print a section header"""
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}\n")


def print_subsection(title: str):
    """Print a subsection header"""
    print(f"\n{'-'*70}")
    print(f"  {title}")
    print(f"{'-'*70}\n")


class IntegrationTest:
    """Comprehensive integration test suite"""

    def __init__(self):
        self.test_dir = "E:/Project/DocuFlow/test_integration"
        self.results = []
        self.total_tests = 0
        self.passed_tests = 0

        # Ensure test directory exists
        os.makedirs(self.test_dir, exist_ok=True)

    def run_test(self, name: str, func):
        """Run a single test and track results"""
        self.total_tests += 1
        print(f"测试 {self.total_tests}: {name}")

        try:
            start_time = time.time()
            result = func()
            assert isinstance(result, dict), "Expected dict result"
            elapsed = time.time() - start_time

            if result:
                self.passed_tests += 1
                print(f"  [OK] 通过 ({elapsed:.3f}秒)")
                self.results.append((name, True, elapsed, None))
            else:
                print(f"  [FAIL] 失败 ({elapsed:.3f}秒)")
                self.results.append((name, False, elapsed, "测试返回False"))
        except Exception as e:
            elapsed = time.time() - start_time
            print(f"  [ERROR] 异常: {str(e)[:100]} ({elapsed:.3f}秒)")
            self.results.append((name, False, elapsed, str(e)[:200]))

    def cleanup_test_files(self):
        """Clean up test files"""
        try:
            if os.path.exists(self.test_dir):
                for file in os.listdir(self.test_dir):
                    file_path = os.path.join(self.test_dir, file)
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                os.rmdir(self.test_dir)
        except Exception as e:
            print(f"清理失败: {e}")

    # ========== Module 1: Document Operations ==========

    def test_document_create_and_read(self):
        """Test doc_create and doc_read"""
        test_file = os.path.join(self.test_dir, "test_create.docx")

        # Create document
        result = dispatch_tool("doc_create", {
            "path": test_file,
            "title": "Test Document"
        })
        if not result.get("success"):
            assert False, "test failed"

        # Read document
        result = dispatch_tool("doc_read", {
            "path": test_file
        })
        if not result.get("success"):
            assert False, "test failed"

        os.remove(test_file)
    def test_document_info(self):
        """Test doc_info"""
        test_file = os.path.join(self.test_dir, "test_info.docx")

        dispatch_tool("doc_create", {"path": test_file})
        result = dispatch_tool("doc_info", {"path": test_file})
        assert isinstance(result, dict), "Expected dict result"

        os.remove(test_file)
        return result.get("success") and "statistics" in result

    def test_document_merge(self):
        """Test doc_merge"""
        doc1 = os.path.join(self.test_dir, "merge1.docx")
        doc2 = os.path.join(self.test_dir, "merge2.docx")
        output = os.path.join(self.test_dir, "merged.docx")

        dispatch_tool("doc_create", {"path": doc1})
        dispatch_tool("doc_create", {"path": doc2})

        result = dispatch_tool("doc_merge", {
            "paths": [doc1, doc2],
            "output_path": output
        })

        os.remove(doc1)
        os.remove(doc2)
        os.remove(output)
        return result.get("success")

    # ========== Module 2: Paragraph Operations ==========

    def test_paragraph_operations(self):
        """Test paragraph add, modify, get, delete"""
        test_file = os.path.join(self.test_dir, "test_para.docx")
        dispatch_tool("doc_create", {"path": test_file})

        # Add paragraph
        result1 = dispatch_tool("paragraph_add", {
            "path": test_file,
            "text": "Test paragraph",
            "font_name": "宋体",
            "font_size": "12pt"
        })

        # Get paragraph
        result2 = dispatch_tool("paragraph_get", {
            "path": test_file,
            "index": 0
        })

        # Modify paragraph
        result3 = dispatch_tool("paragraph_modify", {
            "path": test_file,
            "index": 0,
            "text": "Modified paragraph"
        })

        # Delete paragraph
        result4 = dispatch_tool("paragraph_delete", {
            "path": test_file,
            "index": 0
        })

        os.remove(test_file)
        return all([
            result1.get("success"),
            result2.get("success"),
            result3.get("success"),
            result4.get("success")
        ])

    # ========== Module 3: Heading and TOC ==========

    def test_heading_and_outline(self):
        """Test heading_add and heading_get_outline"""
        test_file = os.path.join(self.test_dir, "test_heading.docx")
        dispatch_tool("doc_create", {"path": test_file})

        # Add headings
        dispatch_tool("heading_add", {
            "path": test_file,
            "text": "Chapter 1",
            "level": 1
        })
        dispatch_tool("heading_add", {
            "path": test_file,
            "text": "Section 1.1",
            "level": 2
        })

        # Get outline
        result = dispatch_tool("heading_get_outline", {
            "path": test_file
        })

        os.remove(test_file)
        return result.get("success") and len(result.get("outline", [])) == 2

    def test_toc_add(self):
        """Test toc_add"""
        test_file = os.path.join(self.test_dir, "test_toc.docx")
        dispatch_tool("doc_create", {"path": test_file})

        result = dispatch_tool("toc_add", {"path": test_file})
        assert isinstance(result, dict), "Expected dict result"

        os.remove(test_file)
        return result.get("success")

    # ========== Module 4: Table Operations ==========

    def test_table_full_workflow(self):
        """Test complete table workflow"""
        test_file = os.path.join(self.test_dir, "test_table.docx")
        dispatch_tool("doc_create", {"path": test_file})

        # Add table
        result1 = dispatch_tool("table_add", {
            "path": test_file,
            "rows": 3,
            "cols": 2
        })

        # Set cell
        result2 = dispatch_tool("table_set_cell", {
            "path": test_file,
            "table_index": 0,
            "row": 0,
            "col": 0,
            "text": "Header"
        })

        # Get table
        result3 = dispatch_tool("table_get", {
            "path": test_file,
            "table_index": 0
        })

        # Add row
        result4 = dispatch_tool("table_add_row", {
            "path": test_file,
            "table_index": 0
        })

        # Set column width
        result5 = dispatch_tool("table_set_column_width", {
            "path": test_file,
            "table_index": 0,
            "col_index": 0,
            "width": "5cm"
        })

        # Merge cells
        result6 = dispatch_tool("table_merge_cells", {
            "path": test_file,
            "table_index": 0,
            "start_row": 0,
            "start_col": 0,
            "end_row": 0,
            "end_col": 1
        })

        os.remove(test_file)
        return all([
            result1.get("success"),
            result2.get("success"),
            result3.get("success"),
            result4.get("success"),
            result5.get("success"),
            result6.get("success")
        ])

    # ========== Module 5: Image and Link ==========

    def test_image_add(self):
        """Test image_add (without actual image file)"""
        test_file = os.path.join(self.test_dir, "test_image.docx")
        dispatch_tool("doc_create", {"path": test_file})

        # Note: This will fail without a real image file, which is expected
        # We're just testing the tool is registered and callable
        result = dispatch_tool("image_add", {
            "path": test_file,
            "image_path": "nonexistent.png"
        })

        os.remove(test_file)
        # Expected to fail, but should return error gracefully
        return not result.get("success") and "error" in result

    def test_hyperlink_add(self):
        """Test hyperlink_add"""
        test_file = os.path.join(self.test_dir, "test_link.docx")
        dispatch_tool("doc_create", {"path": test_file})

        result = dispatch_tool("hyperlink_add", {
            "path": test_file,
            "text": "Example",
            "url": "https://example.com"
        })

        os.remove(test_file)
        return result.get("success")

    # ========== Module 6: List Operations ==========

    def test_lists(self):
        """Test list_add_bullet and list_add_numbered"""
        test_file = os.path.join(self.test_dir, "test_list.docx")
        dispatch_tool("doc_create", {"path": test_file})

        result1 = dispatch_tool("list_add_bullet", {
            "path": test_file,
            "items": ["Item 1", "Item 2", "Item 3"]
        })

        result2 = dispatch_tool("list_add_numbered", {
            "path": test_file,
            "items": ["Step 1", "Step 2", "Step 3"]
        })

        os.remove(test_file)
        return result1.get("success") and result2.get("success")

    # ========== Module 7: Page Setup ==========

    def test_page_setup(self):
        """Test page setup operations"""
        test_file = os.path.join(self.test_dir, "test_page.docx")
        dispatch_tool("doc_create", {"path": test_file})

        result1 = dispatch_tool("page_set_margins", {
            "path": test_file,
            "top": "2cm",
            "bottom": "2cm",
            "left": "3cm",
            "right": "2cm"
        })

        result2 = dispatch_tool("page_set_size", {
            "path": test_file,
            "width": "21cm",
            "height": "29.7cm"
        })

        result3 = dispatch_tool("page_add_break", {
            "path": test_file
        })

        result4 = dispatch_tool("header_set", {
            "path": test_file,
            "text": "Test Header"
        })

        result5 = dispatch_tool("footer_set", {
            "path": test_file,
            "text": "Test Footer"
        })

        os.remove(test_file)
        return all([
            result1.get("success"),
            result2.get("success"),
            result3.get("success"),
            result4.get("success"),
            result5.get("success")
        ])

    # ========== Module 8: Template System ==========

    def test_template_system(self):
        """Test template operations"""
        test_file = os.path.join(self.test_dir, "test_template.docx")

        # List presets
        result1 = dispatch_tool("template_list_presets", {})

        # Create from preset
        result2 = dispatch_tool("template_create_from_preset", {
            "preset_name": "mba_thesis",
            "output_path": test_file,
            "title": "Test Thesis"
        })

        os.remove(test_file)
        return result1.get("success") and result2.get("success")

    # ========== Module 9: Style Management ==========

    def test_style_management(self):
        """Test style operations"""
        test_file = os.path.join(self.test_dir, "test_style.docx")
        dispatch_tool("doc_create", {"path": test_file})

        # Get styles
        result1 = dispatch_tool("doc_get_styles", {
            "path": test_file
        })

        # Create style
        result2 = dispatch_tool("style_create", {
            "path": test_file,
            "style_name": "CustomStyle",
            "base_style": "Normal"
        })

        # Export styles
        result3 = dispatch_tool("style_export", {
            "path": test_file
        })

        os.remove(test_file)
        return all([
            result1.get("success"),
            result2.get("success"),
            result3.get("success")
        ])

    # ========== Module 10: Batch Operations ==========

    def test_batch_operations(self):
        """Test batch operations"""
        test_file = os.path.join(self.test_dir, "test_batch.docx")
        dispatch_tool("doc_create", {"path": test_file})

        # Add multiple paragraphs
        for i in range(10):
            dispatch_tool("paragraph_add", {
                "path": test_file,
                "text": f"Paragraph {i}"
            })

        # Batch format range
        result1 = dispatch_tool("batch_format_range", {
            "path": test_file,
            "start_index": 0,
            "end_index": 4,
            "font_name": "Arial",
            "font_size": "14pt"
        })

        os.remove(test_file)
        return result1.get("success")

    # ========== Module 11: Format Validation ==========

    def test_validation(self):
        """Test validation operations"""
        test_file = os.path.join(self.test_dir, "test_validate.docx")
        dispatch_tool("template_create_from_preset", {
            "preset_name": "business_report",
            "output_path": test_file
        })

        result1 = dispatch_tool("validate_format", {
            "path": test_file,
            "preset_rules": "business_report"
        })

        result2 = dispatch_tool("validate_check_consistency", {
            "path": test_file
        })

        os.remove(test_file)
        return result1.get("success") and result2.get("success")

    # ========== Module 12: Search and Replace ==========

    def test_search_replace(self):
        """Test search and replace operations"""
        test_file = os.path.join(self.test_dir, "test_search.docx")
        dispatch_tool("doc_create", {"path": test_file})
        dispatch_tool("paragraph_add", {
            "path": test_file,
            "text": "Hello World. Hello everyone."
        })

        result1 = dispatch_tool("search_find", {
            "path": test_file,
            "text": "Hello"
        })

        result2 = dispatch_tool("search_replace", {
            "path": test_file,
            "old_text": "Hello",
            "new_text": "Hi"
        })

        os.remove(test_file)
        return result1.get("success") and result2.get("success")

    # ========== Module 13: Advanced Operations ==========

    def test_advanced_operations(self):
        """Test advanced operations"""
        test_file1 = os.path.join(self.test_dir, "test_adv1.docx")
        test_file2 = os.path.join(self.test_dir, "test_adv2.docx")

        # Create two documents
        dispatch_tool("doc_create", {"path": test_file1})
        dispatch_tool("paragraph_add", {
            "path": test_file1,
            "text": "Version 1 content"
        })

        dispatch_tool("doc_create", {"path": test_file2})
        dispatch_tool("paragraph_add", {
            "path": test_file2,
            "text": "Version 2 content"
        })

        # Compare documents
        result1 = dispatch_tool("doc_compare", {
            "path1": test_file1,
            "path2": test_file2
        })

        # Analyze statistics
        result2 = dispatch_tool("doc_analyze_statistics", {
            "path": test_file1
        })

        # Metadata operations
        result3 = dispatch_tool("doc_set_metadata", {
            "path": test_file1,
            "title": "Test Doc",
            "author": "Test Author"
        })

        result4 = dispatch_tool("doc_get_metadata", {
            "path": test_file1
        })

        # Word frequency
        result5 = dispatch_tool("doc_word_frequency", {
            "path": test_file1,
            "top_n": 10
        })

        os.remove(test_file1)
        os.remove(test_file2)

        return all([
            result1.get("success"),
            result2.get("success"),
            result3.get("success"),
            result4.get("success"),
            result5.get("success")
        ])

    # ========== Module 14: Export Operations ==========

    def test_export_operations(self):
        """Test export operations"""
        test_file = os.path.join(self.test_dir, "test_export.docx")
        dispatch_tool("doc_create", {"path": test_file})
        dispatch_tool("paragraph_add", {
            "path": test_file,
            "text": "Export test content"
        })

        result1 = dispatch_tool("export_to_text", {
            "path": test_file
        })

        result2 = dispatch_tool("export_to_markdown", {
            "path": test_file
        })

        os.remove(test_file)
        return result1.get("success") and result2.get("success")

    # ========== Complete Workflow Test ==========

    def test_complete_workflow(self):
        """Test a complete real-world workflow"""
        test_file = os.path.join(self.test_dir, "complete_workflow.docx")

        # Step 1: Create document from template
        result1 = dispatch_tool("template_create_from_preset", {
            "preset_name": "mba_thesis",
            "output_path": test_file,
            "title": "Complete Test Document"
        })
        if not result1.get("success"):
            assert False, "test failed"

        # Step 2: Add content structure
        dispatch_tool("heading_add", {
            "path": test_file,
            "text": "Chapter 1: Introduction",
            "level": 1
        })
        dispatch_tool("paragraph_add", {
            "path": test_file,
            "text": "This is the introduction paragraph with important information."
        })

        # Step 3: Add a table
        dispatch_tool("table_add", {
            "path": test_file,
            "rows": 3,
            "cols": 2
        })
        dispatch_tool("table_set_cell", {
            "path": test_file,
            "table_index": 0,
            "row": 0,
            "col": 0,
            "text": "Category"
        })

        # Step 4: Set metadata
        result4 = dispatch_tool("doc_set_metadata", {
            "path": test_file,
            "title": "Complete Test",
            "author": "Test Suite",
            "keywords": "test,integration,workflow"
        })
        if not result4.get("success"):
            assert False, "test failed"

        # Step 5: Analyze statistics
        result5 = dispatch_tool("doc_analyze_statistics", {
            "path": test_file,
            "detailed": True
        })
        if not result5.get("success"):
            assert False, "test failed"

        # Step 6: Validate format
        result6 = dispatch_tool("validate_format", {
            "path": test_file,
            "preset_rules": "mba_thesis"
        })
        if not result6.get("success"):
            assert False, "test failed"

        # Step 7: Export to markdown
        result7 = dispatch_tool("export_to_markdown", {
            "path": test_file
        })
        if not result7.get("success"):
            assert False, "test failed"

        os.remove(test_file)
    # ========== Tool Registration Test ==========

    def test_all_tools_registered(self):
        """Verify all tools are registered"""
        tools = get_all_registered_tools()

        # Minimum expected count (grows as new tools are added)
        min_expected = 140
        if len(tools) < min_expected:
            print(f"  警告: 预期至少 {min_expected} 个工具，实际找到 {len(tools)} 个")
            assert False, "警告: 预期至少 ... 个工具，实际找到 ... 个"

        # Check critical tools exist
        critical_tools = [
            "doc_create", "doc_read", "doc_info",
            "paragraph_add", "heading_add", "table_add",
            "template_create_from_preset", "style_create",
            "batch_format_range", "validate_format",
            "doc_compare", "doc_analyze_statistics"
        ]

        for tool in critical_tools:
            if tool not in tools:
                print(f"  错误: 关键工具 '{tool}' 未注册")
                assert False, "错误: 关键工具"


def main():
    """Run all integration tests"""
    print_section("DocuFlow MCP - Day 12 综合集成测试")

    # Import all modules to trigger registration
    from docuflow_mcp.extensions import templates, styles, batch, validator, advanced

    tester = IntegrationTest()
    start_time = time.time()

    # Run all tests
    print_subsection("模块1: 文档操作")
    tester.run_test("文档创建和读取", tester.test_document_create_and_read)
    tester.run_test("文档信息", tester.test_document_info)
    tester.run_test("文档合并", tester.test_document_merge)

    print_subsection("模块2: 段落操作")
    tester.run_test("段落全流程", tester.test_paragraph_operations)

    print_subsection("模块3: 标题和目录")
    tester.run_test("标题和大纲", tester.test_heading_and_outline)
    tester.run_test("添加目录", tester.test_toc_add)

    print_subsection("模块4: 表格操作")
    tester.run_test("表格全流程", tester.test_table_full_workflow)

    print_subsection("模块5: 图片和链接")
    tester.run_test("添加图片", tester.test_image_add)
    tester.run_test("添加超链接", tester.test_hyperlink_add)

    print_subsection("模块6: 列表操作")
    tester.run_test("列表", tester.test_lists)

    print_subsection("模块7: 页面设置")
    tester.run_test("页面设置", tester.test_page_setup)

    print_subsection("模块8: 模板系统")
    tester.run_test("模板系统", tester.test_template_system)

    print_subsection("模块9: 样式管理")
    tester.run_test("样式管理", tester.test_style_management)

    print_subsection("模块10: 批量操作")
    tester.run_test("批量操作", tester.test_batch_operations)

    print_subsection("模块11: 格式验证")
    tester.run_test("格式验证", tester.test_validation)

    print_subsection("模块12: 搜索和替换")
    tester.run_test("搜索和替换", tester.test_search_replace)

    print_subsection("模块13: 高级操作")
    tester.run_test("高级操作", tester.test_advanced_operations)

    print_subsection("模块14: 导出操作")
    tester.run_test("导出操作", tester.test_export_operations)

    print_subsection("完整工作流测试")
    tester.run_test("完整工作流", tester.test_complete_workflow)

    print_subsection("工具注册测试")
    tester.run_test("所有工具已注册", tester.test_all_tools_registered)

    total_elapsed = time.time() - start_time

    # Clean up
    tester.cleanup_test_files()

    # Print summary
    print_section("测试总结")

    print(f"总测试数: {tester.total_tests}")
    print(f"通过: {tester.passed_tests}")
    print(f"失败: {tester.total_tests - tester.passed_tests}")
    print(f"通过率: {tester.passed_tests / tester.total_tests * 100:.1f}%")
    print(f"总耗时: {total_elapsed:.2f}秒")
    print()

    # Detailed results
    print("详细结果:")
    for name, passed, elapsed, error in tester.results:
        status = "[OK]  " if passed else "[FAIL]"
        print(f"  {status} {name} ({elapsed:.3f}秒)")
        if error:
            print(f"        错误: {error}")

    print()
    if tester.passed_tests == tester.total_tests:
        print("="*70)
        print("  [SUCCESS] All tests passed! System is working properly")
        print("="*70)
        return 0
    else:
        print("="*70)
        print(f"  [WARNING] {tester.total_tests - tester.passed_tests} tests failed")
        print("="*70)
        return 1


if __name__ == "__main__":
    exit(main())
