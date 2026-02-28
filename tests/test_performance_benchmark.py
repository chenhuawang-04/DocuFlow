"""
DocuFlow MCP - Performance Benchmark Test
Day 12 - Performance testing and optimization

Tests the performance of various operations to establish baselines
and identify potential bottlenecks.
"""

import sys
import os
import time
import statistics
from typing import Dict, Any, List, Callable

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from docuflow_mcp.core.registry import dispatch_tool
from docx import Document


class PerformanceBenchmark:
    """Performance benchmark suite"""

    def __init__(self):
        self.test_dir = "E:/Project/DocuFlow/test_performance"
        self.results = {}

        # Ensure test directory exists
        os.makedirs(self.test_dir, exist_ok=True)

    def benchmark(self, name: str, func: Callable, iterations: int = 10) -> Dict[str, float]:
        """
        Run a function multiple times and collect timing statistics

        Args:
            name: Name of the benchmark
            func: Function to benchmark
            iterations: Number of iterations (default: 10)

        Returns:
            Dictionary with timing statistics
        """
        print(f"\n测试: {name}")
        print(f"  迭代次数: {iterations}")

        times = []
        for i in range(iterations):
            start = time.time()
            try:
                func()
                elapsed = time.time() - start
                times.append(elapsed)
            except Exception as e:
                print(f"  错误 (迭代 {i+1}): {str(e)[:100]}")
                continue

        if not times:
            print("  [FAIL] 所有迭代都失败")
            return {
                "success": False,
                "iterations": iterations,
                "successful": 0
            }

        result = {
            "success": True,
            "iterations": iterations,
            "successful": len(times),
            "min": min(times),
            "max": max(times),
            "mean": statistics.mean(times),
            "median": statistics.median(times),
            "stdev": statistics.stdev(times) if len(times) > 1 else 0
        }

        print(f"  平均: {result['mean']*1000:.2f}ms")
        print(f"  中位数: {result['median']*1000:.2f}ms")
        print(f"  最小: {result['min']*1000:.2f}ms")
        print(f"  最大: {result['max']*1000:.2f}ms")
        if len(times) > 1:
            print(f"  标准差: {result['stdev']*1000:.2f}ms")

        self.results[name] = result
        return result

    def cleanup(self):
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

    # ========== Benchmark Tests ==========

    def bench_doc_create(self):
        """Benchmark document creation"""
        def create_doc():
            path = os.path.join(self.test_dir, f"bench_create_{time.time()}.docx")
            dispatch_tool("doc_create", {"path": path})
            os.remove(path)

        self.benchmark("doc_create - 创建空文档", create_doc, iterations=50)

    def bench_paragraph_add(self):
        """Benchmark adding paragraphs"""
        def add_paragraph():
            path = os.path.join(self.test_dir, f"bench_para_{time.time()}.docx")
            dispatch_tool("doc_create", {"path": path})
            dispatch_tool("paragraph_add", {
                "path": path,
                "text": "Test paragraph with some content",
                "font_name": "宋体",
                "font_size": "12pt"
            })
            os.remove(path)

        self.benchmark("paragraph_add - 添加单个段落", add_paragraph, iterations=30)

    def bench_paragraph_batch(self):
        """Benchmark batch paragraph operations"""
        def batch_paragraphs():
            path = os.path.join(self.test_dir, f"bench_batch_{time.time()}.docx")
            dispatch_tool("doc_create", {"path": path})

            # Add 50 paragraphs
            for i in range(50):
                dispatch_tool("paragraph_add", {
                    "path": path,
                    "text": f"Paragraph {i}"
                })

            # Batch format
            dispatch_tool("batch_format_range", {
                "path": path,
                "start_index": 0,
                "end_index": 49,
                "font_name": "Arial",
                "font_size": "12pt"
            })

            os.remove(path)

        self.benchmark("batch_format_range - 批量格式化50段落", batch_paragraphs, iterations=5)

    def bench_table_operations(self):
        """Benchmark table operations"""
        def table_ops():
            path = os.path.join(self.test_dir, f"bench_table_{time.time()}.docx")
            dispatch_tool("doc_create", {"path": path})

            # Add table
            dispatch_tool("table_add", {
                "path": path,
                "rows": 10,
                "cols": 5
            })

            # Set cells
            for row in range(5):
                for col in range(5):
                    dispatch_tool("table_set_cell", {
                        "path": path,
                        "table_index": 0,
                        "row": row,
                        "col": col,
                        "text": f"R{row}C{col}"
                    })

            os.remove(path)

        self.benchmark("table_operations - 创建表格并填充", table_ops, iterations=10)

    def bench_template_create(self):
        """Benchmark template creation"""
        def create_from_template():
            path = os.path.join(self.test_dir, f"bench_template_{time.time()}.docx")
            dispatch_tool("template_create_from_preset", {
                "preset_name": "mba_thesis",
                "output_path": path
            })
            os.remove(path)

        self.benchmark("template_create_from_preset - 从模板创建", create_from_template, iterations=20)

    def bench_doc_compare(self):
        """Benchmark document comparison"""
        path1 = os.path.join(self.test_dir, "compare1.docx")
        path2 = os.path.join(self.test_dir, "compare2.docx")

        # Setup: Create two documents with differences
        dispatch_tool("doc_create", {"path": path1})
        dispatch_tool("doc_create", {"path": path2})

        for i in range(20):
            dispatch_tool("paragraph_add", {
                "path": path1,
                "text": f"Paragraph {i} version 1"
            })
            dispatch_tool("paragraph_add", {
                "path": path2,
                "text": f"Paragraph {i} version 2" if i % 3 == 0 else f"Paragraph {i} version 1"
            })

        def compare_docs():
            dispatch_tool("doc_compare", {
                "path1": path1,
                "path2": path2
            })

        self.benchmark("doc_compare - 对比20段文档", compare_docs, iterations=10)

        # Cleanup
        os.remove(path1)
        os.remove(path2)

    def bench_doc_statistics(self):
        """Benchmark document statistics analysis"""
        path = os.path.join(self.test_dir, "stats_test.docx")

        # Setup: Create a document with content
        dispatch_tool("doc_create", {"path": path})
        dispatch_tool("heading_add", {"path": path, "text": "Chapter 1", "level": 1})

        for i in range(50):
            dispatch_tool("paragraph_add", {
                "path": path,
                "text": f"This is paragraph {i} with some test content for statistics analysis."
            })

        dispatch_tool("table_add", {"path": path, "rows": 5, "cols": 3})

        def analyze_stats():
            dispatch_tool("doc_analyze_statistics", {
                "path": path,
                "detailed": True
            })

        self.benchmark("doc_analyze_statistics - 分析50段文档", analyze_stats, iterations=10)

        # Cleanup
        os.remove(path)

    def bench_word_frequency(self):
        """Benchmark word frequency analysis"""
        path = os.path.join(self.test_dir, "freq_test.docx")

        # Setup: Create document with text
        dispatch_tool("doc_create", {"path": path})

        for i in range(100):
            dispatch_tool("paragraph_add", {
                "path": path,
                "text": "This is a test document with repeated words. Analysis statistics frequency test document content."
            })

        def word_freq():
            dispatch_tool("doc_word_frequency", {
                "path": path,
                "top_n": 20
            })

        self.benchmark("doc_word_frequency - 词频分析100段", word_freq, iterations=10)

        # Cleanup
        os.remove(path)

    def bench_validate_format(self):
        """Benchmark format validation"""
        path = os.path.join(self.test_dir, "validate_test.docx")

        # Setup: Create document from template
        dispatch_tool("template_create_from_preset", {
            "preset_name": "business_report",
            "output_path": path
        })

        for i in range(30):
            dispatch_tool("paragraph_add", {
                "path": path,
                "text": f"Content paragraph {i}"
            })

        def validate():
            dispatch_tool("validate_format", {
                "path": path,
                "preset_rules": "business_report"
            })

        self.benchmark("validate_format - 验证30段文档", validate, iterations=10)

        # Cleanup
        os.remove(path)

    def bench_search_replace(self):
        """Benchmark search and replace"""
        path = os.path.join(self.test_dir, "search_test.docx")

        # Setup: Create document with searchable content
        dispatch_tool("doc_create", {"path": path})

        for i in range(100):
            dispatch_tool("paragraph_add", {
                "path": path,
                "text": f"This is paragraph {i} with target word appearing multiple times target target."
            })

        def search_replace():
            dispatch_tool("search_replace", {
                "path": path,
                "old_text": "target",
                "new_text": "replaced"
            })

        self.benchmark("search_replace - 替换100段中的词", search_replace, iterations=10)

        # Cleanup
        os.remove(path)

    def bench_export_operations(self):
        """Benchmark export operations"""
        path = os.path.join(self.test_dir, "export_test.docx")

        # Setup: Create document with content
        dispatch_tool("doc_create", {"path": path})

        for i in range(50):
            dispatch_tool("paragraph_add", {
                "path": path,
                "text": f"Paragraph {i} with content to export."
            })

        def export_text():
            dispatch_tool("export_to_text", {"path": path})

        def export_markdown():
            dispatch_tool("export_to_markdown", {"path": path})

        self.benchmark("export_to_text - 导出50段为文本", export_text, iterations=10)
        self.benchmark("export_to_markdown - 导出50段为Markdown", export_markdown, iterations=10)

        # Cleanup
        os.remove(path)

    def print_summary(self):
        """Print benchmark summary"""
        print("\n" + "="*70)
        print("  性能基准测试总结")
        print("="*70)
        print()

        # Sort by mean time
        sorted_results = sorted(
            [(name, data) for name, data in self.results.items() if data.get("success")],
            key=lambda x: x[1]["mean"]
        )

        print("按平均时间排序 (快 -> 慢):")
        print()
        print(f"{'操作':<50} {'平均时间':<15} {'中位数':<15}")
        print("-"*70)

        for name, data in sorted_results:
            mean_ms = data["mean"] * 1000
            median_ms = data["median"] * 1000
            print(f"{name:<50} {mean_ms:>12.2f}ms {median_ms:>12.2f}ms")

        print()
        print(f"总测试数: {len(self.results)}")
        print(f"成功: {sum(1 for r in self.results.values() if r.get('success'))}")
        print()

        # Performance categories
        fast = [name for name, data in self.results.items() if data.get("success") and data["mean"] < 0.05]
        medium = [name for name, data in self.results.items() if data.get("success") and 0.05 <= data["mean"] < 0.2]
        slow = [name for name, data in self.results.items() if data.get("success") and data["mean"] >= 0.2]

        print("性能分类:")
        print(f"  快速 (<50ms): {len(fast)} 个")
        print(f"  中等 (50-200ms): {len(medium)} 个")
        print(f"  较慢 (>200ms): {len(slow)} 个")

        if slow:
            print(f"\n  较慢操作: {', '.join(slow)}")

        print()
        print("="*70)


def main():
    """Run all performance benchmarks"""
    print("="*70)
    print("  DocuFlow MCP - Day 12 性能基准测试")
    print("="*70)

    # Import all modules to trigger registration
    from docuflow_mcp.extensions import templates, styles, batch, validator, advanced

    benchmark = PerformanceBenchmark()
    total_start = time.time()

    try:
        # Run all benchmarks
        print("\n" + "-"*70)
        print("  基础文档操作")
        print("-"*70)
        benchmark.bench_doc_create()
        benchmark.bench_paragraph_add()
        benchmark.bench_template_create()

        print("\n" + "-"*70)
        print("  批量操作")
        print("-"*70)
        benchmark.bench_paragraph_batch()
        benchmark.bench_table_operations()

        print("\n" + "-"*70)
        print("  高级分析")
        print("-"*70)
        benchmark.bench_doc_compare()
        benchmark.bench_doc_statistics()
        benchmark.bench_word_frequency()

        print("\n" + "-"*70)
        print("  格式验证")
        print("-"*70)
        benchmark.bench_validate_format()

        print("\n" + "-"*70)
        print("  搜索和导出")
        print("-"*70)
        benchmark.bench_search_replace()
        benchmark.bench_export_operations()

        total_elapsed = time.time() - total_start

        # Print summary
        benchmark.print_summary()

        print(f"总耗时: {total_elapsed:.2f}秒")
        print()

    finally:
        # Cleanup
        benchmark.cleanup()

    return 0


if __name__ == "__main__":
    exit(main())
