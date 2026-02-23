"""
DocuFlow Converter - 基于pandoc的文档格式转换

支持40+格式互转，包括docx/pdf/md/html/latex/epub等
"""
import os
import subprocess
from pathlib import Path
from typing import Optional, List, Dict, Any

from ..core.registry import register_tool
from ..utils.deps import check_command


class ConverterOperations:
    """文档格式转换操作"""

    # pandoc支持的主要输入格式
    INPUT_FORMATS = [
        'docx', 'markdown', 'md', 'html', 'latex', 'tex',
        'epub', 'odt', 'rst', 'textile', 'mediawiki', 'org',
        'json', 'csv', 'tsv', 'rtf', 'txt', 'docbook',
        'fb2', 'ipynb', 'man', 'muse', 'opml', 't2t', 'wiki'
    ]

    # pandoc支持的主要输出格式
    OUTPUT_FORMATS = [
        'docx', 'pdf', 'markdown', 'md', 'html', 'latex', 'tex',
        'epub', 'odt', 'rst', 'beamer', 'pptx', 'rtf', 'plain',
        'asciidoc', 'mediawiki', 'org', 'json', 'docbook',
        'fb2', 'ipynb', 'man', 'ms', 'muse', 'opml', 'texinfo',
        'textile', 'slideous', 'slidy', 'dzslides', 'revealjs', 's5'
    ]

    # 格式别名映射
    FORMAT_ALIASES = {
        'md': 'markdown',
        'tex': 'latex',
        'txt': 'plain',
        'word': 'docx',
        'powerpoint': 'pptx',
        'text': 'plain',
    }

    @staticmethod
    def _normalize_format(fmt: str) -> str:
        """标准化格式名称"""
        fmt = fmt.lower().strip('.')
        return ConverterOperations.FORMAT_ALIASES.get(fmt, fmt)

    @staticmethod
    def _detect_format(file_path: str) -> str:
        """从文件扩展名检测格式"""
        ext = Path(file_path).suffix.lower().strip('.')
        return ConverterOperations._normalize_format(ext)

    @staticmethod
    def _run_pandoc(args: List[str]) -> Dict[str, Any]:
        """执行pandoc命令"""
        try:
            result = subprocess.run(
                ['pandoc'] + args,
                capture_output=True,
                text=True,
                timeout=300  # 5分钟超时
            )
            if result.returncode == 0:
                return {"success": True, "output": result.stdout}
            else:
                return {"success": False, "error": result.stderr or "Pandoc执行失败"}
        except subprocess.TimeoutExpired:
            return {"success": False, "error": "转换超时（超过5分钟）"}
        except FileNotFoundError:
            return {"success": False, "error": "pandoc未安装或不在PATH中"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @register_tool("convert",
                   required_params=['source'],
                   optional_params=['target', 'source_format', 'target_format', 'extra_args'])
    @staticmethod
    def convert(source: str,
                target: Optional[str] = None,
                source_format: Optional[str] = None,
                target_format: Optional[str] = None,
                extra_args: Optional[List[str]] = None) -> Dict[str, Any]:
        """
        通用文档格式转换

        Args:
            source: 源文件路径
            target: 目标文件路径（可选，不指定则自动生成）
            source_format: 源格式（可选，自动检测）
            target_format: 目标格式（可选，从target路径推断）
            extra_args: pandoc额外参数（如 ['--toc', '--standalone']）

        Returns:
            {success, source, target, source_format, target_format, message}
        """
        try:
            source_path = Path(source)
            if not source_path.exists():
                return {"success": False, "error": f"源文件不存在: {source}"}

            # 检测/验证源格式
            src_fmt = source_format or ConverterOperations._detect_format(source)
            src_fmt = ConverterOperations._normalize_format(src_fmt)

            # 确定目标格式和路径
            if target:
                tgt_fmt = target_format or ConverterOperations._detect_format(target)
            elif target_format:
                tgt_fmt = target_format
                target = str(source_path.with_suffix(f'.{tgt_fmt}'))
            else:
                return {"success": False, "error": "必须指定target或target_format"}

            tgt_fmt = ConverterOperations._normalize_format(tgt_fmt)

            # 构建pandoc命令参数
            args = [
                '-f', src_fmt,
                '-t', tgt_fmt,
                '-o', target,
                source
            ]

            # 添加额外参数
            extra = extra_args or []

            # PDF需要特殊处理（支持中文）
            if tgt_fmt == 'pdf':
                # 检查是否已指定pdf-engine
                has_engine = any('--pdf-engine' in arg for arg in extra)
                if not has_engine:
                    extra.extend(['--pdf-engine=xelatex'])
                # 添加中文支持
                if not any('-V' in arg and 'CJK' in arg for arg in extra):
                    extra.extend(['-V', 'CJKmainfont=SimSun'])

            args = extra + args  # extra_args放在前面

            # 执行转换
            result = ConverterOperations._run_pandoc(args)

            if result["success"]:
                return {
                    "success": True,
                    "source": source,
                    "target": target,
                    "source_format": src_fmt,
                    "target_format": tgt_fmt,
                    "message": f"转换成功: {src_fmt} -> {tgt_fmt}"
                }
            else:
                return {
                    "success": False,
                    "source": source,
                    "target_format": tgt_fmt,
                    "error": result["error"]
                }

        except Exception as e:
            return {"success": False, "error": str(e)}

    @register_tool("convert_batch",
                   required_params=['sources', 'target_format'],
                   optional_params=['output_dir', 'extra_args'])
    @staticmethod
    def convert_batch(sources: List[str],
                      target_format: str,
                      output_dir: Optional[str] = None,
                      extra_args: Optional[List[str]] = None) -> Dict[str, Any]:
        """
        批量转换多个文件

        Args:
            sources: 源文件路径列表
            target_format: 目标格式
            output_dir: 输出目录（可选，默认同目录）
            extra_args: pandoc额外参数

        Returns:
            {success, total, converted, failed, results}
        """
        results = []
        converted = 0
        failed = 0

        tgt_fmt = ConverterOperations._normalize_format(target_format)

        for src in sources:
            src_path = Path(src)
            if output_dir:
                out_dir = Path(output_dir)
                out_dir.mkdir(parents=True, exist_ok=True)
                target = str(out_dir / f"{src_path.stem}.{tgt_fmt}")
            else:
                target = str(src_path.with_suffix(f'.{tgt_fmt}'))

            result = ConverterOperations.convert(
                source=src,
                target=target,
                target_format=tgt_fmt,
                extra_args=extra_args
            )

            results.append(result)
            if result.get("success"):
                converted += 1
            else:
                failed += 1

        return {
            "success": failed == 0,
            "total": len(sources),
            "converted": converted,
            "failed": failed,
            "results": results
        }

    @register_tool("convert_formats",
                   required_params=[],
                   optional_params=[])
    @staticmethod
    def get_formats() -> Dict[str, Any]:
        """
        获取支持的格式列表

        Returns:
            {success, input_formats, output_formats, popular_conversions, pandoc_available}
        """
        popular = [
            {"from": "docx", "to": "pdf", "desc": "Word转PDF"},
            {"from": "docx", "to": "markdown", "desc": "Word转Markdown"},
            {"from": "markdown", "to": "docx", "desc": "Markdown转Word"},
            {"from": "markdown", "to": "pdf", "desc": "Markdown转PDF"},
            {"from": "markdown", "to": "html", "desc": "Markdown转HTML"},
            {"from": "html", "to": "docx", "desc": "HTML转Word"},
            {"from": "latex", "to": "pdf", "desc": "LaTeX转PDF"},
            {"from": "epub", "to": "docx", "desc": "电子书转Word"},
            {"from": "docx", "to": "epub", "desc": "Word转电子书"},
            {"from": "markdown", "to": "pptx", "desc": "Markdown转PowerPoint"},
        ]

        pandoc_ok = check_command("pandoc")

        return {
            "success": True,
            "pandoc_available": pandoc_ok,
            "input_formats": ConverterOperations.INPUT_FORMATS,
            "output_formats": ConverterOperations.OUTPUT_FORMATS,
            "popular_conversions": popular,
            "total_input": len(ConverterOperations.INPUT_FORMATS),
            "total_output": len(ConverterOperations.OUTPUT_FORMATS),
            "message": "pandoc已就绪" if pandoc_ok else "警告: pandoc未安装或不可用"
        }

    @register_tool("convert_with_template",
                   required_params=['source', 'target'],
                   optional_params=['template', 'css', 'reference_doc', 'extra_args'])
    @staticmethod
    def convert_with_template(source: str,
                              target: str,
                              template: Optional[str] = None,
                              css: Optional[str] = None,
                              reference_doc: Optional[str] = None,
                              extra_args: Optional[List[str]] = None) -> Dict[str, Any]:
        """
        带模板/样式的转换

        Args:
            source: 源文件路径
            target: 目标文件路径
            template: pandoc模板文件（用于HTML/LaTeX等）
            css: CSS样式文件（用于HTML输出）
            reference_doc: 参考文档（用于docx/pptx，继承样式）
            extra_args: 其他pandoc参数

        Returns:
            {success, source, target, message}
        """
        extra = list(extra_args) if extra_args else []

        # 验证模板/参考文档是否存在
        if template:
            if not Path(template).exists():
                return {"success": False, "error": f"模板文件不存在: {template}"}
            extra.extend(['--template', template])

        if css:
            if not Path(css).exists():
                return {"success": False, "error": f"CSS文件不存在: {css}"}
            extra.extend(['--css', css])

        if reference_doc:
            if not Path(reference_doc).exists():
                return {"success": False, "error": f"参考文档不存在: {reference_doc}"}
            extra.extend(['--reference-doc', reference_doc])

        result = ConverterOperations.convert(
            source=source,
            target=target,
            extra_args=extra if extra else None
        )

        if result.get("success"):
            result["message"] = f"带模板转换成功: {source} -> {target}"

        return result
