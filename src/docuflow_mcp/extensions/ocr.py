"""
DocuFlow OCR - 图像文字识别模块

支持：
- Tesseract OCR（本地、免费）
- Claude Vision（AI增强、复杂文档）
- PDF扫描件识别
- 识别结果转Word
"""
import os
import base64
import tempfile
import subprocess
from pathlib import Path
from typing import Optional, List, Dict, Any, Union
from dataclasses import dataclass

from ..core.registry import register_tool
from ..utils.deps import check_import, check_command


@dataclass
class OCRResult:
    """OCR识别结果"""
    text: str
    confidence: float  # 0-1
    engine: str  # 'tesseract' or 'claude'
    page: int = 1

    def to_dict(self) -> Dict[str, Any]:
        return {
            "text": self.text,
            "confidence": self.confidence,
            "engine": self.engine,
            "page": self.page
        }


class OCROperations:
    """OCR识别操作"""

    # 支持的图像格式
    IMAGE_FORMATS = ['.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp', '.gif', '.webp']

    # Tesseract语言映射
    LANG_MAP = {
        'chinese': 'chi_sim+chi_tra',
        'chinese_simplified': 'chi_sim',
        'chinese_traditional': 'chi_tra',
        'english': 'eng',
        'japanese': 'jpn',
        'korean': 'kor',
        'french': 'fra',
        'german': 'deu',
        'spanish': 'spa',
        'russian': 'rus',
        'auto': 'chi_sim+eng',  # 默认中英混合
    }

    @staticmethod
    def _image_to_base64(image_path: str) -> str:
        """将图片转为base64"""
        with open(image_path, 'rb') as f:
            return base64.standard_b64encode(f.read()).decode('utf-8')

    @staticmethod
    def _get_image_media_type(image_path: str) -> str:
        """获取图片的media type"""
        ext = Path(image_path).suffix.lower()
        media_types = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.webp': 'image/webp',
            '.bmp': 'image/bmp',
            '.tiff': 'image/tiff',
            '.tif': 'image/tiff',
        }
        return media_types.get(ext, 'image/png')

    @staticmethod
    def _pdf_to_images(pdf_path: str, dpi: int = 200) -> List[str]:
        """将PDF转换为图片列表"""
        if not check_import("pdf2image"):
            raise ImportError("需要安装pdf2image: pip install pdf2image")

        from pdf2image import convert_from_path

        # 创建临时目录
        temp_dir = tempfile.mkdtemp(prefix='docuflow_ocr_')

        # 转换PDF为图片
        images = convert_from_path(pdf_path, dpi=dpi)

        image_paths = []
        for i, image in enumerate(images):
            image_path = os.path.join(temp_dir, f'page_{i+1}.png')
            image.save(image_path, 'PNG')
            image_paths.append(image_path)

        return image_paths

    @staticmethod
    def _ocr_with_tesseract(image_path: str, lang: str = 'auto') -> OCRResult:
        """使用Tesseract进行OCR"""
        if not check_command("tesseract"):
            raise RuntimeError("Tesseract未安装或不可用")

        # 映射语言代码
        tess_lang = OCROperations.LANG_MAP.get(lang, lang)

        try:
            # 使用subprocess直接调用tesseract
            result = subprocess.run(
                ['tesseract', image_path, 'stdout', '-l', tess_lang],
                capture_output=True,
                text=True,
                timeout=60,
                encoding='utf-8'
            )

            if result.returncode == 0:
                text = result.stdout.strip()
                # 简单估算置信度（基于文本长度和字符类型）
                confidence = OCROperations._estimate_confidence(text)
                return OCRResult(
                    text=text,
                    confidence=confidence,
                    engine='tesseract'
                )
            else:
                raise RuntimeError(f"Tesseract错误: {result.stderr}")

        except subprocess.TimeoutExpired:
            raise RuntimeError("Tesseract处理超时")

    @staticmethod
    def _estimate_confidence(text: str) -> float:
        """估算OCR结果的置信度"""
        if not text:
            return 0.0

        # 基于文本特征估算置信度
        total_chars = len(text)
        if total_chars == 0:
            return 0.0

        # 计算有效字符比例（字母、数字、中文、常用标点）
        valid_chars = sum(1 for c in text if c.isalnum() or '\u4e00' <= c <= '\u9fff' or c in '，。！？、；：""''（）')
        valid_ratio = valid_chars / total_chars

        # 计算乱码特征（连续特殊字符）
        import re
        garbage_pattern = re.compile(r'[^\w\s\u4e00-\u9fff，。！？、；：""''（）]{3,}')
        garbage_matches = garbage_pattern.findall(text)
        garbage_ratio = sum(len(m) for m in garbage_matches) / total_chars if garbage_matches else 0

        # 综合评分
        confidence = valid_ratio * 0.7 + (1 - garbage_ratio) * 0.3
        return max(0.0, min(1.0, confidence))

    @staticmethod
    def _ocr_with_claude(image_path: str, api_key: Optional[str] = None,
                         prompt: Optional[str] = None,
                         model: Optional[str] = None) -> OCRResult:
        """使用Claude Vision进行OCR"""
        if not check_import("anthropic"):
            raise ImportError("需要安装anthropic: pip install anthropic")

        import anthropic

        # 获取API key
        key = api_key or os.environ.get('ANTHROPIC_AUTH_TOKEN') or os.environ.get('ANTHROPIC_API_KEY')
        if not key:
            raise ValueError("需要提供ANTHROPIC_AUTH_TOKEN")

        client = anthropic.Anthropic(api_key=key)

        # 准备图片数据
        image_data = OCROperations._image_to_base64(image_path)
        media_type = OCROperations._get_image_media_type(image_path)

        # 构建prompt
        default_prompt = """请仔细识别这张图片中的所有文字内容。

要求：
1. 完整提取图片中的所有文字，包括标题、正文、表格、页眉页脚等
2. 保持原文的段落结构和层级关系
3. 如果有表格，用markdown表格格式呈现
4. 如果有列表，保持列表格式
5. 只输出识别到的文字内容，不要添加任何解释或说明

请开始识别："""

        actual_prompt = prompt or default_prompt

        # 调用Claude Vision API
        message = client.messages.create(
            model=model or "claude-sonnet-4-20250514",
            max_tokens=4096,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": image_data,
                            },
                        },
                        {
                            "type": "text",
                            "text": actual_prompt
                        }
                    ],
                }
            ],
        )

        text = message.content[0].text.strip()

        return OCRResult(
            text=text,
            confidence=0.95,  # Claude通常有较高准确率
            engine='claude'
        )

    @register_tool("ocr_image",
                   required_params=['image_path'],
                   optional_params=['lang', 'engine', 'api_key', 'prompt', 'model'])
    @staticmethod
    def ocr_image(image_path: str,
                  lang: str = 'auto',
                  engine: str = 'auto',
                  api_key: Optional[str] = None,
                  prompt: Optional[str] = None,
                  model: Optional[str] = None) -> Dict[str, Any]:
        """
        OCR识别单张图片

        Args:
            image_path: 图片文件路径
            lang: 识别语言 (auto/chinese/english/japanese等)
            engine: OCR引擎 (auto/tesseract/claude)
            api_key: Claude API密钥（可选，也可用环境变量）
            prompt: 自定义Claude识别提示词

        Returns:
            {success, text, confidence, engine, message}
        """
        try:
            path = Path(image_path)
            if not path.exists():
                return {"success": False, "error": f"图片不存在: {image_path}"}

            if path.suffix.lower() not in OCROperations.IMAGE_FORMATS:
                return {"success": False, "error": f"不支持的图片格式: {path.suffix}"}

            # 选择引擎
            if engine == 'auto':
                # 先尝试Tesseract
                if check_command("tesseract"):
                    result = OCROperations._ocr_with_tesseract(image_path, lang)
                    # 如果置信度低且Claude可用，使用Claude增强
                    if result.confidence < 0.6 and check_import("anthropic"):
                        try:
                            result = OCROperations._ocr_with_claude(image_path, api_key, prompt, model)
                        except (RuntimeError, ValueError, OSError):
                            pass  # 保持Tesseract结果
                elif check_import("anthropic"):
                    result = OCROperations._ocr_with_claude(image_path, api_key, prompt, model)
                else:
                    return {"success": False, "error": "没有可用的OCR引擎（需要Tesseract或anthropic）"}

            elif engine == 'tesseract':
                if not check_command("tesseract"):
                    return {"success": False, "error": "Tesseract未安装"}
                result = OCROperations._ocr_with_tesseract(image_path, lang)

            elif engine == 'claude':
                if not check_import("anthropic"):
                    return {"success": False, "error": "需要安装anthropic: pip install anthropic"}
                result = OCROperations._ocr_with_claude(image_path, api_key, prompt, model)

            else:
                return {"success": False, "error": f"未知引擎: {engine}"}

            return {
                "success": True,
                "text": result.text,
                "confidence": result.confidence,
                "engine": result.engine,
                "message": f"识别成功，使用{result.engine}引擎"
            }

        except Exception as e:
            return {"success": False, "error": str(e)}

    @register_tool("ocr_pdf",
                   required_params=['pdf_path'],
                   optional_params=['pages', 'lang', 'engine', 'dpi', 'api_key', 'prompt', 'model'])
    @staticmethod
    def ocr_pdf(pdf_path: str,
                pages: Optional[List[int]] = None,
                lang: str = 'auto',
                engine: str = 'auto',
                dpi: int = 200,
                api_key: Optional[str] = None,
                prompt: Optional[str] = None,
                model: Optional[str] = None) -> Dict[str, Any]:
        """
        OCR识别PDF文档（支持扫描件）

        Args:
            pdf_path: PDF文件路径
            pages: 要识别的页码列表（从1开始，None表示全部）
            lang: 识别语言
            engine: OCR引擎 (auto/tesseract/claude)
            dpi: PDF转图片的DPI（越高越清晰，但越慢）
            api_key: Claude API密钥
            prompt: 自定义识别提示词

        Returns:
            {success, total_pages, results, full_text, message}
        """
        try:
            path = Path(pdf_path)
            if not path.exists():
                return {"success": False, "error": f"PDF不存在: {pdf_path}"}

            if path.suffix.lower() != '.pdf':
                return {"success": False, "error": "文件不是PDF格式"}

            if not check_import("pdf2image"):
                return {"success": False, "error": "需要安装pdf2image: pip install pdf2image"}

            # PDF转图片
            image_paths = OCROperations._pdf_to_images(pdf_path, dpi)
            if not image_paths:
                return {"success": False, "error": "PDF转图片失败：未生成任何图片"}

            temp_dir = os.path.dirname(image_paths[0])
            total_pages = len(image_paths)

            try:
                # 确定要处理的页面
                if pages:
                    target_pages = [p for p in pages if 1 <= p <= total_pages]
                else:
                    target_pages = list(range(1, total_pages + 1))

                results = []
                full_text_parts = []

                for page_num in target_pages:
                    image_path = image_paths[page_num - 1]

                    # 对每页进行OCR
                    ocr_result = OCROperations.ocr_image(
                        image_path=image_path,
                        lang=lang,
                        engine=engine,
                        api_key=api_key,
                        prompt=prompt,
                        model=model
                    )

                    if ocr_result.get("success"):
                        page_result = {
                            "page": page_num,
                            "text": ocr_result["text"],
                            "confidence": ocr_result["confidence"],
                            "engine": ocr_result["engine"]
                        }
                        results.append(page_result)
                        full_text_parts.append(f"=== 第 {page_num} 页 ===\n{ocr_result['text']}")
                    else:
                        results.append({
                            "page": page_num,
                            "error": ocr_result.get("error", "识别失败")
                        })

            finally:
                # 无论成功或异常，都清理临时文件
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)

            # 计算平均置信度
            valid_results = [r for r in results if "confidence" in r]
            avg_confidence = sum(r["confidence"] for r in valid_results) / len(valid_results) if valid_results else 0

            return {
                "success": True,
                "total_pages": total_pages,
                "processed_pages": len(target_pages),
                "average_confidence": avg_confidence,
                "results": results,
                "full_text": "\n\n".join(full_text_parts),
                "message": f"成功识别 {len(valid_results)}/{len(target_pages)} 页"
            }

        except Exception as e:
            return {"success": False, "error": str(e)}

    @register_tool("ocr_to_docx",
                   required_params=['source', 'output_path'],
                   optional_params=['lang', 'engine', 'dpi', 'api_key', 'prompt', 'title', 'model'])
    @staticmethod
    def ocr_to_docx(source: str,
                    output_path: str,
                    lang: str = 'auto',
                    engine: str = 'auto',
                    dpi: int = 200,
                    api_key: Optional[str] = None,
                    prompt: Optional[str] = None,
                    title: Optional[str] = None,
                    model: Optional[str] = None) -> Dict[str, Any]:
        """
        OCR识别后直接生成Word文档

        Args:
            source: 源文件路径（PDF或图片）
            output_path: 输出Word文档路径
            lang: 识别语言
            engine: OCR引擎
            dpi: PDF转换DPI
            api_key: Claude API密钥
            prompt: 自定义识别提示词
            title: 文档标题

        Returns:
            {success, output_path, pages, message}
        """
        try:
            source_path = Path(source)
            if not source_path.exists():
                return {"success": False, "error": f"源文件不存在: {source}"}

            # 根据文件类型选择处理方式
            ext = source_path.suffix.lower()

            if ext == '.pdf':
                ocr_result = OCROperations.ocr_pdf(
                    pdf_path=source,
                    lang=lang,
                    engine=engine,
                    dpi=dpi,
                    api_key=api_key,
                    prompt=prompt,
                    model=model
                )
                if not ocr_result.get("success"):
                    return ocr_result
                full_text = ocr_result.get("full_text", "")
                pages = ocr_result.get("processed_pages", 1)

            elif ext in OCROperations.IMAGE_FORMATS:
                ocr_result = OCROperations.ocr_image(
                    image_path=source,
                    lang=lang,
                    engine=engine,
                    api_key=api_key,
                    prompt=prompt,
                    model=model
                )
                if not ocr_result.get("success"):
                    return ocr_result
                full_text = ocr_result.get("text", "")
                pages = 1

            else:
                return {"success": False, "error": f"不支持的文件格式: {ext}"}

            # 生成Word文档
            try:
                from docx import Document
                from docx.shared import Pt
                from docx.enum.text import WD_ALIGN_PARAGRAPH
            except ImportError:
                return {"success": False, "error": "需要安装python-docx"}

            doc = Document()

            # 添加标题
            if title:
                heading = doc.add_heading(title, 0)
                heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 添加内容
            for line in full_text.split('\n'):
                line = line.strip()
                if not line:
                    continue

                # 检测是否是页面分隔符
                if line.startswith('=== 第') and line.endswith('==='):
                    doc.add_page_break()
                    doc.add_heading(line.strip('= '), level=1)
                else:
                    para = doc.add_paragraph(line)
                    para.paragraph_format.first_line_indent = Pt(21)  # 首行缩进

            # 保存文档
            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            doc.save(output_path)

            return {
                "success": True,
                "output_path": output_path,
                "pages": pages,
                "message": f"OCR识别完成，已生成Word文档: {output_path}"
            }

        except Exception as e:
            return {"success": False, "error": str(e)}

    @register_tool("ocr_status",
                   required_params=[],
                   optional_params=[])
    @staticmethod
    def get_status() -> Dict[str, Any]:
        """
        获取OCR模块状态和可用引擎

        Returns:
            {success, engines, languages, message}
        """
        engines = {
            "tesseract": {
                "available": check_command("tesseract"),
                "description": "开源OCR引擎，本地运行，免费",
                "best_for": "简单文档、批量处理"
            },
            "claude": {
                "available": check_import("anthropic"),
                "api_key_set": bool(os.environ.get('ANTHROPIC_AUTH_TOKEN') or os.environ.get('ANTHROPIC_API_KEY')),
                "description": "Claude Vision AI，需要API密钥",
                "best_for": "复杂布局、手写体、需要理解的文档"
            }
        }

        dependencies = {
            "pdf2image": {
                "available": check_import("pdf2image"),
                "purpose": "PDF转图片（OCR PDF必需）"
            },
            "PIL": {
                "available": check_import("PIL"),
                "purpose": "图像处理"
            }
        }

        available_engines = [name for name, info in engines.items() if info.get("available")]

        return {
            "success": True,
            "engines": engines,
            "dependencies": dependencies,
            "available_engines": available_engines,
            "supported_languages": list(OCROperations.LANG_MAP.keys()),
            "supported_image_formats": OCROperations.IMAGE_FORMATS,
            "message": f"可用引擎: {', '.join(available_engines) if available_engines else '无'}"
        }
