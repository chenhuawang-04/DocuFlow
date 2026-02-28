"""
DocuFlow OCR - 功能测试脚本

测试OCR识别功能
"""
import os
import sys
import tempfile
import pytest
import shutil

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from docuflow_mcp.extensions.ocr import OCROperations
from docuflow_mcp.core.registry import get_all_registered_tools


def test_tool_registration():
    """测试工具是否正确注册"""
    print("=" * 60)
    print("测试1: OCR工具注册检查")
    print("=" * 60)

    tools = get_all_registered_tools()
    ocr_tools = ['ocr_image', 'ocr_pdf', 'ocr_to_docx', 'ocr_status']

    registered = []
    missing = []

    for tool in ocr_tools:
        if tool in tools:
            registered.append(tool)
        else:
            missing.append(tool)

    print(f"已注册OCR工具: {registered}")
    if missing:
        print(f"未注册工具: {missing}")
        assert False, "未注册工具: ..."

    print(f"✓ 所有4个OCR工具已注册")
    print(f"  当前总工具数: {len(tools)}")

def test_ocr_status():
    """测试OCR状态检查"""
    print("\n" + "=" * 60)
    print("测试2: OCR状态检查")
    print("=" * 60)

    result = OCROperations.get_status()
    assert isinstance(result, dict), "Expected dict result"

    print(f"状态检查成功: {result.get('success', False)}")

    engines = result.get('engines', {})
    for name, info in engines.items():
        status = "✓ 可用" if info.get('available') else "✗ 不可用"
        print(f"  {name}: {status}")
        print(f"    描述: {info.get('description', 'N/A')}")

    deps = result.get('dependencies', {})
    print("\n依赖状态:")
    for name, info in deps.items():
        status = "✓" if info.get('available') else "✗"
        print(f"  {status} {name}: {info.get('purpose', '')}")

    print(f"\n支持的语言: {result.get('supported_languages', [])}")
    print(f"支持的图片格式: {result.get('supported_image_formats', [])}")

    assert result.get('success', False), "OCR状态检查失败"


def _create_test_image():
    """创建测试图片（如果PIL可用）"""
    try:
        from PIL import Image, ImageDraw, ImageFont

        # 创建临时目录
        temp_dir = tempfile.mkdtemp()
        image_path = os.path.join(temp_dir, "test_ocr.png")

        # 创建图片
        img = Image.new('RGB', (400, 200), color='white')
        draw = ImageDraw.Draw(img)

        # 尝试使用系统字体
        try:
            # Windows
            font = ImageFont.truetype("arial.ttf", 24)
        except (OSError, IOError):
            try:
                # 尝试其他字体
                font = ImageFont.truetype("simsun.ttc", 24)
            except (OSError, IOError):
                font = ImageFont.load_default()

        # 绘制文字
        draw.text((50, 30), "Hello World!", fill='black', font=font)
        draw.text((50, 80), "测试文字识别", fill='black', font=font)
        draw.text((50, 130), "OCR Test 2024", fill='black', font=font)

        img.save(image_path)
        return image_path, temp_dir
    except ImportError:
        return None, None


def test_ocr_image():
    """测试图片OCR"""
    print("\n" + "=" * 60)
    print("测试3: 图片OCR识别")
    print("=" * 60)

    # 创建测试图片
    image_path, temp_dir = _create_test_image()

    if image_path is None:
        print("跳过: PIL未安装，无法创建测试图片")
        pytest.skip("PIL未安装，无法创建测试图片")

    try:
        print(f"创建测试图片: {image_path}")

        # 测试OCR
        result = OCROperations.ocr_image(
            image_path=image_path,
            lang='auto',
            engine='auto'
        )

        print(f"识别结果: {result.get('success', False)}")

        if result.get('success'):
            print(f"使用引擎: {result.get('engine', 'N/A')}")
            print(f"置信度: {result.get('confidence', 0):.2f}")
            print(f"识别文本:\n{result.get('text', '')[:200]}")
            print("✓ 图片OCR测试通过!")
        else:
            error = result.get('error', '未知错误')
            if 'Tesseract' in error and 'anthropic' in error:
                print(f"跳过: 没有可用的OCR引擎")
                pytest.skip("没有可用的OCR引擎")
            print(f"✗ 识别失败: {error}")
            assert False, "识别失败: ..."

    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)


def test_confidence_estimation():
    """测试置信度估算"""
    print("\n" + "=" * 60)
    print("测试4: 置信度估算算法")
    print("=" * 60)

    test_cases = [
        ("Hello World 你好世界", "正常文本"),
        ("", "空文本"),
        ("###@@@!!!???", "全乱码"),
        ("这是一段正常的中文文本，包含标点符号。", "纯中文"),
        ("This is English text with numbers 123.", "纯英文"),
    ]

    all_pass = True
    for text, desc in test_cases:
        conf = OCROperations._estimate_confidence(text)
        status = "✓" if (conf > 0.5 and "乱码" not in desc) or (conf <= 0.5 and "乱码" in desc) or (conf == 0 and text == "") else "?"
        print(f"  {status} {desc}: {conf:.2f}")


def test_format_detection():
    """测试格式支持"""
    print("\n" + "=" * 60)
    print("测试5: 支持的格式")
    print("=" * 60)

    print("支持的图片格式:")
    for fmt in OCROperations.IMAGE_FORMATS:
        print(f"  - {fmt}")

    print("\n支持的语言:")
    for lang, code in OCROperations.LANG_MAP.items():
        print(f"  - {lang}: {code}")


def main():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("DocuFlow OCR 功能测试")
    print("=" * 60)

    results = {
        "工具注册": test_tool_registration(),
        "状态检查": test_ocr_status(),
        "图片OCR": test_ocr_image(),
        "置信度估算": test_confidence_estimation(),
        "格式支持": test_format_detection(),
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
    print("\n" + "=" * 60)
    print("使用建议")
    print("=" * 60)

    status = OCROperations.get_status()
    engines = status.get('engines', {})

    if not engines.get('tesseract', {}).get('available'):
        print("提示: 安装Tesseract可启用本地免费OCR")
        print("  Windows: https://github.com/UB-Mannheim/tesseract/wiki")
        print("  添加到PATH后重启终端")

    if not engines.get('claude', {}).get('available'):
        print("\n提示: 安装anthropic可启用Claude Vision增强OCR")
        print("  pip install anthropic")
        print("  设置环境变量: ANTHROPIC_API_KEY=your_key")

    deps = status.get('dependencies', {})
    if not deps.get('pdf2image', {}).get('available'):
        print("\n提示: 安装pdf2image可启用PDF OCR")
        print("  pip install pdf2image")
        print("  Windows还需安装poppler: https://github.com/oschwartz10612/poppler-windows/releases")

    return passed == total


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
