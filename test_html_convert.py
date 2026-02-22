# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, 'src')

from docuflow_mcp.extensions.html_to_pptx import HTMLToPPTXOperations

# 测试状态
print("=== 测试 html_to_pptx 模块 ===\n")

status = HTMLToPPTXOperations.get_status()
print('模块状态:')
print(f"  可用: {status['available']}")
print(f"  依赖: {status['dependencies']}")
print(f"  支持元素: {status['supported_elements']}")
print()

# 转换HTML到PPTX
print("开始转换 HTML -> PPTX ...")
result = HTMLToPPTXOperations.convert(
    html_source='test_output/slide_page.html',
    output_path='test_output/html_converted.pptx'
)

print(f"转换结果: {result}")
