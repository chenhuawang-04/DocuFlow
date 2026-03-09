# DocuFlow MCP

英文文档见 [README.md](README.md)。

DocuFlow 是一个面向文档处理的 MCP (Model Context Protocol) 服务，覆盖 Word、Excel、PowerPoint、PDF、格式转换、OCR、HTML 转 PPTX 以及 AI 图像生成等工作流，目前共 **149 个工具**。

项目目标是提供统一的 MCP 工具面，便于 Claude Code、Codex 等 MCP 客户端直接对文档执行操作。

## 功能覆盖

DocuFlow 的工具按模块划分如下：

| 模块 | 工具数 | 能力范围 |
| --- | ---: | --- |
| Word (.docx) | 49 | 文档、段落、标题、表格、图片、列表、分页、页眉/页脚、批注、样式、模板、导出 |
| Excel (.xlsx) | 33 | 工作簿、工作表、单元格、公式、统计、图表、数据透视表、条件格式、数据处理 |
| PowerPoint (.pptx) | 30 | 幻灯片、文本框、形状、图表、动画、切换、母版、占位符 |
| PDF | 23 | 提取、合并、拆分、旋转、水印、表单、涂黑、转换 |
| 格式转换 | 4 | 基于 pandoc 的 40+ 格式互转 |
| OCR | 4 | Tesseract 或 OpenAI 兼容 completion 的图像与 PDF OCR |
| HTML -> PPTX | 3 | 将幻灯片式 HTML 转为 PPTX |
| AI 图像生成 | 3 | 文生图与 PPT 插图生成 |

## 安装

### 方式一：安装脚本

```bash
cd DocuFlow
python install.py
```

### 方式二：可编辑安装

```bash
pip install -e .
```

安装完成后，MCP 入口命令为：

```bash
docuflow-mcp
```

## MCP 客户端配置

在 MCP 客户端配置中加入：

```json
{
  "mcpServers": {
    "docuflow": {
      "command": "docuflow-mcp"
    }
  }
}
```

修改配置后重启客户端。

常见提示词示例：

- “创建一个名为 report.docx 的 Word 文档，标题是 Monthly Sales Report。”
- “从 invoice.pdf 中提取表格到 Excel。”
- “将 report.docx 转成 PDF。”
- “对 scan.png 执行 OCR，只返回识别文本。”

## 使用方式

DocuFlow 支持两种使用方式。

### 1. MCP 客户端模式

推荐方式。Claude Code、Codex 等客户端连接 `docuflow-mcp` 并通过 MCP 调用工具。

### 2. 直接 Python API 模式

可直接导入内部模块，例如：

```python
from docuflow_mcp.extensions.ocr import OCROperations
```

该方式适用于开发、测试或自定义集成，与 MCP 客户端模式不同。

## OCR 架构

远程 OCR 统一通过 **OpenAI 兼容的 `chat/completions` 接口** 实现，不依赖 Anthropic 或 OpenAI 的 Python SDK。

OCR 工具：

- `ocr_image`
- `ocr_pdf`
- `ocr_to_docx`
- `ocr_status`

### OCR 引擎

- `tesseract`
  本地 OCR，不需要远程 API。
- `completion`
  通过 OpenAI 兼容的 `chat/completions` 端点进行远程 OCR。
- `claude`
  兼容别名，内部映射到 `completion`。
- `auto`
  优先使用 Tesseract，不可用时回退到 `completion`。

## OCR 配置

远程 OCR 从项目根目录的 `ocr_config.json` 读取配置：

```json
{
  "api_url": "https://your-api.example.com/v1/chat/completions",
  "model": "grok-4.1-thinking",
  "timeout": 120,
  "api_key": "your-api-key"
}
```

字段说明：

- `api_url`: OpenAI 兼容的 completion 完整端点，通常需要以 `/v1/chat/completions` 结尾。
- `model`: 默认远程模型。
- `timeout`: 请求超时（秒）。
- `api_key`: 远程 API Key。

### 参数优先级

`completion` 路径的参数优先级：

1. 工具调用参数
2. `ocr_config.json`
3. 内置默认值

示例：

- `ocr_image(image_path="scan.png", engine="completion")`
  使用 `ocr_config.json` 中的 `api_url`、`model`、`timeout`、`api_key`。
- `ocr_image(image_path="scan.png", engine="completion", model="grok-4")`
  `model` 使用显式参数，其他字段仍取自 `ocr_config.json`。

### 默认 OCR Prompt 策略

默认 OCR prompt 较为严格，目标是：

- 仅返回可见文字
- 不输出解释、摘要或回答问题
- 不输出 Markdown 标题、列表或代码块
- 尽量保留行序与换行
- 避免无意义的重复输出

如需特定排版策略，可通过 `prompt` 参数覆盖默认 prompt。

## OCR 依赖与前置条件

### 图像 OCR

- 本地 Tesseract OCR 需要安装 `tesseract`。
- 远程 completion OCR 需要有效的 `ocr_config.json`。

### 扫描 PDF OCR

`ocr_pdf` 或 `ocr_to_docx` 处理扫描件时依赖：

- `pdf2image`
- `Pillow`
- Windows 下通常需要 **Poppler** 以将 PDF 页渲染为图像

推荐安装：

```bash
pip install pdf2image Pillow
```

Windows 用户需单独安装 Poppler 并确保可执行文件在系统路径中，否则扫描 PDF 的 OCR 可能无法完成。

## OCR 示例

### MCP 客户端自然语言示例

- “对 scan.png 进行 OCR，只返回识别文本。”
- “对 scan.pdf 前三页进行 OCR，并导出到 Word。”
- “显示当前 OCR 模型与端点配置。”

### 直接 Python API 示例

#### 单张图片 OCR

```python
from docuflow_mcp.extensions.ocr import OCROperations

result = OCROperations.ocr_image(
    image_path="scan.png",
    engine="completion",
)
```

#### PDF OCR

```python
result = OCROperations.ocr_pdf(
    pdf_path="scan.pdf",
    engine="completion",
    pages=[1, 2, 3],
)
```

#### OCR 转 Word

```python
result = OCROperations.ocr_to_docx(
    source="scan.pdf",
    output_path="scan_ocr.docx",
    engine="completion",
)
```

#### OCR 状态

```python
status = OCROperations.get_status()
```

`ocr_status` 会输出：

- 解析到的 `ocr_config.json` 路径
- 生效的 `api_url`、`model`、`timeout`
- Tesseract 与 completion 可用性
- `claude -> completion` 兼容别名

## 项目结构

```text
DocuFlow/
|-- src/docuflow_mcp/
|   |-- server.py                # MCP 服务入口
|   |-- document.py              # Word 文档操作
|   |-- core/
|   |   |-- registry.py          # 工具注册与分发
|   |   `-- middleware.py        # 中间件
|   |-- extensions/
|   |   |-- excel.py             # Excel 操作
|   |   |-- pdf.py               # PDF 操作
|   |   |-- ppt.py               # PowerPoint 操作
|   |   |-- converter.py         # 格式转换
|   |   |-- ocr.py               # OCR
|   |   |-- image_gen.py         # AI 图像生成
|   |   |-- html_to_pptx.py      # HTML 转 PPTX
|   |   |-- styles.py            # 样式管理
|   |   |-- templates.py         # 模板管理
|   |   `-- validator.py         # 校验与修复
|   `-- utils/
|       |-- deps.py              # 依赖检查
|       `-- paths.py             # 路径校验
|-- tests/                       # 测试
|-- scripts/                     # 维护脚本
|-- install.py                   # 安装脚本
|-- install_codex.py             # Codex 安装脚本
|-- pyproject.toml               # 项目元数据
|-- LICENSE                      # Apache 2.0 许可证
|-- README.md                    # 英文文档
`-- README.zh-CN.md              # 中文文档
```

## 开发与测试

安装开发依赖：

```bash
pip install -e .[dev]
```

运行测试：

```bash
pytest -q
```

仅运行 OCR 测试：

```bash
pytest -q tests/test_ocr.py
```

## 故障排查

### 1. `ocr_image` 能运行但返回解释性内容

请优先检查：

- 所选模型是否严格遵循 OCR 指令
- 是否覆盖了默认 `prompt`
- 上游是否将你的模型名映射到了其他模型

### 2. `ocr_pdf` 缺少依赖

请优先检查：

- `pdf2image` 是否安装
- Windows 是否安装 Poppler
- `ocr_status` 是否显示 `pdf2image` 和 `PIL` 可用

### 3. 远程 OCR 失败或超时

请优先检查：

- `ocr_config.json` 是否使用完整的 `/v1/chat/completions` 端点
- `api_key` 与 `model` 是否有效
- 上游服务是否能稳定处理多模态请求
- 是否存在代理或上游 TLS 问题

## 安全说明

- `ocr_config.json` 默认被 Git 忽略。
- 不要在 README、测试或提交历史中写入真实 API Key。
- 若在 CI 中使用远程 OCR，建议使用受限权限的专用 Key。

## 许可证

Apache License 2.0，详见 `LICENSE`。
