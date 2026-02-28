# DocuFlow MCP

All-in-One 文档处理 MCP (Model Context Protocol) 服务器，提供 **149 个工具**，覆盖 Word / Excel / PowerPoint / PDF / 格式转换 / OCR / AI 图片生成。

## 功能模块

| 模块 | 工具数 | 说明 |
|------|--------|------|
| Word (.docx) | 49 | 文档/段落/表格/图片/列表/页面/页眉页脚/搜索替换/样式/模板/批注/导出 |
| Excel (.xlsx) | 33 | 工作簿/工作表/单元格/公式/数据处理/统计/图表/透视表 |
| PowerPoint (.pptx) | 30 | 幻灯片/形状/图表/动画/切换/母版/占位符 |
| PDF | 23 | 提取/合并/拆分/旋转/水印/加密/表单/转换 |
| 格式转换 | 4 | 40+ 格式互转 (pandoc) |
| OCR | 4 | 图片/PDF 文字识别 (Tesseract/Claude) |
| HTML→PPTX | 3 | HTML 幻灯片转 PowerPoint |
| AI 图片生成 | 3 | 文本描述生成图片 |

## 安装

### 方式一：一键安装（推荐）

```bash
cd DocuFlow
python install.py
```

### 方式二：手动安装

```bash
pip install -e .
```

然后在 `~/.claude/settings.json` 中添加：

```json
{
  "mcpServers": {
    "docuflow": {
      "command": "docuflow-mcp"
    }
  }
}
```

## 使用示例

配置完成后重启 Claude Code，即可用自然语言操作文档：

- "创建一个新文档 report.docx，标题为'月度销售报告'"
- "在 report.docx 中添加表格和图表"
- "创建 Excel 文件，写入数据并生成柱状图"
- "创建 PPT 演示文稿，添加 5 张幻灯片"
- "提取 PDF 中的表格并转为 Excel"
- "将 report.docx 转换为 PDF"

## 项目结构

```
DocuFlow/
├── src/docuflow_mcp/
│   ├── __init__.py          # 包初始化
│   ├── server.py            # MCP 服务器入口
│   ├── document.py          # Word 文档核心逻辑
│   ├── core/
│   │   ├── registry.py      # 工具注册框架
│   │   └── middleware.py     # 中间件管理
│   ├── extensions/
│   │   ├── excel.py         # Excel 操作
│   │   ├── pdf.py           # PDF 操作
│   │   ├── ppt.py           # PowerPoint 操作
│   │   ├── converter.py     # 格式转换
│   │   ├── ocr.py           # OCR 文字识别
│   │   ├── image_gen.py     # AI 图片生成
│   │   ├── html_to_pptx.py  # HTML→PPTX
│   │   ├── styles.py        # 样式管理
│   │   ├── templates.py     # 模板管理
│   │   ├── validator.py     # 格式校验
│   │   └── advanced.py      # 高级文档分析
│   └── utils/
│       ├── deps.py          # 依赖检查
│       └── paths.py         # 路径校验
├── install.py               # Claude Code 安装脚本
├── install_codex.py         # Codex 安装脚本
├── tests/                   # 测试文件
├── pyproject.toml           # 项目配置
└── CLAUDE.md                # AI 代理指令
```

## 格式参数

- **尺寸**: `"12pt"`, `"1in"`, `"2.54cm"`, `"25.4mm"`
- **颜色**: `"#FF0000"`, `"rgb(255,0,0)"`, `"red"`
- **对齐**: `"left"`, `"center"`, `"right"`, `"justify"`

## 许可证

MIT License
