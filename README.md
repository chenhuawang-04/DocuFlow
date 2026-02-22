# DocuFlow MCP

用于 Word 文档处理的 MCP (Model Context Protocol) 服务器。让 Claude Code 等 AI CLI 工具能够创建、读取和修改 Word 文档。

## 功能特性

### 文档操作
- 创建新文档（支持模板）
- 读取文档内容
- 获取文档信息和属性
- 设置文档属性（标题、作者等）
- 合并多个文档
- 获取可用样式列表

### 段落操作
- 添加段落（支持丰富格式）
- 修改段落内容和格式
- 删除段落
- 获取段落详细信息

### 标题操作
- 添加各级标题
- 获取文档大纲结构

### 表格操作
- 创建表格
- 获取表格内容
- 设置单元格内容和格式
- 添加/删除行列
- 合并单元格
- 设置列宽
- 删除表格

### 图片操作
- 插入图片
- 设置图片尺寸和对齐

### 列表操作
- 添加无序列表
- 添加有序列表

### 页面设置
- 设置页边距
- 设置页面大小和方向
- 添加分页符/分节符

### 页眉页脚
- 设置页眉
- 设置页脚
- 添加页码

### 搜索替换
- 查找文本
- 替换文本（支持全部替换）

### 特殊内容
- 添加超链接
- 添加目录
- 添加换行符
- 添加水平线

### 导出功能
- 导出为纯文本
- 导出为 Markdown

## 安装

### 方式一：直接安装依赖

```bash
pip install mcp python-docx Pillow
```

### 方式二：使用 pip 安装包

```bash
cd E:/Project/DocuFlow
pip install -e .
```

## 配置 Claude Code

在 `~/.claude/settings.json` 中添加 MCP 服务器配置：

```json
{
  "mcpServers": {
    "docuflow": {
      "command": "python",
      "args": ["E:/Project/DocuFlow/src/docuflow_mcp/server.py"]
    }
  }
}
```

或者如果已安装包：

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

配置完成后，重启 Claude Code，即可使用自然语言操作 Word 文档：

### 创建文档

> "创建一个新文档 report.docx，标题为'月度销售报告'"

### 添加内容

> "在 report.docx 中添加一个二级标题'销售概况'，然后添加一段描述文字"

### 添加表格

> "在文档中添加一个 4 行 3 列的表格，包含产品名称、销量和金额三列"

### 设置格式

> "将第一段文字设置为微软雅黑字体，14 号，加粗，居中对齐"

### 搜索替换

> "将文档中所有的 '2024' 替换为 '2025'"

### 导出

> "将 report.docx 导出为 Markdown 格式"

## 工具列表

| 工具名 | 功能 |
|--------|------|
| `doc_create` | 创建新文档 |
| `doc_read` | 读取文档内容 |
| `doc_info` | 获取文档信息 |
| `doc_set_properties` | 设置文档属性 |
| `doc_merge` | 合并文档 |
| `doc_get_styles` | 获取样式列表 |
| `paragraph_add` | 添加段落 |
| `paragraph_modify` | 修改段落 |
| `paragraph_delete` | 删除段落 |
| `paragraph_get` | 获取段落信息 |
| `heading_add` | 添加标题 |
| `heading_get_outline` | 获取大纲 |
| `table_add` | 添加表格 |
| `table_get` | 获取表格内容 |
| `table_set_cell` | 设置单元格 |
| `table_add_row` | 添加行 |
| `table_add_column` | 添加列 |
| `table_delete_row` | 删除行 |
| `table_merge_cells` | 合并单元格 |
| `table_set_column_width` | 设置列宽 |
| `table_delete` | 删除表格 |
| `image_add` | 插入图片 |
| `image_add_to_paragraph` | 在段落中插入图片 |
| `list_add_bullet` | 添加无序列表 |
| `list_add_numbered` | 添加有序列表 |
| `page_set_margins` | 设置页边距 |
| `page_set_size` | 设置页面大小 |
| `page_add_break` | 添加分页符 |
| `page_add_section_break` | 添加分节符 |
| `header_set` | 设置页眉 |
| `footer_set` | 设置页脚 |
| `page_number_add` | 添加页码 |
| `search_find` | 查找文本 |
| `search_replace` | 替换文本 |
| `hyperlink_add` | 添加超链接 |
| `toc_add` | 添加目录 |
| `line_break_add` | 添加换行符 |
| `horizontal_line_add` | 添加水平线 |
| `export_to_text` | 导出为文本 |
| `export_to_markdown` | 导出为 Markdown |

## 格式参数说明

### 尺寸单位
- `pt` - 点 (默认)
- `in` / `inch` / `inches` - 英寸
- `cm` - 厘米
- `mm` - 毫米

示例：`"12pt"`, `"1in"`, `"2.54cm"`

### 颜色格式
- Hex: `"#FF0000"`, `"#F00"`
- RGB: `"rgb(255, 0, 0)"`
- 预定义: `"red"`, `"blue"`, `"green"`, `"black"`, `"white"`, `"yellow"`, `"orange"`, `"purple"`, `"gray"`, `"pink"`, `"brown"`, `"cyan"`, `"magenta"`

### 对齐方式
- `"left"` - 左对齐
- `"center"` - 居中
- `"right"` - 右对齐
- `"justify"` - 两端对齐

## 项目结构

```
DocuFlow/
├── src/
│   └── docuflow_mcp/
│       ├── __init__.py      # 包初始化
│       ├── server.py        # MCP 服务器入口
│       ├── tools.py         # 工具定义
│       └── document.py      # 文档操作核心逻辑
├── pyproject.toml           # 项目配置
└── README.md               # 使用说明
```

## 注意事项

1. **文件格式**: 仅支持 `.docx` 格式（Office Open XML），不支持旧版 `.doc` 格式
2. **目录更新**: 添加目录后需要在 Word 中按 F9 更新
3. **中文字体**: 设置字体时会同时设置中文字体（eastAsia）
4. **路径格式**: 建议使用绝对路径

## 许可证

MIT License
