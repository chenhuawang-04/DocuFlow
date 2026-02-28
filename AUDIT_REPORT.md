# DocuFlow 项目全面审查报告

> **日期**: 2026-02-27
> **范围**: src/docuflow_mcp/ 全部源码、扩展模块、安装脚本、Skill 文件、测试、项目配置
> **审查方式**: 4 个并行探索代理 + 1 个代码采集代理 + 3 个交叉验证代理
> **修订**: v3 — 经三次审查修正计数一致性和技术细节

---

## 目录

1. [审查总览](#1-审查总览)
2. [严重问题 (CRITICAL)](#2-严重问题-critical)
3. [高优先级问题 (HIGH)](#3-高优先级问题-high)
4. [中优先级问题 (MEDIUM)](#4-中优先级问题-medium)
5. [低优先级问题 (LOW)](#5-低优先级问题-low)
6. [已修复问题确认](#6-已修复问题确认)
7. [修复优先级建议](#7-修复优先级建议)

---

## 1. 审查总览

| 严重性 | 数量 | 主要类别 |
|--------|------|----------|
| **CRITICAL** | 3 | 资源泄漏(31函数/32调用点)、目录创建竞态(23处)、PDF 异常吞没(7处) |
| **HIGH** | 10 | API Key 泄漏、路径校验绕过、表单假成功、text_replace 不生效、异常过宽、测试质量、gitignore |
| **MEDIUM** | 8 | 页码校验、度量解析、文档过时、测试硬编码、命名遮蔽 |
| **LOW** | 8 | 命名规范、魔法数字、导入组织、测试覆盖率 |

---

## 2. 严重问题 (CRITICAL)

### C1. Excel 工作簿资源泄漏（31 个函数，32 个调用点）

**文件**: `src/docuflow_mcp/extensions/excel.py`
**风险**: 文件锁未释放、内存泄漏、Windows 上无法删除/移动文件

**问题描述**: `load_workbook()` 成功后，若后续操作抛出异常或提前 return，工作簿永远不会被关闭。全部 31 个使用 `load_workbook()` 的函数（共 32 个调用点，`stats_summary` 有两次）均未使用 `try-finally` 或上下文管理器，**无一例外存在泄漏路径**。

**A. 提前 return 泄漏（29 个函数）** — 参数校验失败时直接 return，跳过 `wb.close()`：

| # | 函数 | load_workbook 行 | 泄漏 return 行 |
|---|------|-----------------|---------------|
| 1 | `read()` | 134 | 139, 151 |
| 2 | `save_as()` | 296 | 308 |
| 3 | `add_sheet()` | 373 | 376 |
| 4 | `delete_sheet()` | 407 | 410, 413 |
| 5 | `rename_sheet()` | 445 | 448, 451 |
| 6 | `copy_sheet()` | 485 | 488, 491 |
| 7 | `read_cell()` | 540 | 545 |
| 8 | `write_cell()` | 612 | 617, 658 |
| 9 | `format_cell()` | 706 | 711 |
| 10 | `merge_cell()` | 821 | 826 |
| 11 | `set_formula()` | 877 | 882 |
| 12 | `insert_row()` | 927 | 932 |
| 13 | `delete_row()` | 971 | 976 |
| 14 | `insert_col()` | 1015 | 1020 |
| 15 | `delete_col()` | 1061 | 1066 |
| 16 | `to_word()` | 1126 | 1130 |
| 17 | `formula_batch()` | 1267 | 1272 |
| 18 | `formula_quick()` | 1368 | 1373 |
| 19 | `data_sort()` | 1430 | 1435, 1452 |
| 20 | `data_filter()` | 1532 | 1537 |
| 21 | `data_validate()` | 1600 | 1605, 1614, 1629, 1632 |
| 22 | `data_deduplicate()` | 1684 | 1689, 1704 |
| 23 | `data_fill()` | 1804 | 1809, 1851 |
| 24 | `stats_summary()` | 1902 (+1963) | 1907, 1925 |
| 25 | `conditional_format()` | 2032 | 2037, 2107 |
| 26 | `named_range()` | 2154 | 2158, 2163 |
| 27 | `chart_create()` | 2271 | 2276, 2296 |
| 28 | `chart_modify()` | 2399 | 2404, 2413 |
| 29 | `pivot_create()` | 2496 | 2501, 2518, 2523, 2548 |

**B. 异常路径泄漏（2 个函数）** — 无提前 return，但外层 `except Exception` 跳过 `wb.close()`：

| # | 函数 | load_workbook 行 | wb.close() 行 | 泄漏原因 |
|---|------|-----------------|--------------|---------|
| 30 | `get_info()` | 223 | 246 | 223-246 间任意异常被外层 `except`(line 258) 捕获，跳过 close |
| 31 | `list_sheets()` | 340 | 343 | 340-343 间异常被外层 `except`(line 353) 捕获，跳过 close |

> 注：`stats_summary()` 在 line 1963 有第二次 `load_workbook()`（写入模式），故全文件共 32 个调用点。

**典型泄漏示例** — `read()` (line 134-151):
```python
# excel.py:134
wb = load_workbook(path, data_only=True)

# line 139 — 提前 return，wb 未 close
if sheet not in wb.sheetnames:
    return {"success": False, "error": f"工作表不存在: {sheet}"}  # ← wb 泄漏

# line 151 — 另一个提前 return
try:
    cells = ws[range]
except Exception:
    return {"success": False, "error": f"无效的范围: {range}"}    # ← wb 泄漏

# line 189 — 唯一的 close 调用（仅成功路径）
wb.close()
```

**典型泄漏示例** — `format_cell()` (line 706-791):
```python
# excel.py:706
wb = load_workbook(path)

# line 711 — 提前 return
if sheet not in wb.sheetnames:
    return {"success": False, "error": f"工作表不存在: {sheet}"}  # ← wb 泄漏

# line 790-791 — 仅成功路径才 close
wb.save(path)
wb.close()
```

**修复方案**:
```python
def excel_read(path, ...):
    wb = load_workbook(path, data_only=True)
    try:
        # ... 所有操作逻辑 ...
        return {"success": True, ...}
    except (KeyError, ValueError) as e:
        return {"success": False, "error": str(e)}
    finally:
        wb.close()
```

---

### C2. `os.makedirs()` 缺少 `exist_ok=True`（23 处）

**风险**: 并发请求或目录已存在时 `FileExistsError` 导致崩溃
**根因**: `os.path.exists()` + `os.makedirs()` 存在 TOCTOU 竞态条件

> 注：所有 `Path.mkdir()` 调用（middleware.py:105, templates.py:355, image_gen.py:299, converter.py:196）及部分 `os.makedirs` 调用（document.py:147,318, config.py:219, templates.py:93,407, pdf.py:265,485）均已正确使用 `exist_ok=True`，无问题。

**完整位置清单**:

| # | 文件 | 行号 | 函数/上下文 |
|---|------|------|------------|
| 1 | `html_to_pptx.py` | 478 | `convert()` |
| 2 | `html_to_pptx.py` | 522 | `convert_multi()` |
| 3 | `ppt.py` | 95 | `ppt_create()` |
| 4 | `ppt.py` | 407 | `ppt_merge()` |
| 5 | `excel.py` | 74 | `excel_create()` |
| 6 | `excel.py` | 294 | `excel_save_as()` |
| 7 | `excel.py` | 1160 | `excel_to_word()` |
| 8 | `ocr.py` | 502 | `ocr_to_docx()` |
| 9 | `pdf.py` | 423 | `pdf_merge()` |
| 10 | `pdf.py` | 572 | `pdf_extract_pages()` |
| 11 | `pdf.py` | 641 | `pdf_rotate()` |
| 12 | `pdf.py` | 708 | `pdf_delete_pages()` |
| 13 | `pdf.py` | 808 | `pdf_add_watermark()` |
| 14 | `pdf.py` | 863 | `pdf_tables_to_word()` |
| 15 | `pdf.py` | 966 | `pdf_tables_to_excel()` |
| 16 | `pdf.py` | 1069 | `pdf_to_text()` |
| 17 | `pdf.py` | 1240 | `pdf_to_editable()` |
| 18 | `pdf.py` | 1472 | `pdf_text_replace()` |
| 19 | `pdf.py` | 1639 | `pdf_redact()` |
| 20 | `pdf.py` | 1753 | `pdf_annotate_text()` |
| 21 | `pdf.py` | 1831 | `pdf_encrypt()` |
| 22 | `pdf.py` | 1900 | `pdf_decrypt()` |
| 23 | `pdf.py` | 2070 | `pdf_form_fill()` |

> 按文件分布：pdf.py 15 处、excel.py 3 处、html_to_pptx.py 2 处、ppt.py 2 处、ocr.py 1 处。

**典型模式** (所有 23 处相同):
```python
output_dir = os.path.dirname(output_path)
if output_dir and not os.path.exists(output_dir):
    os.makedirs(output_dir)         # ← 竞态：检查与创建之间目录可能被其他进程创建
```

**修复方案** (统一替换):
```python
# 替换所有 3 行模式为 1 行：
os.makedirs(os.path.dirname(path), exist_ok=True)
```

---

### C3. PDF 模块 7 处 `except Exception: pass`

**文件**: `src/docuflow_mcp/extensions/pdf.py`
**风险**: 静默吞掉 PDF 损坏、权限错误、内存不足等关键异常，用户无法知道操作是否真正成功

**位置 1-2** — `pdf_get_outline` (line 365, 373):
```python
# pdf.py:363-374
                    try:
                        page_num = reader.get_destination_page_number(item) + 1
                    except Exception:           # ← C3a: 吞没页码解析异常
                        pass

                    outline_items.append({...})
                except Exception:               # ← C3b: 吞没整个 outline 条目异常
                    pass
```

**位置 3** — `pdf_add_watermark` (line 797):
```python
# pdf.py:795-799
                writer.add_annotation(page_number=i, annotation=annotation)
                watermarked_count += 1
            except Exception:                   # ← C3c: 水印添加失败被静默忽略
                pass
```

**位置 4** — `pdf_tables_to_word` (line 906):
```python
# pdf.py:904-907
            try:
                table.style = table_style
            except Exception:                   # ← C3d: 表格样式应用失败被忽略
                table.style = 'Table Grid'
```

**位置 5** — `pdf_text_replace` (line 1464):
```python
# pdf.py:1463-1465
            except Exception:                   # ← C3e: 内容流操作失败被忽略
                pass
```

**位置 6-7** — `pdf_form_fill` (line 2056, 2064):
```python
# pdf.py:2054-2065
        try:
            writer.update_page_form_field_values(None, fields, auto_regenerate=True)
        except Exception:                       # ← C3f: 表单填充失败
            for page_num in range(len(writer.pages)):
                try:
                    writer.update_page_form_field_values(writer.pages[page_num], fields)
                except Exception:               # ← C3g: 逐页填充也失败被忽略
                    pass
```

**修复方案**:
```python
# C3a: 页码解析
except (ValueError, KeyError, AttributeError):
    page_num = -1  # 标记无法解析

# C3c: 水印添加
except (ValueError, TypeError) as e:
    logger.warning(f"Failed to add watermark to page {i}: {e}")

# C3d: 表格样式
except (KeyError, ValueError):
    table.style = 'Table Grid'

# C3f-g: 表单填充
except (ValueError, AttributeError, TypeError) as e:
    logger.warning(f"Form fill failed: {e}")
```

---

## 3. 高优先级问题 (HIGH)

### H1. API Key 泄漏风险 — HTTP 重定向跟随

**文件**: `src/docuflow_mcp/extensions/image_gen.py:108-122`
**风险**: 重定向到恶意域时，Authorization Bearer token 随请求发送到第三方服务器

```python
# image_gen.py:108-122
try:
    with opener.open(request, timeout=timeout) as response:
        raw = response.read()
except HTTPError as exc:
    if exc.code in (301, 302, 303, 307, 308) and redirects_remaining > 0:
        location = exc.headers.get("Location")
        if location:
            redirected_url = urljoin(api_url, location)
            return _request_chat_completion(
                redirected_url,             # ← 可能是完全不同的域
                api_key,                    # ← API Key 被发送到重定向目标！
                payload, timeout, ssl_context, redirects_remaining - 1
            )
```

**修复方案**:
```python
from urllib.parse import urlparse

redirected_url = urljoin(api_url, location)
# 检查重定向目标域名是否一致
if urlparse(redirected_url).netloc != urlparse(api_url).netloc:
    raise RuntimeError(f"Cross-origin redirect blocked: {redirected_url}")
```

---

### H2. 路径校验绕过 — `startswith('<')` 跳过所有参数

**文件**: `src/docuflow_mcp/core/registry.py:99-117`
**风险**: 任何以 `<` 开头的路径参数可以绕过 `validate_path()` 校验

```python
# registry.py:99-117
_PATH_PARAMS = {
    'path', 'output_path', 'ppt_path', 'input_path',
    'template', 'path1', 'path2', 'html_source',
    'source', 'target', 'reference_doc', 'css',
}
try:
    for param_name in _PATH_PARAMS:
        if param_name in args and isinstance(args[param_name], str):
            val = args[param_name]
            if val.strip().startswith('<'):      # ← 所有参数都会被跳过，不仅是 html_source
                continue
            args[param_name] = validate_path(val)
```

**问题**: 此跳过逻辑是为 `html_source` 参数设计的（可能传入 HTML 内容），但应用于所有路径参数。攻击者可以在 `path` 参数中传入 `<script>...` 来绕过校验。

**修复方案**:
```python
_HTML_CONTENT_PARAMS = {'html_source'}  # 仅这些参数可以是 HTML 内容

for param_name in _PATH_PARAMS:
    if param_name in args and isinstance(args[param_name], str):
        if param_name in _HTML_CONTENT_PARAMS:
            continue  # HTML 内容参数跳过路径校验
        args[param_name] = validate_path(args[param_name])
```

---

### H3. converter `extra_args` 未过滤

**文件**: `src/docuflow_mcp/extensions/converter.py:58-66`
**风险**: 用户可通过 `extra_args` 传入任意 pandoc 参数（如 `--output=/etc/passwd`）

```python
# converter.py:58-66
@staticmethod
def _run_pandoc(args: List[str]) -> Dict[str, Any]:
    try:
        result = subprocess.run(
            ['pandoc'] + args,              # ← args 直接拼接，包含用户提供的 extra_args
            capture_output=True,
            text=True,
            timeout=300
        )
```

**修复方案**:
```python
# 白名单允许的 pandoc 参数
_ALLOWED_PANDOC_FLAGS = {'--toc', '--standalone', '--wrap=auto', '--columns=80', ...}

def _sanitize_extra_args(extra_args: List[str]) -> List[str]:
    safe = []
    for arg in extra_args:
        if arg in _ALLOWED_PANDOC_FLAGS or arg.startswith('--metadata='):
            safe.append(arg)
        else:
            logger.warning(f"Blocked pandoc argument: {arg}")
    return safe
```

---

### H4. `pdf_form_fill` 失败后仍返回成功

**文件**: `src/docuflow_mcp/extensions/pdf.py:2054-2082`
**风险**: 表单填充完全失败时，用户收到 `"success": True`，误以为填写成功

```python
# pdf.py:2054-2065 — 双层 except 静默吞没所有异常
try:
    writer.update_page_form_field_values(None, fields, auto_regenerate=True)
except Exception:                           # ← 第一次失败
    for page_num in range(len(writer.pages)):
        try:
            writer.update_page_form_field_values(writer.pages[page_num], fields)
        except Exception:                   # ← 逐页也全部失败
            pass                            # ← 但无任何记录

# pdf.py:2075-2082 — 无论上面是否成功，都返回 success: True
return {
    "success": True,                        # ← 即使所有填充都失败
    "path": path,
    "output_path": output_path,
    "fields_filled": list(fields.keys()),   # ← 误导：声称已填充
    "message": f"已填写 {len(fields)} 个字段到 {output_path}"
}
```

**修复方案**: 记录填充结果，根据实际情况返回：
```python
filled_ok = False
try:
    writer.update_page_form_field_values(None, fields, auto_regenerate=True)
    filled_ok = True
except (ValueError, AttributeError, TypeError):
    for page_num in range(len(writer.pages)):
        try:
            writer.update_page_form_field_values(writer.pages[page_num], fields)
            filled_ok = True
        except (ValueError, AttributeError, TypeError):
            pass

if not filled_ok:
    return {"success": False, "error": "表单填充失败，字段可能不存在或不可写"}
```

---

### H5. `pdf_text_replace` 计算替换但未写回 — 功能性 Bug

**文件**: `src/docuflow_mcp/extensions/pdf.py:1454-1467`
**风险**: 函数报告替换成功并返回替换计数，但输出 PDF **未做任何改动**

```python
# pdf.py:1454-1467
for page in reader.pages:
    try:
        if '/Contents' in page:
            contents = page['/Contents']
            if hasattr(contents, 'get_data'):
                data = contents.get_data().decode('latin-1')
                if old_text in data:
                    new_data = data.replace(old_text, new_text)     # ← line 1460: 计算了替换结果
                    # 注意：直接修改内容流可能导致问题                    # ← 注释承认有问题
                    # 这是简化实现                                      # ← 但未实现写回！
                    replacement_count += data.count(old_text)        # ← line 1463: 计数增加
    except Exception:
        pass
    writer.add_page(page)                                           # ← line 1467: 添加的是原始 page，不是修改后的

# pdf.py:1475-1481 — 返回"成功"
return {
    "success": True,
    "replacements": replacement_count,       # ← 声称替换了 N 处，实际为 0
    "message": f"已替换 {replacement_count} 处..."
}
```

**根因**: `new_data` 在 line 1460 计算后从未写回页面内容流。`writer.add_page(page)` 添加的是未修改的原始页面。

**修复方案**: 此为已知的 PyPDF 限制。建议：
1. 方案 A：使用 `pypdf` 的 `ContentStream` API 写回修改后的数据
2. 方案 B：文档中明确标注此工具为"查找并计数"而非"查找并替换"
3. 方案 C：对于简单文本替换，使用覆盖绘制方式（当前 `pdf_redact` + `pdf_annotate_text` 组合）

---

### H6. Excel `except Exception:` 异常过宽

**文件**: `src/docuflow_mcp/extensions/excel.py:150`
**风险**: 掩盖真实错误（如磁盘满、权限拒绝），用户看到的是误导性错误信息

```python
# excel.py:150-151
try:
    cells = ws[range]
except Exception:                           # ← 捕获所有异常，包括 MemoryError, SystemExit 等
    return {"success": False, "error": f"无效的范围: {range}"}
```

**修复方案**:
```python
except (KeyError, ValueError, IndexError) as e:
    return {"success": False, "error": f"无效的范围: {range} ({e})"}
```

---

### H7. 测试断言缺失 — 17 个 test_*.py 文件，0 个 assert

**文件**: 全部 17 个 `test_*.py` 文件（另有 `verify_format.py` 验证脚本同样无 assert）
**风险**: 测试永远"通过"，CI 永远不会因回归而失败

静态扫描结果：17 个测试文件中 `assert` 和 `self.assert` 出现次数 = **0**。所有检查均以 `print()` + `if/else` 形式存在：

```python
# test_document.py:245 — 典型问题
result = DocumentOperations.get_info(doc_path)
print(f"    段落数: {result['statistics']['paragraph_count']}")
# ← 没有 assert，即使 paragraph_count 是 0 也"通过"

# test_excel.py — 类似
if result.get('success'):
    print(f"   [OK] 创建文档成功")
else:
    print(f"   [FAIL] 创建文档失败")
# ← 即使 FAIL 也不会让测试退出失败
```

**修复方案**: 转为 pytest 风格，所有 print 检查替换为 assert：
```python
result = DocumentOperations.get_info(doc_path)
assert result['success'], f"get_info failed: {result.get('error')}"
assert result['statistics']['paragraph_count'] > 0
```

---

### H8. .gitignore 不完整

**文件**: `.gitignore`
**当前缺失**:

| 文件/目录 | 状态 | 原因 |
|-----------|------|------|
| `.venv-pack/` | 未忽略 | 虚拟环境打包，可能数百 MB |
| `*.zip` | 未忽略 | 3 个源码压缩包 (28MB+) |
| `SESSION_SUMMARY.md` | 未忽略 | 会话生成文档 |
| `CONTINUE_GUIDE.md` | 未忽略 | 会话生成文档 |
| `PROJECT_INDEX.md` | 未忽略 | 会话生成文档 |
| `ppt_agent/` | 未忽略 | 已废弃 (agent→skill 迁移) |
| `AUDIT_REPORT.md` | 未忽略 | 审查报告 |
| `nul` | 已忽略但仍存在 | Windows 重定向产物 |

**修复方案**: 追加到 `.gitignore`:
```gitignore
# Virtual environment packs
.venv-pack/

# Archives
*.zip
*.tar.gz

# Generated session docs
SESSION_SUMMARY.md
CONTINUE_GUIDE.md
PROJECT_INDEX.md
AUDIT_REPORT.md

# Obsolete directories
ppt_agent/
```

---

### H9. styles.py 模块级导入无保护

**文件**: `src/docuflow_mcp/extensions/styles.py:10-11`、`src/docuflow_mcp/extensions/__init__.py:5`
**风险**: `python-docx` 未安装时，`styles.py` 的 `ImportError` 会沿 `__init__.py` 的导入链向上传播，导致整个 extensions 包加载失败，进而使 MCP 服务器启动崩溃

```python
# styles.py:10-11
from docx import Document                   # ← 模块级导入，无 check_import 保护
from docx.enum.style import WD_STYLE_TYPE   # ← 同上
```

对比其他扩展模块的做法（正确的模式）：
```python
# excel.py — 每个函数内检查
from ..utils.deps import check_import
def excel_create(...):
    if not check_import("openpyxl"):
        return {"success": False, "error": "openpyxl is required"}
```

**修复方案**: 将 `from docx import ...` 移入函数内部，或在文件顶部添加：
```python
from ..utils.deps import check_import

# 替换模块级 import
Document = None
WD_STYLE_TYPE = None

def _ensure_docx():
    global Document, WD_STYLE_TYPE
    if Document is None:
        from docx import Document as _Doc
        from docx.enum.style import WD_STYLE_TYPE as _WST
        Document, WD_STYLE_TYPE = _Doc, _WST
```

---

### H10. 测试依赖外部工具无跳过机制

**文件**: `test_converter.py:58`, `test_ocr.py:238`
**风险**: 未安装 pandoc/tesseract 时测试报错而非跳过

```python
# test_converter.py:58
print(f"Pandoc可用: {result.get('pandoc_available', False)}")
# ← 没有 @pytest.mark.skipif(not shutil.which('pandoc'))

# test_ocr.py:238
if not engines.get('tesseract', {}).get('available'):
    print("  Windows: https://github.com/UB-Mannheim/tesseract/wiki")
# ← 没有跳过逻辑
```

---

## 4. 中优先级问题 (MEDIUM)

### M1. PDF 页码校验不足

**文件**: `src/docuflow_mcp/extensions/pdf.py`
**函数**: `pdf_extract_pages`(line 550), `pdf_rotate`(line 609), `pdf_delete_pages`(line 676)
**库**: 均使用 **pypdf**（`from pypdf import PdfWriter, PdfReader`），非 pdfplumber
**问题**: 非法页码（负数、0、超范围）被 `1 <= p <= total_pages` 条件静默过滤，不向用户报错。其中 `pdf_rotate()` 甚至在全部页码无效时仍返回成功（旋转 0 页）。

**修复方案**:
```python
# 在校验条件前，先检查非法页码并报告
invalid = [p for p in pages if p < 1 or p > total_pages]
if invalid:
    return {"success": False, "error": f"页码超出范围(1-{total_pages}): {invalid}"}
```

---

### M2. 测试文件硬编码绝对路径

**文件**: `test_day10_11.py`, `test_day3.py`, `test_day6_7.py`, `test_day8_9.py`, `test_document.py`, `test_performance_benchmark.py`

```python
# test_day10_11.py:70
doc1_path = "E:/Project/DocuFlow/test_compare1.docx"    # ← 硬编码绝对路径

# test_performance_benchmark.py:26
"E:/Project/DocuFlow/test_performance"                   # ← 同上
```

**修复方案**:
```python
TEST_DIR = Path(__file__).parent
doc1_path = str(TEST_DIR / "test_compare1.docx")
```

---

### M3. validator 度量单位解析脆弱

**文件**: `src/docuflow_mcp/extensions/validator.py:44,55,76,86`
**问题**: 仅处理 `cm` 单位，其他单位（`in`, `pt`, `px`）会导致 ValueError

```python
# validator.py:44
tolerance = Cm(float(tolerance_str.replace('cm', '')))
# '1in' → Cm(float('1'))  = 错误值！
# '12pt' → ValueError!
```

**修复方案**:
```python
import re

def parse_measurement(value_str: str) -> float:
    """解析度量值，支持 cm/in/pt/px"""
    m = re.match(r'([\d.]+)\s*(cm|in|pt|px)?', value_str.strip())
    if not m:
        raise ValueError(f"无法解析度量值: {value_str}")
    val, unit = float(m.group(1)), m.group(2) or 'cm'
    converters = {'cm': Cm, 'in': Inches, 'pt': Pt, 'px': Emu}
    return converters[unit](val)
```

---

### M4. 文档工具数不一致

| 文件 | 声称的工具数 |
|------|------------|
| `CLAUDE.md` | 149 |
| `SESSION_SUMMARY.md` | 170 |
| `CONTINUE_GUIDE.md` | 134 |
| 实际 `get_all_tools()` | **149** |

需要统一所有文档中的工具数为实际值 **149**。

---

### M5. README.md 项目结构过时

**文件**: `README.md:202-214`

```markdown
## 当前显示
DocuFlow/
├── src/
│   └── docuflow_mcp/
│       ├── __init__.py
│       ├── server.py
│       ├── tools.py
│       └── document.py          ← 仅列 4 个文件
```

实际结构有 25+ 个源文件：`core/`, `extensions/`, `utils/` 目录。

---

### M6. advanced.py 相似度阈值硬编码

**文件**: `src/docuflow_mcp/extensions/advanced.py:78-79`

```python
# advanced.py:78-79
similarity = difflib.SequenceMatcher(None, rem, add).ratio()
if similarity > 0.6:               # ← 不可配置
```

**修复方案**: 添加 `similarity_threshold` 参数，默认 0.6。

---

### M7. Excel `range` 参数遮蔽内建函数

**文件**: `src/docuflow_mcp/extensions/excel.py`
**影响**: 15+ 个函数使用 `range` 作为参数名

```python
# excel.py:109-112
def read(path: str,
         sheet: Optional[str] = None,
         range: Optional[str] = None,       # ← 遮蔽 Python 内建 range()
         include_formatting: bool = False):
```

**修复方案**: 重命名为 `cell_range` 或 `data_range`。
> 注意：此修改会影响已注册的工具参数名，需同步更新 `register_tool` 装饰器和 CLAUDE.md。

---

### M8. OCR bare except

**文件**: `src/docuflow_mcp/extensions/ocr.py:269`

```python
# ocr.py:266-270
if result.confidence < 0.6 and check_import("anthropic"):
    try:
        result = OCROperations._ocr_with_claude(image_path, api_key, prompt)
    except Exception:                       # ← 静默忽略 Claude API 错误
        pass
```

**修复方案**: `except (anthropic.APIError, ConnectionError, TimeoutError):`

---

## 5. 低优先级问题 (LOW)

### L1. 字符串格式化风格不一致

**范围**: 全项目
**问题**: 混用 f-string、`.format()`、`%` 三种风格
**建议**: 统一使用 f-string

### L2. 魔法数字

**文件**: `src/docuflow_mcp/core/middleware.py:189,225-230`
```python
slow_threshold: float = 1.0      # 什么单位？秒？
self.max_param_length = 200      # 为什么是 200？
```

### L3. image_gen 模型名硬编码

**文件**: `src/docuflow_mcp/extensions/image_gen.py:32`
```python
DEFAULT_MODEL = "gpt-4o-mini"    # API 废弃后工具直接失效
```

### L4. 测试覆盖率不足

| 模块 | 工具数 | 测试状态 |
|------|--------|----------|
| Image Generation | 3 | **完全缺失** |
| Styles | 6 | 部分 |
| Batch Operations | 4 | 部分 |
| Validation | 4 | 部分 |

### L5. pytest 配置不完善

**文件**: `pyproject.toml:89-91`
```toml
[tool.pytest.ini_options]
asyncio_mode = "auto"
testpaths = ["."]               # ← 会扫描 .venv-pack/ 等无关目录
```

### L6. 返回值格式不一致

**问题**: 部分函数成功时用 `"message"` key，部分用 `"data"` key，无统一 schema。

### L7. html_to_pptx 文字宽度估算不准

**文件**: `src/docuflow_mcp/extensions/html_to_pptx.py:442`
```python
sum(1.0 if ord(c) > 0x2000 else 0.6 for c in text)  # ← 不考虑字体差异
```

### L8. document.py 样式应用静默失败

**文件**: `src/docuflow_mcp/document.py:425-428`
```python
try:
    para.style = style
except KeyError:
    pass                        # ← 用户不知道样式是否生效
```

---

## 6. 已修复问题确认

以下问题在之前的审计中已修复，本次确认状态：

| Commit | 修复内容 | 状态 |
|--------|----------|------|
| `2308c24` | 共享依赖检查 `utils/deps.py` + 路径校验 `utils/paths.py` | ✅ 已验证 |
| `eb234c5` | 18 处裸 `except:` → 具体异常类型 | ✅ 已验证 |
| `45680c7` | 移除 `shell=True`，使用 `shlex.split` | ✅ 已验证 |
| `f0fba72` | 权限合并、卸载补全、动态步骤数、Skill 内容同步 | ✅ 已验证 |

**安装脚本 (install.py, install_codex.py)**: 全部检查通过，0 问题。
**6 个 Skill 文件**: 内容正确、工具名正确、流程完整，0 问题。
**settings.local.json**: 149 个工具权限正确，0 问题。

---

## 7. 修复优先级建议

### 第一批：安全与稳定性（预计 6-8 小时）

| 编号 | 问题 | 估时 |
|------|------|------|
| C1 | Excel 工作簿 try-finally 包装 (31 函数) | 3h |
| C2 | `os.makedirs` 加 `exist_ok=True` (23 处) | 1h |
| C3 | PDF bare except → 具体异常 (7 处) | 1h |
| H1 | image_gen 重定向域名校验 | 30min |
| H2 | 路径校验仅对 `html_source` 跳过 | 15min |
| H3 | converter extra_args 白名单 | 30min |
| H4 | pdf_form_fill 失败时返回错误 | 30min |
| H5 | pdf_text_replace 标注或修复 | 1h |

### 第二批：健壮性（预计 3-4 小时）

| 编号 | 问题 | 估时 |
|------|------|------|
| H6 | Excel `except Exception` → 具体类型 | 1h |
| H9 | styles.py 导入保护 | 30min |
| M1 | PDF 页码校验 | 30min |
| M3 | validator 度量解析 | 1h |
| M8 | OCR bare except | 15min |

### 第三批：项目治理（预计 2-3 小时）

| 编号 | 问题 | 估时 |
|------|------|------|
| H8 | .gitignore 补全 | 30min |
| M4 | 文档工具数统一 | 30min |
| M5 | README 更新 | 1h |

### 第四批：测试改进（预计 8-10 小时）

| 编号 | 问题 | 估时 |
|------|------|------|
| H7 | 测试断言替换 (17 文件) | 4h |
| H10 | 外部工具 skipif 标记 | 1h |
| M2 | 硬编码路径改 pathlib | 2h |
| L4 | 补充 image_gen 测试 | 1h |
| L5 | pytest 配置优化 | 30min |

---

---

## 附录：修订记录

基于多轮人工核实（对照源码逐行验证），本报告历经如下修订：

| 变更 | 初审 (v1) | v2 修订 | v3 修订 | 说明 |
|------|-----------|---------|---------|------|
| C1 数量 | "30+ 函数" | 16 函数 | **31 函数 / 32 调用点** | v2 误标 14 个泄漏函数为安全；v3 确认全部 31 个函数均因无 try-finally 而泄漏（含 get_info/list_sheets 异常路径、stats_summary 双调用点） |
| C2 数量 | "6 处" | 24 处 | **23 处** | v2 混入"未来注意事项"非缺陷条目；pdf.py:1069 定位 tables_to_word→to_text；pdf.py 实为 15 处 |
| H4 (新增) | — | pdf_form_fill 假成功 | *(同 v2)* | 双层 except 后仍返回 success:True |
| H5 (新增) | — | pdf_text_replace 不生效 | *(同 v2)* | new_data 计算后从未写回内容流 |
| M2 (删除) | "PPT 索引未校验" | 已删除 | — | 交叉验证确认 6 个函数均已校验 |
| M7 (删除) | "缓存不可清" | 已删除 | — | deps.py:57 已有 `clear_cache()` |
| H7 测试数量 | "9/18 文件" | 17 文件 | *(同 v2)* 明确口径 | 17 个 test_*.py + 1 个 verify_format.py |
| H9 影响链 | 仅 styles.py | + `__init__.py:5` | *(同 v2)* | 导入错误会沿加载链传播 |
| 总览 HIGH | — | 9 | **10** | v2 遗漏更新总览表 |
| M1 库名 | — | "pdfplumber" | **"pypdf"** | extract_pages/rotate/delete_pages 均用 pypdf |

---

> **报告生成**: DocuFlow 全面审查代理
> **修订版本**: v3（经三次核实修正计数与技术细节）
> **审查文件数**: 60+
> **发现问题**: 29 项（3 严重 / 10 高 / 8 中 / 8 低）— 另有 2 项初审误报已删除
> **已确认修复**: 4 项（前次审计）
> **v3 修正**: C1 从 16→31 函数/32 调用点(含 get_info/list_sheets 异常路径泄漏、stats_summary 双调用点) / C2 从 24→23 处(移除非缺陷条目、修正函数定位) / M1 库名 pdfplumber→pypdf / 总览 HIGH 9→10 / 尾部总计一致
