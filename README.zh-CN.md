# DocuFlow MCP

DocuFlow 鏄竴涓潰鍚戞枃妗ｅ鐞嗙殑 MCP (Model Context Protocol) 鏈嶅姟鍣紝褰撳墠鎻愪緵 **149 涓伐鍏?*锛岃鐩?Word銆丒xcel銆丳owerPoint銆丳DF銆佹牸寮忚浆鎹€丱CR銆丠TML 杞?PPTX 浠ュ強 AI 鍥剧墖鐢熸垚绛夊父瑙佸伐浣滄祦銆?
椤圭洰鐩爣濡備笅锛?
- 閫氳繃缁熶竴鐨?MCP 鎺ュ彛鏆撮湶鏂囨。澶勭悊鑳藉姏
- 鍚屾椂鏀寔鏈湴鏂囨。宸ュ叿閾惧拰杩滅▼ AI 鑳藉姏
- 璁?Claude Code銆丆odex 鍙婂叾浠栧吋瀹?MCP 鐨勫鎴风鍙互鐩存帴鎿嶄綔鏂囨。

## 鍔熻兘鐭╅樀

| 妯″潡 | 宸ュ叿鏁?| 涓昏鑳藉姏 |
| --- | ---: | --- |
| Word (.docx) | 49 | 鏂囨。銆佹钀姐€佹爣棰樸€佽〃鏍笺€佸浘鐗囥€佸垪琛ㄣ€佸垎椤点€侀〉鐪夐〉鑴氥€佹壒娉ㄣ€佹牱寮忋€佹ā鏉裤€佸鍑?|
| Excel (.xlsx) | 33 | 宸ヤ綔绨裤€佸伐浣滆〃銆佸崟鍏冩牸銆佸叕寮忋€佺粺璁°€佸浘琛ㄣ€侀€忚琛ㄣ€佹潯浠舵牸寮忋€佹暟鎹鐞?|
| PowerPoint (.pptx) | 30 | 骞荤伅鐗囥€佹枃鏈銆佸舰鐘躲€佸浘琛ㄣ€佸姩鐢汇€佸垏鎹€佹瘝鐗堛€佸崰浣嶇 |
| PDF | 23 | 鎻愬彇銆佸悎骞躲€佹媶鍒嗐€佹棆杞€佹按鍗般€佽〃鍗曘€佽劚鏁忋€佽浆鎹?|
| 鏍煎紡杞崲 | 4 | 鍩轰簬 pandoc 鐨?40+ 鏍煎紡浜掕浆 |
| OCR | 4 | 鍥剧墖 / PDF 鏂囧瓧璇嗗埆锛屾敮鎸?Tesseract 涓?OpenAI 鍏煎 completion 鎺ュ彛 |
| HTML -> PPTX | 3 | 灏?HTML 椤甸潰杞崲涓?PowerPoint |
| AI 鍥剧墖鐢熸垚 | 3 | 鏂囩敓鍥俱€丳PT 鎻掑浘鐢熸垚 |

## 瀹夎

### 鏂瑰紡涓€锛氫竴閿畨瑁?
```bash
cd DocuFlow
python install.py
```

### 鏂瑰紡浜岋細鎵嬪姩瀹夎

```bash
pip install -e .
```

瀹夎瀹屾垚鍚庯紝椤圭洰浼氭毚闇?MCP 鏈嶅姟鍏ュ彛锛?
```bash
docuflow-mcp
```

## 鍦?MCP 瀹㈡埛绔腑鎺ュ叆

鍦ㄦ敮鎸?MCP 鐨勫鎴风閰嶇疆涓姞鍏ワ細

```json
{
  "mcpServers": {
    "docuflow": {
      "command": "docuflow-mcp"
    }
  }
}
```

閰嶇疆瀹屾垚鍚庨噸鍚鎴风锛屽嵆鍙€氳繃鑷劧璇█璋冪敤宸ュ叿锛屼緥濡傦細

- 鈥滃垱寤轰竴涓柊鐨?Word 鏂囨。 `report.docx`锛屾爣棰樻槸銆婃湀搴﹂攢鍞姤鍛娿€嬨€傗€?- 鈥滄彁鍙?`invoice.pdf` 涓殑琛ㄦ牸骞跺啓鍏?Excel銆傗€?- 鈥滄妸 `report.docx` 杞垚 PDF銆傗€?- 鈥滃 `scan.png` 鍋?OCR锛屽彧杩斿洖璇嗗埆鏂囨湰銆傗€?
## 涓ょ浣跨敤鏂瑰紡

DocuFlow 鍚屾椂鏀寔涓ょ浣跨敤鏂瑰紡锛屾枃妗ｅ垎鍒鏄庡涓嬨€?
### 1. MCP 瀹㈡埛绔娇鐢?
杩欐槸榛樿涓旀帹鑽愮殑浣跨敤鏂瑰紡銆備綘鍦?Claude Code銆丆odex 绛夊鎴风閲岄厤缃?`docuflow-mcp` 鍚庯紝閫氳繃 MCP 宸ュ叿鎴栬嚜鐒惰瑷€杩涜璋冪敤銆?
### 2. Python 鍐呴儴 API 鐩存帴璋冪敤

椤圭洰鍐呴儴妯″潡涔熷彲浠ョ洿鎺ュ鍏ワ紝渚嬪锛?
```python
from docuflow_mcp.extensions.ocr import OCROperations
```

杩欑鏂瑰紡閫傚悎寮€鍙戙€佹祴璇曟垨浜屾闆嗘垚锛屼笉绛夊悓浜?MCP 瀹㈡埛绔娇鐢ㄦ柟寮忋€傛湰鏂囦腑鐨?Python 绀轰緥榛樿闈㈠悜寮€鍙戣€呫€?
## OCR 鏋舵瀯璇存槑

褰撳墠杩滅▼ OCR 宸茬粺涓€鍒?**OpenAI 鍏煎鐨?`chat/completions` 鎺ュ彛**锛屼笉鍐嶄緷璧?Anthropic / OpenAI 鐨?Python SDK銆?
OCR 鐩稿叧宸ュ叿鍖呮嫭锛?
- `ocr_image`
- `ocr_pdf`
- `ocr_to_docx`
- `ocr_status`

### 鏀寔鐨?OCR 寮曟搸

- `tesseract`
  鏈湴 OCR锛屼笉闇€瑕佽繙绋?API銆?- `completion`
  杩滅▼ OCR锛岃蛋 OpenAI 鍏煎鐨?`chat/completions` 鎺ュ彛銆?- `claude`
  鍏煎鍒悕锛屽唴閮ㄤ細鏄犲皠鍒?`completion`銆?- `auto`
  浼樺厛灏濊瘯 Tesseract锛涘繀瑕佹椂鍥為€€鍒?`completion`銆?
## OCR 閰嶇疆

杩滅▼ OCR 榛樿璇诲彇椤圭洰鏍圭洰褰曚笅鐨?`ocr_config.json`锛?
```json
{
  "api_url": "https://your-api.example.com/v1/chat/completions",
  "model": "grok-4.1-thinking",
  "timeout": 120,
  "api_key": "your-api-key"
}
```

瀛楁璇存槑锛?
- `api_url`锛歄penAI 鍏煎 completion 鎺ュ彛鍦板潃銆傚繀椤诲啓瀹屾暣璺緞锛岄€氬父鏄?`/v1/chat/completions`銆?- `model`锛氶粯璁や娇鐢ㄧ殑杩滅▼妯″瀷銆?- `timeout`锛氳姹傝秴鏃舵椂闂达紝鍗曚綅绉掋€?- `api_key`锛氳繙绋嬫湇鍔″瘑閽ャ€?
### OCR 鍙傛暟浼樺厛绾?
瀵逛簬 `completion` 璺緞锛屽弬鏁扮敓鏁堥『搴忓涓嬶細

1. 宸ュ叿璋冪敤鍙傛暟
2. `ocr_config.json`
3. 浠ｇ爜鍐呴粯璁ゅ€?
绀轰緥锛?
- `ocr_image(image_path="scan.png", engine="completion")`
  浣跨敤 `ocr_config.json` 涓殑 `api_url`銆乣model`銆乣timeout`銆乣api_key`
- `ocr_image(image_path="scan.png", engine="completion", model="grok-4")`
  浣跨敤鏄惧紡浼犲叆鐨?`model="grok-4"`锛屽叾浣欏弬鏁扮户缁粠 `ocr_config.json` 璇诲彇

### OCR 榛樿鎻愮ず璇嶇瓥鐣?
榛樿 completion OCR 鎻愮ず璇嶅凡缁忔敹绱т负涓ユ牸 OCR 妯″紡锛岀洰鏍囨槸锛?
- 鍙緭鍑哄浘鐗囦腑鐨勬枃瀛?- 涓嶈В閲娿€佷笉鎬荤粨銆佷笉鍥炵瓟闂
- 涓嶈緭鍑?Markdown 鏍囬銆佸垪琛ㄣ€佷唬鐮佸潡
- 灏介噺淇濈暀鍘熷鎹㈣鍜岄槄璇婚『搴?- 灏介噺閬垮厤閲嶅杈撳嚭

濡傛灉浣犳湁鐗规畩鐗堝紡闇€姹傦紝鍙互鍦?`ocr_image`銆乣ocr_pdf`銆乣ocr_to_docx` 涓樉寮忎紶鍏?`prompt` 瑕嗙洊榛樿鎻愮ず璇嶃€?
## OCR 渚濊禆涓庡墠缃潯浠?
### 鍥剧墖 OCR

- 鏈湴 Tesseract 璺緞闇€瑕佸畨瑁?`tesseract`
- 杩滅▼ completion 璺緞闇€瑕侀厤缃?`ocr_config.json`

### 鎵弿 PDF OCR

`ocr_pdf` 鍜?`ocr_to_docx` 鍦ㄥ鐞嗘壂鎻?PDF 鏃朵緷璧栵細

- `pdf2image`
- `Pillow`
- Windows 涓婇€氬父杩橀渶瑕佸畨瑁?**Poppler**锛屽苟纭繚鍏跺彲鎵ц鏂囦欢鍙绯荤粺璁块棶

鎺ㄨ崘瀹夎锛?
```bash
pip install pdf2image Pillow
```

濡傛灉鏄?Windows锛岃棰濆瀹夎 Poppler 鍚庡啀閲嶈瘯銆傚惁鍒欐壂鎻?PDF 铏界劧鑳借繘鍏?OCR 娴佺▼锛屼絾浼氬湪 PDF 杞浘鐗囬樁娈靛け璐ャ€?
## OCR 浣跨敤绀轰緥

### MCP 鍦烘櫙涓殑鑷劧璇█绀轰緥

- 鈥滃 `scan.png` 鍋?OCR锛屽彧杩斿洖璇嗗埆鏂囨湰銆傗€?- 鈥滄妸 `scan.pdf` 鐨勫墠 3 椤靛仛 OCR锛屽啀瀵煎嚭涓?Word銆傗€?- 鈥滄煡鐪?OCR 褰撳墠浣跨敤鐨勬ā鍨嬪拰鎺ュ彛閰嶇疆銆傗€?
### Python 鐩存帴璋冪敤绀轰緥

#### 鍗曞紶鍥剧墖 OCR

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

#### OCR 鍚庣敓鎴?Word

```python
result = OCROperations.ocr_to_docx(
    source="scan.pdf",
    output_path="scan_ocr.docx",
    engine="completion",
)
```

#### 鏌ョ湅 OCR 鐘舵€?
```python
status = OCROperations.get_status()
```

`ocr_status` 浼氳繑鍥烇細

- 褰撳墠 `ocr_config.json` 璺緞
- 褰撳墠鐢熸晥鐨?`api_url` / `model` / `timeout`
- `tesseract` 鍜?`completion` 鐨勫彲鐢ㄦ€?- `claude -> completion` 鐨勫吋瀹瑰埆鍚嶅叧绯?
## 椤圭洰缁撴瀯

```text
DocuFlow/
|-- src/docuflow_mcp/
|   |-- server.py                # MCP 鏈嶅姟鍏ュ彛
|   |-- document.py              # Word 鏂囨。鐩稿叧鎿嶄綔
|   |-- core/
|   |   |-- registry.py          # 宸ュ叿娉ㄥ唽涓庡垎鍙?|   |   `-- middleware.py        # 涓棿浠?|   |-- extensions/
|   |   |-- excel.py             # Excel 鎿嶄綔
|   |   |-- pdf.py               # PDF 鎿嶄綔
|   |   |-- ppt.py               # PowerPoint 鎿嶄綔
|   |   |-- converter.py         # 鏍煎紡杞崲
|   |   |-- ocr.py               # OCR
|   |   |-- image_gen.py         # AI 鍥剧墖鐢熸垚
|   |   |-- html_to_pptx.py      # HTML 杞?PPTX
|   |   |-- styles.py            # 鏍峰紡绠＄悊
|   |   |-- templates.py         # 妯℃澘绠＄悊
|   |   `-- validator.py         # 鏍￠獙涓庝慨澶?|   `-- utils/
|       |-- deps.py              # 渚濊禆妫€鏌?|       `-- paths.py             # 璺緞鏍￠獙
|-- tests/                       # 娴嬭瘯
|-- scripts/                     # 鎵归噺淇涓庣淮鎶よ剼鏈?|-- install.py                   # 瀹夎鑴氭湰
|-- install_codex.py             # Codex 瀹夎鑴氭湰
|-- pyproject.toml               # 椤圭洰閰嶇疆
|-- README.md                    # English documentation
`-- README.zh-CN.md              # Chinese documentation
```

## 寮€鍙戜笌娴嬭瘯

瀹夎寮€鍙戜緷璧栵細

```bash
pip install -e .[dev]
```

杩愯鍏ㄩ儴娴嬭瘯锛?
```bash
pytest -q
```

浠呰繍琛?OCR 娴嬭瘯锛?
```bash
pytest -q tests/test_ocr.py
```

## 甯歌闂

### 1. `ocr_image` 鑳借繍琛岋紝浣嗚緭鍑烘湁瑙ｉ噴鎬у唴瀹?
浼樺厛妫€鏌ワ細

- 褰撳墠妯″瀷鏄惁閫傚悎 OCR 鎸囦护閬靛惊
- 鏄惁瑕嗙洊浜嗛粯璁?`prompt`
- 鏈嶅姟绔槸鍚﹀仛浜嗘ā鍨嬪埆鍚嶆槧灏?
### 2. `ocr_pdf` 澶辫触锛屾彁绀虹己灏戜緷璧?
浼樺厛妫€鏌ワ細

- 鏄惁瀹夎浜?`pdf2image`
- Windows 鏄惁瀹夎浜?Poppler
- `ocr_status` 涓?`pdf2image` 鍜?`PIL` 鏄惁鍙敤

### 3. 杩滅▼ OCR 璇锋眰澶辫触鎴栬秴鏃?
浼樺厛妫€鏌ワ細

- `ocr_config.json` 涓殑 `api_url` 鏄惁涓哄畬鏁寸殑 `/v1/chat/completions`
- `api_key` 鍜?`model` 鏄惁鏈夋晥
- 鏈嶅姟绔槸鍚﹁兘绋冲畾澶勭悊澶氭ā鎬佽姹?- 鍙嶅悜浠ｇ悊鎴栦笂娓告湇鍔℃槸鍚﹀瓨鍦ㄨ秴鏃?/ TLS 闂

## 瀹夊叏璇存槑

- `ocr_config.json` 榛樿宸插姞鍏?`.gitignore`
- 涓嶈鎶婄湡瀹?API 瀵嗛挜鍐欏叆 README銆佹祴璇曟牱渚嬫垨鎻愪氦璁板綍
- 濡傛灉闇€瑕佸湪 CI 涓繍琛岃繙绋?OCR锛岃浣跨敤鐙珛娴嬭瘯鐜鍜屽彈闄愬瘑閽?
## 璁稿彲璇?
Apache License 2.0。完整文本见 `LICENSE`。