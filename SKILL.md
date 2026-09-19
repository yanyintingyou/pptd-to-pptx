---
name: pptd-to-pptx
description: 把 Kimi / neo-design 生成的声明式幻灯片工程目录（*.pptd 主题 + pages/*.page + media/）确定性转换为 PPTX。触发场景：用户给出含 .pptd 或 .page 文件的文件夹，要求"转成 pptx / 导出成 PPT / 变成幻灯片"；或需要只读地把这类第三方幻灯片工程落成可编辑的 .pptx。
agent_created: true
version: "1.2.0"
---

# pptd-to-pptx

把 Kimi（neo-design）导出的声明式幻灯片工程转成**可编辑的 PPTX**：几何 1:1，主题、三线表、公式、图片全部还原。

## 一、输入长什么样

```
<project>/
├── <name>.pptd        # YAML：version / title / size:[960,540] / theme{colors,textStyles,tableStyles} / pages 列表
├── pages/NN_xxx.page  # YAML：pageType / background / elements[]
└── media/*.png        # elements 里 src 引用的图片
```

`.pptd` 提供主题；`.page` 提供内容。每个元素的结构：

```yaml
- elementId: title
  elementType: text            # text | shape | line | image | table
  bounds: [60, 40, 840, 32]    # [x, y, w, h]，单位 px
  content: {...}               # 文本元素
  fill / border / shapeName    # 图形元素
  src / fit:{mode}             # 图片元素
  rows / columnWidths / rowHeights / style   # 表格元素
```

引用语法：`$primary` → `theme.colors.primary`；`$pageTitle` / `$body` / `$note` → `theme.textStyles.*`；
`$threeLine` → `theme.tableStyles.*`。元素内联属性（`fontSize` / `color` / `align` / `wrap` / `lineHeight`）覆盖主题值。
文本字段是 HTML 片段，支持 `<p>`、`<span style>`、`<strong>`、`<sub>`、`<br/>` 与 `\(LaTeX\)`。

## 二、关键决策：不要用 slidep

`slidep` 只吃 **SlideDSL**（`<Slide>` JSX、flex 布局、1280×720），**没有 pptd 导入命令**，
两套格式不通用，硬翻译会丢几何精度。本技能的 `scripts/ppd2pptx.py` 是**确定性 python-pptx 转换器**，
`.page` 的绝对坐标 1:1 落到 PPTX。

## 三、安装与用法

作为技能使用时，把本目录放进你所使用 agent 的技能目录即可（各家 agent 的路径约定不同，本仓库不假定具体位置）；
也可以完全脱离 agent，当独立脚本用：

```powershell
pip install python-pptx
python scripts/ppd2pptx.py <含 .pptd 的工程目录> <输出.pptx>
```

脚本自动找目录里唯一的 `.pptd`，按 `pages:` 顺序逐页生成，**全程只读源目录**。
跑完在 `%TEMP%\_build_stats.txt` 写一份每页元素计数（`text/shape/line/image/table`）——
数字与源文件对得上才算没漏元素。失败会打印 `page <xxx.page> element <id>` 精确定位。

## 四、映射规则（改脚本时别踩的坑）

| DSL | PPTX |
|---|---|
| 1 px | 1 pt = 12700 EMU；画布 `size` × 12700 即 slide_width / slide_height |
| `text` | `add_textbox`，四边 margin 归零、`auto_size=NONE`、`word_wrap` 取 `wrap`（缺省 True）、`vertical_anchor` 取 `align[1]` |
| `shape`(rect) | `MSO_SHAPE.RECTANGLE`；无 `border` 时**必须 `line.fill.background()`**，否则有默认描边 |
| `line` | `add_connector(STRAIGHT)`，箭头写 `a:headEnd` / `a:tailEnd type="triangle"` |
| `image` `contain` | 等比缩放后按 bounds **居中** |
| `image` `cover` | 铺满 bounds + `crop_left/right` 或 `crop_top/bottom = (1 − 目标宽高比 / 原图宽高比) / 2` |
| `table` | `add_table` + 自建网格处理 `rowSpan`（行数组会**参差**，靠 `occupied` 集合跳格）+ `cell.merge` |

### 三线表（`$threeLine`）三件套，缺一不可

1. `tblPr` 里 `firstRow=0 bandRow=0`，`tableStyleId` 换成 **`{2D5ABB26-0587-4C30-8999-92F81FD0307C}`（No Style, No Grid）**；
2. 每格 `cell.fill.background()`，否则残留默认蓝斑马纹；
3. 框线只能写 XML：`tcPr` 里插 `a:lnL/lnR/lnT/lnB`，**必须按 CT_TableCellProperties 顺序插到填充元素之前**，
   其余边一律 `a:noFill` 显式清掉。三线 = 表头行 `lnT` 1.5 + `lnB` 1.0、末行 `lnB` 1.5，其余无边。

### python-pptx 够不到的地方

`a:ea` / `a:cs` typeface（按 `a:latin, a:ea, a:cs` 顺序手插 `rPr` 子元素）、
上下标（`rPr@baseline`，下标 −25000 / 上标 30000）、字距（`rPr@spc`，px×100）。

### HTML 片段

自写状态机（`parse_inline`）处理 `<span style="color:#..;font-family:..">` 的 push/pop 栈、
`<strong>`、`<sub>`、`<br/>`（`paragraph.add_line_break()`）、`<p style="margin-top:6px">`（`space_before`）、
`<p style="text-align:center">`（段落级覆盖元素级对齐）；`html.unescape` 处理 `&amp;` / `&nbsp;`。

## 五、公式：双层策略

`\(...\)` 分两类处理：

- **整块公式**（用 `re.fullmatch(r'<p[^>]*>\s*\\\((.*?)\\\)\s*</p>')` 判定）→ matplotlib mathtext 渲成
  600dpi 透明 PNG 居中贴入。中文与堆叠分式能同时排对的关键：
  `mathtext.fontset='custom'` + `mathtext.rm='SimSun'` + `mathtext.it='Times New Roman:italic'`，
  并把 `\text{}` 预处理成 `\mathrm{}`。PNG 逻辑尺寸 = 像素 × 72 / 600，再等比缩放进 bounds。
- **行内公式** → 转**真实可编辑 run**：Times New Roman 斜体 + baseline 上下标；
  `\frac` / `\dfrac` 展成「（分子）／（分母）」，`\bar{X}` 用组合长音符 U+0304，`\left[` → `[`。
  二元/关系运算符前后补空格，但递归进 `_{}` / `^{}` 时用 `tight=True` 不补。

## 六、字体

`.pptd` 的 `customFonts` 是 webfont（URL），PPT 用不了。脚本用 `EA_MAP` 做回退映射，
默认 `思源宋体 → 宋体`（Windows 通用），`黑体` 原样保留。目标机器确实装了思源宋体时，改 `EA_MAP` 即可。

## 七、视觉校验（slidep screenshot 不可用）

`slidep screenshot` 依赖 editor_sdk 本地服务，通常失败。可靠替代是 LibreOffice + pypdfium2：

```powershell
Get-Process soffice* -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Remove-Item "<outdir>\*.pdf" -Force -EA SilentlyContinue      # 关键：先删旧 PDF
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --norestore `
  --convert-to pdf --outdir <outdir> <file.pptx>
python scripts/pdf2png.py <outdir>\<file>.pdf <shots-dir> 1.6
```

**必须先删旧 PDF，再核对 PDF 的修改时间晚于 pptx。**
soffice 在已有 PDF 时可能打印 `Overwriting: ...` 却**实际没有覆盖**（进程/用户配置锁），于是你渲染的是上一版，
会得到"明明改了源文件、导出结果没变"的假象。核对时间戳是唯一可靠的判据。

**更省事的文本核对**（不渲染，先排除内容层问题）：

```python
from pptx import Presentation
s = Presentation('out.pptx').slides[N-1]
print([sh.text_frame.text for sh in s.shapes if sh.has_text_frame])
```

逐页排除内容问题后再看图，能省掉大量渲染往返。另外，读取图片时若怀疑渲染结果没更新
（同路径图片可能被工具或系统缓存），把 PNG 复制成一个新文件名再读。

soffice 会打一行 `Could not find platform independent libraries <prefix>` 噪声，**不影响转换**。

## 八、已知边界

- 只认 `.pptd` + `.page`；不要拿 SlideDSL 工程或 `.pptx` 试。
- **转换器是忠实的**：导出结果"看着不对"时，先回查源 `.page` 对应的 `text` / `bounds`，再怀疑脚本。
  源端缺陷会被原样带出来（例：某次更新把目录页 n1–n5 五个章节号全写成 `"02"`，导出就是五条 "02"）。
  确认是源端问题时**报告给作者、不要擅自改源**——改源还是对产物做后处理，由作者决定。
- 行内分式为保可编辑性展成斜杠式，不做堆叠；只有独立成块的公式才贴图片。
- PPT 的文字度量与浏览器略有差异，极端密集的页面可能有一两行溢出，需要回源文件微调 `bounds` 或文案。
- Windows 上建议用 PowerShell 调用，命令输出重定向到文件再读，避免 stdout 丢失。
