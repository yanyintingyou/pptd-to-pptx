# pptd-to-pptx

[English](README.md) | **中文文档**

把 **Kimi / neo-design** 导出的声明式幻灯片工程（`.pptd` 主题 + `pages/*.page` 元素表 + `media/` 图片）转换为**完全可编辑的 PowerPoint 演示文稿**。

几何**1:1** 还原——`.page` 用绝对坐标描述每一个元素，因此不存在重新排版、重排流或内容丢失。主题色、字号字体、三线表、LaTeX 公式、图形、连接线与图片全部保留。

```
project/                      ->  deck.pptx（16:9，可编辑）
├── slides.pptd  （主题）          20 页
├── pages/01..20.page （元素）
└── media/*.png
```

## 为什么不用 `slidep`？

`slidep`（SlideDSL 系 PPT 生成器的命令行工具）只吃 **SlideDSL**（`<Slide>` JSX、flex 布局、1280×720 画布），**没有 `.pptd` 的导入命令**。两套格式本身毫无关系：`.page` 是绝对定位的 YAML 声明，SlideDSL 是 flex 布局 DSL。互相翻译必然损失几何精度。

所以本项目改用**确定性的 `python-pptx` 转换器**。源文件里的 1 px 等于输出里的 1 pt，于是 960×540 的画布正好落在 13.333in × 7.5in 的宽屏页面上。

## 安装

需要 Python 3.9+ 与几个依赖：

```bash
pip install python-pptx pillow pyyaml matplotlib
```

克隆本仓库即可；如果想当作技能用，把目录放进 agent 的技能目录（如 `~/.workbuddy/skills/`），`SKILL.md` 是操作手册。

## 用法

```bash
python scripts/ppd2pptx.py <工程目录> <输出.pptx>
```

脚本会自动定位工程目录里唯一的 `.pptd`，按其 `pages:` 列出的顺序逐页渲染，**全程只读源目录**。跑完会在 `%TEMP%\_build_stats.txt` 写一份每页的元素计数（`text / shape / line / image / table`）——数字与源文件对得上，就说明没有漏元素。报错会精确到 `page <xxx.page> element <id>`。

## 映射规则

| 源 DSL | PowerPoint |
|---|---|
| 1 px | 1 pt = 12700 EMU；画布 `size` × 12700 即 slide 宽高 |
| `text` | `add_textbox`，四边 margin 归零、`auto_size=NONE`、`word_wrap` 取 `wrap`、`vertical_anchor` 取 `align[1]` |
| `shape`(rect) | `MSO_SHAPE.RECTANGLE`；未声明 `border` 时**必须**调 `line.fill.background()`，否则会冒出 PowerPoint 默认描边 |
| `line` | `add_connector(STRAIGHT)`；箭头写成 `a:headEnd` / `a:tailEnd type="triangle"` |
| `image` `contain` | 等比缩放后在 `bounds` 内居中 |
| `image` `cover` | 铺满 `bounds` 并用 `crop_left/right` 或 `crop_top/bottom = (1 − 框宽高比 / 图宽高比) / 2` 裁切 |
| `table` | `add_table` + 自建网格处理 `rowSpan`（源数组里被合并的行是**参差**的，用一个 `occupied` 集合跳格）+ `cell.merge` |

主题引用（`$primary` → `theme.colors`，`$pageTitle` / `$body` / `$note` → `theme.textStyles`，`$threeLine` → `theme.tableStyles`）先解析，再被元素内联属性覆盖。

### 三线表：三件套缺一不可

否则 PowerPoint 默认的斑马纹表样式会渗出来：

1. `tblPr` 里设 `firstRow=0`、`bandRow=0`，并把 `tableStyleId` 换成 **`{2D5ABB26-0587-4C30-8999-92F81FD0307C}`（No Style, No Grid）**；
2. 每个单元格调 `cell.fill.background()`，不然默认蓝色斑马纹留着；
3. 框线没有 API：往 `tcPr` 里插 `a:lnL / lnR / lnT / lnB`，**必须按 CT_TableCellProperties 顺序插在填充元素之前**，其余边用 `a:noFill` 显式清掉。三线 = 表头 `lnT` 1.5 + `lnB` 1.0、末行 `lnB` 1.5，其余无边。

### `python-pptx` 没有 API 的地方

- 东亚字体：按 `a:latin → a:ea → a:cs` 顺序往 `rPr` 里插 `a:ea` / `a:cs`。
- 上下标：`rPr@baseline`（下标 −25000，上标 30000）。
- 字距：`rPr@spc`（px × 100）。

### HTML 片段

一个小的状态机（`parse_inline`）处理 `<span style="color:#..;font-family:..">` 的进栈出栈，以及 `<strong>`、`<sub>`、`<br/>`（`paragraph.add_line_break()`）、`<p style="margin-top:6px">`（`space_before`）、`<p style="text-align:center">`（段落级对齐覆盖元素级）。`html.unescape` 负责 `&amp;` / `&nbsp;`。

## 公式：双层策略

行内数学（`\(...\)`）按出现形态分两种处理：

- **独立成块的公式**（用 `re.fullmatch(r'<p[^>]*>\s*\\\((.*?)\\\)\s*</p>')` 判定）交给 matplotlib 的 mathtext 以 600 dpi 渲到透明 PNG 上再居中贴入。让中文与堆叠分式**同时**排对的关键：`mathtext.fontset='custom'`、`mathtext.rm='SimSun'`、`mathtext.it='Times New Roman:italic'`，并把 `\text{}` 预处理成 `\mathrm{}`。PNG 的逻辑尺寸 = 像素 × 72 / 600，再等比缩放进 `bounds`。
- **行内公式**转成**真实可编辑的 run**：Times New Roman 斜体 + baseline 上下标。`\frac` / `\dfrac` 展成「（分子）/（分母）」，`\bar{X}` 用组合长音符 U+0304，`\left[` 变 `[`。二元与关系运算符前后补空格——但递归进 `_{}` / `^{}` 时用 `tight=True` 抑制。

## 字体

`.pptd` 的 `customFonts` 指向 web 字体（URL），PowerPoint 用不了。脚本通过 `EA_MAP` 映射回本机已装字族：默认 `思源宋体 → 宋体`（Windows 通用），`黑体` 原样保留。目标机器确实装了思源宋体时，改 `EA_MAP` 即可。

## 视觉校验

`slidep screenshot` 依赖本地 editor SDK 服务，一般会失败。可靠的替代是 LibreOffice + pypdfium2：

```powershell
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --norestore `
  --convert-to pdf --outdir <outdir> <file.pptx>
python scripts/pdf2png.py <outdir>\<file>.pdf <shots-dir> 1.6
```

然后逐页看 PNG 核对。`soffice` 会打一行 `Could not find platform independent libraries <prefix>` 警告，无害，不影响转换。

## 仓库结构

```
.
├── SKILL.md                # 操作手册（中文）
├── README.md               # 英文说明
├── README.zh-CN.md         # 本文件
└── scripts/
    ├── ppd2pptx.py         # 转换器本体
    └── pdf2png.py          # PDF → 逐页 PNG，用于校验
```

## 已知边界

- 只认 `.pptd` + `.page`。读不了 SlideDSL 工程，也读不了现成的 `.pptx`。
- 行内分式为保持可编辑性展成斜杠式，不做堆叠；只有独立成块的公式才贴图片。
- PowerPoint 的文字度量与浏览器略有差异，极端密集的页面可能溢出两三行，需要回源文件微调 `bounds` 或文案。
- Windows 上建议用 PowerShell 调用，并把命令输出重定向到文件再读——stdout 常被宿主 shell 吞掉。

## 许可证

尚未包含许可证文件——再分发前请先补一个。
