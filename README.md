# pptd-to-pptx

**English** | [中文文档](README.zh-CN.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Convert declarative slide projects produced by **Kimi / neo-design** (`.pptd` theme + `pages/*.page` elements + `media/` images) into **fully editable PowerPoint decks**.

Geometry is preserved **1:1** — the source `.page` format describes every element with absolute coordinates, so there is no re-layout, no re-flow and no content loss. Theme colors, font styles, three-line tables, LaTeX formulas, shapes, connectors and images all carry over.

```
project/                      ->  deck.pptx (16:9, editable)
├── slides.pptd  (theme)          20 slides
├── pages/01..20.page (elements)
└── media/*.png
```

## Why not `slidep`?

`slidep` — the CLI used by SlideDSL-based PPT generators — only consumes **SlideDSL** (`<Slide>` JSX, flexbox layout, 1280×720 canvas). It has **no import command for `.pptd`**, and the two formats are unrelated: `.page` files are absolute-positioned YAML declarations, SlideDSL is a flex layout DSL. Translating one into the other loses geometric fidelity.

So this project uses a **deterministic `python-pptx` converter** instead. 1 px in the source equals 1 pt in the output, which makes an 960×540 canvas land exactly on a 13.333in × 7.5in widescreen slide.

## Install

Python 3.9+ and a few packages:

```bash
pip install python-pptx pillow pyyaml matplotlib
```

Clone the repo, or drop the folder into your agent's skills directory (e.g. `~/.workbuddy/skills/`) to use it as an installable skill — `SKILL.md` is the operational reference.

## Usage

```bash
python scripts/ppd2pptx.py <project-dir> <output.pptx>
```

The script locates the single `.pptd` inside the project directory and renders every page listed in its `pages:` key, **read-only** on the source tree. On completion it writes per-page element counts to `%TEMP%\_build_stats.txt` (`text / shape / line / image / table`) — if those numbers match the source, nothing was dropped. Failures are reported as `page <xxx.page> element <id>`.

## How it works

| Source DSL | PowerPoint |
|---|---|
| 1 px | 1 pt = 12700 EMU; canvas `size` × 12700 gives slide width/height |
| `text` | `add_textbox`, zeroed margins, `auto_size=NONE`, `word_wrap` from `wrap`, `vertical_anchor` from `align[1]` |
| `shape` (rect) | `MSO_SHAPE.RECTANGLE`; when no `border` is declared you **must** call `line.fill.background()`, otherwise PowerPoint's default outline shows up |
| `line` | `add_connector(STRAIGHT)`; arrowheads written as `a:headEnd` / `a:tailEnd type="triangle"` |
| `image` `contain` | scale to fit, then center inside `bounds` |
| `image` `cover` | fill `bounds` and crop with `crop_left/right` or `crop_top/bottom = (1 − box_ar / image_ar) / 2` |
| `table` | `add_table` plus a hand-built grid for `rowSpan` (merged rows in the source arrays are **ragged** — an `occupied` set skips taken cells) then `cell.merge` |

Theme references (`$primary` → `theme.colors`, `$pageTitle` / `$body` / `$note` → `theme.textStyles`, `$threeLine` → `theme.tableStyles`) are resolved first, then overridden by inline element properties.

### Three-line tables: the three-part recipe

All three are required, or PowerPoint's default banded table style bleeds through:

1. In `tblPr`, set `firstRow=0` and `bandRow=0`, and swap `tableStyleId` for **`{2D5ABB26-0587-4C30-8999-92F81FD0307C}` (No Style, No Grid)**.
2. Call `cell.fill.background()` on every cell — otherwise the default blue banding stays.
3. Borders have no API: insert `a:lnL / lnR / lnT / lnB` into `tcPr` **in CT_TableCellProperties order, before the fill element**, and explicitly clear the other edges with `a:noFill`. The three lines are header `lnT` 1.5 + `lnB` 1.0, and last row `lnB` 1.5; everything else is borderless.

### Where `python-pptx` has no API

- East-Asian typefaces: insert `a:ea` / `a:cs` into `rPr` in `a:latin → a:ea → a:cs` order.
- Super/subscript: `rPr@baseline` (−25000 subscript, 30000 superscript).
- Letter spacing: `rPr@spc` (px × 100).

### HTML fragments

A small state machine (`parse_inline`) handles the push/pop stack of `<span style="color:#..;font-family:..">`, plus `<strong>`, `<sub>`, `<br/>` (`paragraph.add_line_break()`), `<p style="margin-top:6px">` (`space_before`) and `<p style="text-align:center">` (paragraph-level alignment overriding the element's). `html.unescape` deals with `&amp;` / `&nbsp;`.

## Formulas: a two-tier strategy

Inline math (`\(...\)`) is handled in two different ways depending on how it appears:

- **Standalone formulas** (detected with `re.fullmatch(r'<p[^>]*>\s*\\\((.*?)\\\)\s*</p>')`) are rendered by matplotlib's mathtext at 600 dpi onto a transparent PNG and centered in place. To get CJK *and* stacked fractions right at the same time, the script uses `mathtext.fontset='custom'`, `mathtext.rm='SimSun'`, `mathtext.it='Times New Roman:italic'`, and pre-processes `\text{}` into `\mathrm{}`. The PNG's logical size is `pixels × 72 / 600`, then scaled to fit `bounds`.
- **Inline formulas** become **real, editable runs**: Times New Roman italic with baseline super/subscripts. `\frac` / `\dfrac` expand to `(numerator) / (denominator)`, `\bar{X}` uses the combining macron U+0304, `\left[` becomes `[`. Binary and relational operators get surrounding spaces — except when recursing into `_{}` / `^{}`, where `tight=True` suppresses them.

## Fonts

The `customFonts` block in a `.pptd` points at web fonts (URLs), which PowerPoint cannot use. The script maps them back to installed families through `EA_MAP`; by default `思源宋体 → 宋体` (universally present on Windows) while `黑体` is kept as-is. Edit `EA_MAP` if the target machine really does have Source Han Serif.

## Visual verification

`slidep screenshot` depends on a local editor SDK service and generally fails. A reliable substitute is LibreOffice plus pypdfium2:

```powershell
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --norestore `
  --convert-to pdf --outdir <outdir> <file.pptx>
python scripts/pdf2png.py <outdir>\<file>.pdf <shots-dir> 1.6
```

Then inspect the PNGs page by page. `soffice` prints a `Could not find platform independent libraries <prefix>` warning — it is harmless and does not affect the conversion.

## Repository layout

```
.
├── SKILL.md                # operational reference (Chinese)
├── README.md               # this file
├── README.zh-CN.md         # Chinese README
├── LICENSE                 # MIT
└── scripts/
    ├── ppd2pptx.py         # the converter
    └── pdf2png.py          # PDF -> per-page PNG, for verification
```

## Limitations

- Only `.pptd` + `.page` input. It will not read SlideDSL projects or existing `.pptx` files.
- Inline fractions are expanded to slash form so they stay editable; only standalone formulas are embedded as images.
- PowerPoint's text metrics differ slightly from a browser's, so extremely dense pages may overflow by a line or two. Nudge `bounds` or trim the copy in the source.
- On Windows, prefer PowerShell and redirect command output to a file before reading it — stdout is often swallowed by the host shell.

## License

Released under the [MIT License](LICENSE) © 2026 yanyintingyou.

In short: you may use, modify, redistribute and sell this code, including commercially, as long as the copyright notice and the license text stay with it. The software comes with no warranty.
