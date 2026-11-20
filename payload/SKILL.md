---
name: printable-doc
description: 把文章、笔记、周报这类内容排成可直接打印的 A4 Word（通常两页、正反面一张纸）：一页纸结构、固定行距、版心预算、生成后渲染预览逐页核对。用户说「整理成 word」「排版精美一点」「方便打印」「做成一页纸」「把这篇总结一下给我个文档」时用。只负责排版与成稿；抓网页、提炼内容按一般做法处理。
whenToUse: 用户要一份能直接打印的 Word/PDF，重点在版面与观感；不是改现有文档样式，也不是做 PPT。
---

# 打印版 Word

`scripts/printdoc.py` 是排版引擎，`scripts/example_article.py` 是一份可直接跑的完整样例。
**不要从零手写 python-docx**——坑都封在引擎里了（见 `references/pitfalls.md`）。

## 环境准备（第一次用先查，别等报错）

```sh
python3 -c "import docx"        || pip install python-docx      # 必需
python3 -c "import weasyprint"  || pip install weasyprint       # 预览用
command -v pdftoppm             || apt-get install -y poppler-utils
fc-list :lang=zh | head -3                                      # 中文字体
# 来源是视频/音频时才需要（见 references/video-to-text.md）
command -v ffmpeg               || apt-get install -y ffmpeg
python3 -c "import sherpa_onnx" || pip install sherpa-onnx
```

- **python-docx 缺了不能用**：先装，装不上就直接告诉用户缺什么，别硬跑。
- **weasyprint / poppler 缺了不影响成稿**：照常生成 docx、照常跑 `check_docx.py`，
  但**不要声称看过预览**——说清楚「环境缺预览工具，只做了结构校验」。
- **ffmpeg / sherpa-onnx 只在视频来源时需要**：链接是视频才装；缺了就先问用户要文字稿，
  不要拿同题文章凑正文（见 `references/video-to-text.md`）。
- **中文字体**：没有 CJK 字体时 Word 里会整体回退，行高可能变。先
  `apt-get install -y fonts-noto-cjk`；docx 里的字体名仍写「微软雅黑 / 宋体」，
  本机靠 fontconfig 别名兜底，用户的 Windows/WPS 上有真字体。

## 流程

### 0. 用户给的是视频 / 音频时，先把正文取下来

链接是短视频、播客、直播切片时，先找文字稿，找不到再做转写——**不要凭标题猜内容**。

- 有软字幕流：`ffprobe` 看 `stream=codec_type`，有 `subtitle` 就 `ffmpeg -map 0:s:0` 抽出来。
- 只有 `video` + `audio`：抽音轨做语音转写，**不要 OCR 烧录字幕**——
  实测 0.5 秒/帧、2fps 抽 176 帧约 89 秒，识别出来还是乱码。
- 头条视频取直链、SenseVoice 模型下载与调用、环境坑，全在 `references/video-to-text.md`。
- 同题的其他文章不是原文，不要拿来凑正文。

### 1. 先定页数和每页内容

动笔前决定：几页、每页放哪几块。默认按「一张 A4 双面」设计：

- **第 1 页**：小标签 + 大标题 + 一句话副标 + 导语 + 一块「为什么有效」的原理卡 + 2~3 个板块
- **第 2 页**：时间轴 / 数据对比 + 一块「边界与补充」的卡片 + 居中金句收尾

文章天然的「三个部分 × 每部分三条」最适合做板块；数字、变化、阶段适合做时间轴。

### 2. 算容量，别等生成完再看

照 `references/layout.md` 的预算表把高度加一遍，一页上限 **757pt**；
正文页再留 **1.5–2cm（约 42–57pt）** 机动量，预览和 Word 的中文换行可能差一两行。
超了就删一行说明或缩小 `gap`，不要先写满再缩字号。

### 3. 复制样例改文案

```bash
cp scripts/example_article.py my_doc.py
```

只改文字与板块数量，版式参数（字号、行距、间距）沿用样例——
它们是被验证过能刚好铺满一页的数。

### 4. 生成 → 校验 → 预览

```bash
cd <技能目录>/scripts
python3 my_doc.py out.docx
python3 check_docx.py out.docx          # 必须 0 problems
python3 preview.py out.docx preview/    # 出 PDF + 逐页 PNG
```

然后**逐页看 PNG**，重点看五件事：页数对不对、有没有溢出、有没有重叠、大块留白是否突兀、
第 2 页是否只有表格尾巴（「假三页」，见 `pitfalls.md` #12）。
有问题改脚本重跑，不要手工修 docx。

### 5. 交付

```bash
cp out.docx "/sdcard/Download/DSH-Go/<名字>（A4双面打印版）.docx"
cp preview/preview.pdf "/sdcard/Download/DSH-Go/<名字>（预览）.pdf"
# 需要逐页预览图时，再逐个复制为 <名字>-预览第1页.png、第2页……
```

用 `present` 交付 Word 主件（需要时连同预览 PDF）；说明预览 PDF 与 Word 之间仍有细微字体差异
（中文换行、字重），但页脚已经画出来了（`preview.py` 从 `word/footer1.xml` 取文字，画进 `@bottom-center`）
——正式打印用 docx。

## 引擎速查

```python
from printdoc import PrintDoc
d = PrintDoc(footer="来源 · 一句话")
ACCENT, GREY, LIGHT, LINE = d.t['accent'], d.t['grey'], d.t['light'], d.t['line']

d.kicker("分类 · 副标签")          # 小号宽字距红字
d.title("主标题", size=38)
d.lede("一句话副标")
d.rule(color=ACCENT, sz=16, width=3.4)   # 短装饰线；不传 width 就是整宽细线
d.gap(16); d.pagebreak()

p = d.p(0, 11, 18.5)               # before, after, 固定行距
d.text(p, "正文", size=11, color=GREY)  # 同一段可多次调用拼混排

d.callout(build_fn)                # 整宽淡色卡片，build_fn(cell) 里用 d.cell_para/d.border
d.section_head("01", "板块标题", "TAG")
d.item("①", "条目标题", "一行说明")        # 页紧时再加 body_after=5, head_line=15, body_line=14
d.timeline([("第 1 个月", ["变化只是", "「不用催了」"]), ...])
d.quote("所谓自推", ["把决定变成流程，", "最后孩子自己往前走。"], "补充一句")
d.save(path, title=..., author=..., comments=...)
```

卡片里的段落用 `d.cell_para(cell, first=True, after=5, line=15, left=0.42)`，
文字仍用 `d.text(...)`。

## 验收清单

- [ ] 页数与设计一致，分页位置是你指定的（不是自然流）
- [ ] `check_docx.py` 输出 `ok: 0 schema order problems`
- [ ] 预览每页无文字溢出单元格、无元素压字
- [ ] 正文页页底至少留 1.5cm 机动量（约 42pt）
- [ ] 页码/页脚有一行来源说明
- [ ] 中文引号统一用「」，破折号统一用——
- [ ] 两页的末行落点接近（第 2 页明显更空时调 `d.gap()`，不是把内容硬塞）
- [ ] 来源是视频/音频时：正文来自转写，页脚写了平台与账号
- [ ] 没有半页以上的突兀空白（尾页除外）

## 改引擎之后

改 `scripts/printdoc.py` 后先跑一遍完整示例自检：

```bash
cd scripts
python3 example_article.py /tmp/example.docx
python3 check_docx.py /tmp/example.docx      # 0 problems
python3 preview.py /tmp/example.docx /tmp/example_preview   # 应仍是 2 页
```

示例覆盖标题、卡片、时间轴和 `d.quote()` 的居中短线；确认页数与校验都通过再交付。

## 参考

- `references/layout.md`：单位换算、字号行距表、一页预算表、分页策略
- `references/video-to-text.md`：来源是视频/音频时怎么取正文（字幕 → 转写 → 别 OCR 烧录字幕）
- `references/pitfalls.md`：LibreOffice 不可用、XML 顺序、预览重复文本、假三页、窄栏断词等十五个坑
- `scripts/example_article.py`：成品级样例（两页、三板块、时间轴、卡片、金句）
