# 坑清单

按踩到的顺序记，每条都附「怎么发现的 / 怎么绕开」。

## 1. LibreOffice 在本机（proot）完全不可用

```
ERROR: /proc not mounted - LibreOffice is unlikely to work well if at all
```

它 0.05 秒就退出，不产出任何 PDF。`--headless`、`--norestore`、
`-env:UserInstallation`、`SAL_USE_VCLPLUGIN=svp` 全试过，都无效。

**结论**：不要用 `soffice --convert-to pdf` 做校验，直接跳过。装它也白装。

## 2. python-docx 不保证 OOXML 子元素顺序

往 `pPr` 里先 append `w:spacing` 再 append `w:pBdr`，文件能存，
Word 打开会说「内容有问题，是否恢复」。

**绕开**：所有属性插入都走 `_put()` / `_insert_ordered()`，带一张 schema 顺序表。
写完用 `scripts/check_docx.py` 验一遍，必须 0 problems。

`python-docx` 自己的 `insert_element_before` 会做 `qn()` 转换，
所以只能传 `'w:xxx'` 字符串；但普通 lxml 元素没有这个方法，得自己写。

## 3. weasyprint 预览里文字会重复两遍

```python
# 错：w:r 的 itertext() 会把 w:t 连同 rPr 里的内容一起收，导致重复
t = ''.join(r.itertext())

# 对：只取 w:t 的文本
t = ''.join((e.text or '') for e in r.iter(el('t')))
```

## 4. `w:tcMar` 的宽度属性是 `w:w`，不是 `w:val`

写成 `w:val` 直接 `TypeError`，且报错信息不指向真正原因。

## 5. 中文字体缺失会毁掉整份排版

本机只有 Noto Sans/Serif CJK，没有微软雅黑、宋体。两种兜底：

- 给 fontconfig 建别名（`~/.config/fontconfig/fonts.conf`）：
  微软雅黑 → Noto Sans CJK SC，宋体 → Noto Serif CJK SC；
- 预览 HTML 用 `@font-face` 直接指向 `~/.fonts/NotoSansSC.otf`。

字体名仍写「微软雅黑 / 宋体」——用户拿到的 docx 是在 Windows/WPS 上打开的，
那里有真字体；本机只是预览。

## 6. 固定行距是页数稳定的关键

`w:lineRule="exact"` + `w:line = 行距 × 20`。字体被替换时行高不变，
所以「预览 2 页」和「Word 打开 2 页」能对上。用「单倍行距」就会随字体漂移。

## 7. 表格宽度超过版心不会报错

Word 会默默缩放或挤出页面边界。技能里 `PrintDoc.table()` 直接断言，
宁可在生成时报错。

## 8. 高度 0 的空段落要用 1pt 字号

`doc.add_paragraph()` 默认继承 10pt，即使没内容也会撑出一行。
用 `d.gap(pt)`：1pt 字号 + `space_after`。

## 9. 装饰线不要用整宽段落下边框做「短线」

段落边框横贯「左缩进到右缩进」之间的宽度，所以短线必须同时设 `left`/`right`。
`PrintDoc.rule(width=1.5, align='center')` 会把两侧都缩进 `(W - width) / 2`；
1.0.0 只设了 `right`，短线会贴在左边（`d.quote()` 的金句短线因此偏左），1.0.1 已修。
靠右同理：`align='right'` 只缩左边。

## 10. weasyprint 与 Word 有可见差异

预览是**近似**，不是像素级：
- 中文引号、破折号的字宽略有不同；
- 加粗字重（Noto 只有 400/700）；
- 页脚在预览里不渲染（脚本没解析 footer）。

预览用来抓「页数、溢出、重叠、大块留白」，不要用它逐像素对齐。

## 11. 交付路径

成品放 `/sdcard/Download/DSH-Go/`（用户手机可见），
同时给一份 PDF 预览，方便手机上直接看效果。

## 12. 预览里表格行仍会被拆开，出现「假三页」

给表格行设了 `cantSplit`，Word 会整行下移；但 weasyprint 预览不认，仍会在行中间断开。
症状：预览 3 页，第 2 页只有一两行（正好是第 1 页表格的尾巴），第 3 页才是本想放的第 2 页。

这不是分页坏了，是第 1 页超了 20–60pt。先看第 1 页表格底部还剩多少：

- 余量 < 1.5cm：把正文行距减 0.5–0.7pt、段后减 2pt，或缩一个 gap；
- 不要靠删正文硬塞；`cantSplit` 救不了预览，只会在 Word 里把整行推到下一页。

本次实测：正文栏 12.0cm、11pt/20.5pt 行距时，第 1 页页底只剩约 1cm；
把行距收到 19.8pt、段后收到 8pt 后留 1.7cm。另一次把正文栏缩到 11.7cm，
P4 多出一行，预览立刻变成「假三页」——第 2 页只有两行表格尾巴，第 3 页才是正文第 2 页。

## 13. 窄栏里的孤字、断词与硬断

4.7cm 的边注栏，扣掉左缩进 0.5cm、右缩进 0.15cm，正文宽约 4.05cm：

- 7.5pt 中文 ≈ 15 字/行；8.5pt ≈ 13 字/行。
- 15 字的批注（「点明主题，我和孙悟空参观龙宫。」）在 8.5pt 会把「宫。」挤到下一行；
  降到 7.5pt 正好一行。29 字的批注在 7.5pt 断成 15+14 字，若断在「一 / 天」这种两字词中间，
  用两个段落在「的」后硬断，预览和 Word 都稳定；别用 `w:br`（预览脚本不解析）。

导语同理：中文按字换行，会在「花 / 园。」「想 / 象」处断开。先按「版心宽 ÷ 字号」算每行字数，
把第一句控制在整数行内；差的字不多时，在句号后用两个段落硬断（第二段 `after` 保持原值），
视觉上仍是一段。本次 16.7cm、10.1pt 的导语第一句 46 字，正好一行。
