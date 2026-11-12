# dsh-go-printable-doc

把一篇文章排成**可以直接打印的 A4 Word**（通常两页、正反面一张纸）。

装上这个插件，等于把 `printable-doc` 技能装到本机：之后在对话里说「整理成可以直接打印的 Word」，
AI 就会按固定版面排：一页纸结构、固定行距、版心预算，生成后还会渲染成图片逐页核对。

## 安装

插件市场搜索 **打印文档** / `dsh-go-printable-doc`，或命令行：

```sh
dsh plugin --profile web add dsh-go-printable-doc
```

装完**新开一个对话**（技能在会话创建时加载），输入框上方会出现「排成打印版」按钮，
设置 › 插件 里会多一个「打印文档」面板。

## 它能做什么

- **一页纸排版引擎**：`printdoc.py` 把标题、导语、原理卡、分板块条目、时间轴、金句都封成方法，
  只管填文案，不用手算间距。
- **页数可控**：版心 16.7 × 26.7 cm = 757pt，用固定行距让 Word 和渲染器算出一致的行高，
  内容超了就提前发现，不是打完才发现多出一页。
- **生成后自查**：`check_docx.py` 校验 OOXML 元素顺序（顺序错了 Word 会报「文件已损坏」），
  `preview.py` 渲染成 PNG 逐页看效果。**看过再交付**。
- **踩过的坑都写下来了**：`references/pitfalls.md` 里 13 条，包括 LibreOffice 在手机上起不来、
  `w:tcMar` 用 `w:w`、表格超宽静默溢出、字体缺失等。

## 环境依赖

技能只在你的主机上跑 Python，插件本身不装任何东西。

| 依赖 | 必需 | 说明 |
| --- | --- | --- |
| python3 + `python-docx` | 是 | 生成 docx。缺了技能不能用 |
| `weasyprint` | 否 | 渲染预览图。有它才能「看过再交付」 |
| `pdftoppm`（poppler-utils） | 否 | 把预览 PDF 转成 PNG |
| `fonts-noto-cjk` | 建议 | 本机渲染中文字体；docx 里仍写微软雅黑/宋体，用户机器上有真字体 |

```sh
pip install python-docx weasyprint
apt-get install -y poppler-utils fonts-noto-cjk
```

面板里会实时显示这五项的状态，缺哪项、怎么补都写在上面。

## 面板

`设置 › 插件 › 打印文档`（或对话标签页）能看：

- 已装技能版本 vs 插件包版本（不一致 = 部署没跑或引擎没重启）
- 技能落盘路径
- 五项环境自检
- 「重新部署」按钮

## 技能装到哪

`$DSH_HOME/skills/printable-doc/`（默认 `~/.dsh/skills/printable-doc/`）。

- 升级插件会自动覆盖技能；旧版本留在 `printable-doc.bak-YYYYMMDD-HHmmss/`（只保留最近 2 个）。
- **卸载插件不会删技能**——它是你的文件，要清掉得手动删目录。

## 已知边界

- 技能只在**新会话**生效，装完/升级完要新开对话。
- 只负责排版成稿；抓网页、提炼内容由对话本身完成。
- 生成的是 A4 版式，打印时选「实际大小 / 100%」，不要选「适合页面」。

## 更新记录

- **1.0.1**：修复 `rule(width, align='center')` 只缩右边、金句短线偏左的 bug；补「假三页」「窄栏断词」等坑；交付命名与自检步骤。
- **1.0.0**：首个版本。

## License

MIT
