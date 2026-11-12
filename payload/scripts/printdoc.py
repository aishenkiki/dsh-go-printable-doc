#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""printdoc —— A4 打印版 Word 的排版工具库。

只负责「怎么排」，不负责「写什么」。用法见 example_article.py。

设计约束（都是踩过坑换来的）：
  * 版心宽 16.7cm（A4 21cm − 左右各 2.15cm），所有表格宽度不得超过它；
  * 全部使用固定行距（exact），字体替换不会改变行高，页数才稳定；
  * 每个 pPr/rPr/tblPr/tcPr/trPr 子元素都按 OOXML schema 顺序插入，
    否则 Word 会报「内容有问题」。
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml.etree import QName

# ---------------------------------------------------------------- schema 顺序

PPR = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr',
       'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd',
       'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
       'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
       'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
       'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
       'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
       'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
       'w:pPrChange']

RPR = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps',
       'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow',
       'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish',
       'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position',
       'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd',
       'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang',
       'w:eastAsianLayout', 'w:specVanish', 'w:oMath']

TBLPR = ['w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual',
         'w:tblStyleRowBandSize', 'w:tblStyleColBandSize', 'w:tblW', 'w:jc',
         'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders', 'w:shd',
         'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption',
         'w:tblDescription']

TCPR = ['w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge',
        'w:tcBorders', 'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection',
        'w:tcFitText', 'w:vAlign', 'w:hideMark', 'w:headers', 'w:cellIns',
        'w:cellDel', 'w:cellMerge', 'w:tcPrChange']

TRPR = ['w:cnfStyle', 'w:divId', 'w:gridBefore', 'w:gridAfter', 'w:wBefore',
        'w:wAfter', 'w:cantSplit', 'w:trHeight', 'w:tblHeader',
        'w:tblCellSpacing', 'w:jc', 'w:hidden']

DEFAULT_THEME = {
    'ink': '1B3B33', 'accent': 'B0543A', 'grey': '4E5A56',
    'light': '8A918C', 'band': 'F4F1E9', 'line': 'D6D2C6',
    'title': '微软雅黑', 'body': '宋体', 'sans': '微软雅黑',
}


def _insert_ordered(parent, el, tag, order):
    succ = {t.split(':')[1] for t in order[order.index(tag) + 1:]}
    for child in parent:
        if QName(child).localname in succ:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def _put(parent, tag, order, attrs=None):
    el = parent.find(qn(tag))
    if el is None:
        el = _insert_ordered(parent, OxmlElement(tag), tag, order)
    for k, v in (attrs or {}).items():
        el.set(qn(k), str(v))
    return el


def _fmt(p, before=0, after=0, line=16.0, align=None,
         left=0.0, hanging=0.0, right=0.0, keep=False):
    """pt 单位的间距与固定行距；left/hanging/right 单位 cm。"""
    ppr = p._p.get_or_add_pPr()
    _put(ppr, 'w:snapToGrid', PPR, {'w:val': '0'})
    sp = _put(ppr, 'w:spacing', PPR)
    sp.set(qn('w:before'), str(int(round(before * 20))))
    sp.set(qn('w:after'), str(int(round(after * 20))))
    sp.set(qn('w:line'), str(int(round(line * 20))))
    sp.set(qn('w:lineRule'), 'exact')
    if left or hanging or right:
        ind = _put(ppr, 'w:ind', PPR)
        ind.set(qn('w:left'), str(int(round(left * 567))))
        ind.set(qn('w:right'), str(int(round(right * 567))))
        if hanging:
            ind.set(qn('w:hanging'), str(int(round(hanging * 567))))
    if align:
        _put(ppr, 'w:jc', PPR,
             {'w:val': {'left': 'left', 'center': 'center',
                        'right': 'right'}[align]})
    if keep:
        _put(ppr, 'w:keepNext', PPR, {'w:val': '1'})
    return p


def _border(p, edges):
    pbdr = _put(p._p.get_or_add_pPr(), 'w:pBdr', PPR)
    order = ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:between', 'w:bar']
    for side, (sz, color) in edges.items():
        old = pbdr.find(qn(f'w:{side}'))
        if old is not None:
            pbdr.remove(old)
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '1')
        el.set(qn('w:color'), color)
        _insert_ordered(pbdr, el, f'w:{side}', order)


class PrintDoc:
    """A4 打印文档构建器。"""

    def __init__(self, footer='', theme=None, width=16.7,
                 margins=(1.6, 1.4, 2.15, 2.15)):
        self.t = dict(DEFAULT_THEME)
        if theme:
            self.t.update(theme)
        self.W = width
        self.doc = Document()

        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        top, bottom, left, right = margins
        sec.top_margin, sec.bottom_margin = Cm(top), Cm(bottom)
        sec.left_margin = sec.right_margin = Cm(left)
        sec.header_distance, sec.footer_distance = Cm(0.9), Cm(0.8)

        st = self.doc.styles['Normal']
        st.font.name = self.t['body']
        st.font.size = Pt(10)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), self.t['body'])
        st.paragraph_format.space_before = Pt(0)
        st.paragraph_format.space_after = Pt(0)

        if footer:
            fp = sec.footer.paragraphs[0]
            _fmt(fp, 0, 0, 11, align='center')
            self.text(fp, footer, size=7.5, color=self.t['light'])

    # ------------------------------------------------------------ 原语

    def p(self, before=0, after=0, line=16.0, align=None,
          left=0.0, hanging=0.0, right=0.0, keep=False):
        para = self.doc.add_paragraph()
        return _fmt(para, before, after, line, align, left, hanging, right, keep)

    def text(self, part, s, cn=None, en=None, size=10, bold=False,
             color=None, spacing=None):
        """往段落里加一段文字；cn 控制中文字体，en 控制西文字体。"""
        r = part.add_run(s)
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = RGBColor.from_string(color or self.t['ink'])
        cn = cn or self.t['body']
        en = en or cn
        rpr = r._element.get_or_add_rPr()
        rf = _put(rpr, 'w:rFonts', RPR)
        rf.set(qn('w:ascii'), en)
        rf.set(qn('w:hAnsi'), en)
        rf.set(qn('w:eastAsia'), cn)
        rf.set(qn('w:cs'), en)
        if spacing:
            _put(rpr, 'w:spacing', RPR, {'w:val': str(int(round(spacing * 20)))})
        return r

    def gap(self, pts, line=1.0):
        para = self.doc.add_paragraph()
        _fmt(para, 0, 0, max(line, 1.0))
        para.add_run("").font.size = Pt(1)
        para.paragraph_format.space_after = Pt(pts)
        return para

    def rule(self, color=None, sz=6, after=7, before=2, width=None,
             align='left', line=2):
        """横向装饰线。width 给 cm 时只画一段短线，按 align 靠左/居中/靠右。"""
        kw = {}
        if width is not None:
            if align == 'center':
                side = (self.W - width) / 2
                kw['left'] = side
                kw['right'] = side
            elif align == 'right':
                kw['left'] = self.W - width
            else:
                kw['right'] = self.W - width
        para = self.p(before, after, line, align=align, **kw)
        _border(para, {'bottom': (sz, color or self.t['line'])})
        para.add_run("").font.size = Pt(1)
        return para

    def pagebreak(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ------------------------------------------------------------ 组件

    def kicker(self, s, color=None, size=8, spacing=2.2, after=6, before=0):
        para = self.p(before, after, size + 3)
        self.text(para, s, cn=self.t['sans'], size=size,
                  color=color or self.t['accent'], spacing=spacing)
        return para

    def title(self, s, size=38, after=4, spacing=3):
        para = self.p(0, after, size * 1.32)
        self.text(para, s, cn=self.t['title'], size=size, bold=True,
                  color=self.t['ink'], spacing=spacing)
        return para

    def lede(self, s, size=12.5, after=10):
        para = self.p(0, after, size * 1.6)
        self.text(para, s, cn=self.t['sans'], size=size, color='46544F')
        return para

    def body(self, s, size=10, color=None, after=0, line=None, **kw):
        para = self.p(0, after, line or size * 1.65, **kw)
        self.text(para, s, size=size, color=color or self.t['grey'])
        return para

    def table(self, rows, cols, widths, margins=(0, 0, 0, 0)):
        """无边框固定宽表格，宽度单位 cm。"""
        if sum(widths) > self.W + 0.05:
            raise ValueError(f'表格总宽 {sum(widths):.2f}cm 超过版心 {self.W}cm')
        t = self.doc.add_table(rows=rows, cols=cols)
        tblPr = t._tbl.tblPr
        _put(tblPr, 'w:tblLayout', TBLPR, {'w:type': 'fixed'})
        tb = _put(tblPr, 'w:tblBorders', TBLPR)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'nil')
            tb.append(el)
        _put(tblPr, 'w:tblInd', TBLPR, {'w:w': '0', 'w:type': 'dxa'})
        for row in t.rows:
            _put(row._tr.get_or_add_trPr(), 'w:cantSplit', TRPR, {'w:val': '1'})
            for cell, w in zip(row.cells, widths):
                cell.width = Cm(w)
        for i, w in enumerate(widths):
            t.columns[i].width = Cm(w)
        mar = _put(tblPr, 'w:tblCellMar', TBLPR)
        for side, v in zip(('top', 'left', 'bottom', 'right'), margins):
            el = mar.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}')
                mar.append(el)
            el.set(qn('w:w'), str(int(v)))
            el.set(qn('w:type'), 'dxa')
        return t

    def shade(self, cell, fill=None):
        _put(cell._tc.get_or_add_tcPr(), 'w:shd', TCPR,
             {'w:val': 'clear', 'w:color': 'auto', 'w:fill': fill or self.t['band']})

    def vcenter(self, cell):
        _put(cell._tc.get_or_add_tcPr(), 'w:vAlign', TCPR, {'w:val': 'center'})

    def cell_para(self, cell, first=False, **kw):
        """在单元格里新建（或复用第一个）段落并设版式。"""
        para = cell.paragraphs[0] if first else cell.add_paragraph()
        return _fmt(para, **kw)

    def border(self, para, edges):
        _border(para, edges)

    def callout(self, build, pad=(230, 230, 300, 300), fill=None):
        """整宽淡色卡片；build(cell) 里往单元格添加内容。"""
        t = self.table(1, 1, [self.W], margins=pad)
        cell = t.cell(0, 0)
        self.shade(cell, fill)
        build(cell)
        return t

    def section_head(self, num, title, tag, size=24):
        """板块头：红色大数字 + 标题 + 英文小标签 + 细线。"""
        t = self.table(1, 2, [1.45, self.W - 1.45])
        nc, tc = t.cell(0, 0), t.cell(0, 1)
        self.vcenter(nc)
        self.vcenter(tc)
        pn = nc.paragraphs[0]
        _fmt(pn, 0, 0, size * 1.12)
        self.text(pn, num, cn=self.t['title'], size=size, bold=True,
                  color=self.t['accent'], spacing=0.5)
        pt_ = tc.paragraphs[0]
        _fmt(pt_, 0, 0, size * 0.79)
        self.text(pt_, title, cn=self.t['title'], size=size * 0.54, bold=True)
        self.text(pt_, "   " + tag, cn=self.t['sans'], size=7,
                  color=self.t['light'], spacing=1.2)
        self.rule(sz=6, before=2, after=7)

    def item(self, mark, head, body, body_after=7):
        """编号条目：① 标题 + 缩进一行的说明。"""
        para = self.p(0, 1.5, 15.5, left=0.8, hanging=0.8, keep=True)
        self.text(para, mark + "  ", cn=self.t['sans'], size=9.5,
                  bold=True, color=self.t['accent'])
        self.text(para, head, size=10.5, bold=True)
        p2 = self.p(0, body_after, 14.5, left=0.8)
        self.text(p2, body, size=9, color=self.t['light'])
        return para

    def timeline(self, steps, col_pad=200, tick=20):
        """steps = [(标签, [行1, 行2]), ...]；每列顶部一条短横线。"""
        n = len(steps)
        w = self.W / n
        t = self.table(2, n, [w] * n, margins=(0, 0, 0, col_pad))
        for i, (k, lines) in enumerate(steps):
            top, bot = t.cell(0, i), t.cell(1, i)
            pt_ = top.paragraphs[0]
            _fmt(pt_, 8, 0, 2, right=w - 0.9)
            _border(pt_, {'bottom': (tick, self.t['accent'] if i == n - 1
                                      else self.t['line'])})
            pt_.add_run("").font.size = Pt(1)
            pb = bot.paragraphs[0]
            _fmt(pb, 16, 5, 15)
            self.text(pb, k, size=10, bold=True)
            for ln in lines:
                pl = bot.add_paragraph()
                _fmt(pl, 0, 0, 15)
                self.text(pl, ln, size=9, color=self.t['light'])
        return t

    def quote(self, kicker, lines, tail='', size=15.5):
        """居中收尾金句。"""
        self.rule(color=self.t['ink'], sz=10, after=16, before=0,
                  width=1.5, align='center')
        para = self.p(0, 8, 17, align='center')
        self.text(para, kicker, size=9, color=self.t['light'], spacing=1.8)
        for i, ln in enumerate(lines):
            pp = self.p(10 if i == 0 else 0,
                        6 if i < len(lines) - 1 else 14,
                        size * 1.74, align='center')
            self.text(pp, ln, size=size, bold=True)
        if tail:
            tp = self.p(2, 0, 17, align='center')
            self.text(tp, tail, size=9.5, color=self.t['grey'])

    # ------------------------------------------------------------ 输出

    def save(self, path, title='', author='', comments=''):
        cp = self.doc.core_properties
        cp.title = title
        cp.author = author
        cp.comments = comments
        self.doc.save(path)
        return path
