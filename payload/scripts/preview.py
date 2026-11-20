#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 docx 渲染成 PNG 预览，用来核对分页、溢出、重叠。

运行：python3 preview.py <x.docx> [outdir]
产出：outdir/preview.pdf、outdir/pg-1.png、outdir/pg-2.png ...

注意：这是 weasyprint 的近似渲染，不是 Word 的像素级输出。
字体、标点、加粗细节会有细微差别；页数与大致版式是可依赖的。
"""

import os
import subprocess
import sys

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
FONT_MAP = {'宋体': 'NotoSerifSC', '微软雅黑': 'NotoSansSC'}
FALLBACK = 'NotoSansSC'
FONT_FILES = {
    'NotoSansSC': '/root/.fonts/NotoSansSC.otf',
    'NotoSerifSC': '/root/.fonts/NotoSerifSC.otf',
}


def el(t):
    return '{%s}%s' % (W, t)


def wval(e, attr='w:val', default=None):
    return default if e is None else e.get(qn(attr), default)


def css_color(hexv):
    return '#' + hexv if hexv else None


def para_css(p):
    ppr = p.find(el('pPr'))
    if ppr is None:
        return 'margin:0;'
    styles = []
    sp = ppr.find(el('spacing'))
    before = after = 0
    line = None
    if sp is not None:
        if sp.get(qn('w:before')):
            before = int(sp.get(qn('w:before'))) / 20.0
        if sp.get(qn('w:after')):
            after = int(sp.get(qn('w:after'))) / 20.0
        if sp.get(qn('w:line')) and sp.get(qn('w:lineRule')) in ('exact', 'atLeast'):
            line = int(sp.get(qn('w:line'))) / 20.0
    styles.append(f'margin:{before}pt 0 {after}pt 0')
    if line:
        styles.append(f'line-height:{line}pt')

    ind = ppr.find(el('ind'))
    if ind is not None:
        left, right = ind.get(qn('w:left')), ind.get(qn('w:right'))
        hang, first = ind.get(qn('w:hanging')), ind.get(qn('w:firstLine'))
        if left:
            styles.append(f'margin-left:{int(left) / 567.0:.3f}cm')
        if right:
            styles.append(f'margin-right:{int(right) / 567.0:.3f}cm')
        if hang:
            styles.append(f'text-indent:-{int(hang) / 567.0:.3f}cm')
        elif first:
            styles.append(f'text-indent:{int(first) / 567.0:.3f}cm')

    jc = ppr.find(el('jc'))
    if jc is not None:
        styles.append('text-align:' + {
            'right': 'right', 'center': 'center', 'both': 'justify'
        }.get(wval(jc), 'left'))

    shd = ppr.find(el('shd'))
    if shd is not None and wval(shd, 'w:fill', 'auto') != 'auto':
        styles.append('background:' + css_color(wval(shd, 'w:fill')))

    bdr = ppr.find(el('pBdr'))
    if bdr is not None:
        for side in ('top', 'left', 'bottom', 'right'):
            b = bdr.find(el(side))
            if b is None:
                continue
            sz = int(wval(b, 'w:sz', '4'))
            col = css_color(wval(b, 'w:color', '000000'))
            space = int(wval(b, 'w:space', '0'))
            styles.append(f'border-{side}:{sz / 8:.2f}pt solid {col}')
            if space:
                pad = 'top' if side == 'top' else 'bottom' if side == 'bottom' else side
                styles.append(f'padding-{pad}:{space}pt')
    return ';'.join(styles)


def run_html(r):
    # 只取直属 w:t，用 iter 会重复计数（踩过的坑）
    t = ''.join((e.text or '') for e in r.iter(el('t')))
    t = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    if not t:
        return ''
    styles = []
    rpr = r.find(el('rPr'))
    fam = FALLBACK
    if rpr is not None:
        rf = rpr.find(el('rFonts'))
        if rf is not None:
            cn = rf.get(qn('w:eastAsia')) or rf.get(qn('w:ascii'))
            fam = FONT_MAP.get(cn, FALLBACK)
        sz = rpr.find(el('sz'))
        if sz is not None:
            styles.append(f'font-size:{int(wval(sz)) / 2:.1f}pt')
        b = rpr.find(el('b'))
        if b is not None and wval(b, 'w:val', '1') not in ('0', 'false'):
            styles.append('font-weight:700')
        col = rpr.find(el('color'))
        if col is not None and wval(col) not in (None, 'auto'):
            styles.append('color:' + css_color(wval(col)))
        sp = rpr.find(el('spacing'))
        if sp is not None:
            styles.append(f'letter-spacing:{int(wval(sp)) / 20:.2f}pt')
    styles.append(f'font-family:{fam}')
    return f'<span style="{";".join(styles)}">{t}</span>'


def para_html(p):
    inner = ''.join(run_html(r) for r in p.findall(el('r')))
    return f'<p style="{para_css(p)}">{inner or "&#160;"}</p>'


def table_html(tbl):
    widths = []
    grid = tbl.find(el('tblGrid'))
    if grid is not None:
        for gc in grid.findall(el('gridCol')):
            widths.append(int(gc.get(qn('w:w'))) / 567.0)
    mar = {'top': 0, 'bottom': 0, 'left': 0, 'right': 0}
    tblPr = tbl.find(el('tblPr'))
    if tblPr is not None:
        cm = tblPr.find(el('tblCellMar'))
        if cm is not None:
            for side in mar:
                e = cm.find(el(side))
                if e is not None:
                    mar[side] = int(wval(e, 'w:w', '0')) / 20.0
    rows = []
    for tr in tbl.findall(el('tr')):
        cells = []
        for tc in tr.findall(el('tc')):
            tcPr = tc.find(el('tcPr'))
            span, css = 1, []
            if tcPr is not None:
                gs = tcPr.find(el('gridSpan'))
                if gs is not None:
                    span = int(wval(gs))
                shd = tcPr.find(el('shd'))
                if shd is not None and wval(shd, 'w:fill', 'auto') != 'auto':
                    css.append('background:' + css_color(wval(shd, 'w:fill')))
                va = tcPr.find(el('vAlign'))
                if va is not None:
                    css.append('vertical-align:' + {'center': 'middle',
                                                    'bottom': 'bottom'}.get(wval(va), 'top'))
            inner = ''.join(para_html(p) for p in tc.findall(el('p')))
            cells.append(f'<td colspan="{span}" style="{";".join(css)}">{inner}</td>')
        rows.append('<tr>' + ''.join(cells) + '</tr>')
    cols = ''.join(f'<col style="width:{w:.3f}cm">' for w in widths)
    total = sum(widths) if widths else 16.7
    return (f'<table style="table-layout:fixed;width:{total:.3f}cm;'
            f'border-collapse:collapse;padding:0;margin:0">'
            f'<colgroup>{cols}</colgroup>{"".join(rows)}</table>'
            f'<div style="height:0;line-height:0;font-size:0">&nbsp;</div>')


def footer_css(docx_path):
    """把 word/footer1.xml 里的页脚文字画在 @bottom-center，预览才和打印一致。"""
    import re
    import zipfile
    try:
        z = zipfile.ZipFile(docx_path)
    except Exception:
        return ''
    names = sorted(n for n in z.namelist()
                   if re.match(r'word/footer\d*\.xml$', n))
    if not names:
        return ''
    x = z.read(names[0]).decode('utf-8')
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', x)).strip()
    if not text:
        return ''
    m = re.search(r'<w:sz w:val="(\d+)"', x)
    c = re.search(r'<w:color w:val="([0-9A-Fa-f]{6})"', x)
    sz = int(m.group(1)) / 2 if m else 8
    col = '#' + c.group(1) if c else '#8A918C'
    return (f'@page {{ @bottom-center {{ content: "{text}";'
            f' font-family: NotoSerifSC; font-size: {sz}pt; color: {col}; }} }}')


def render_html(docx_path):
    d = Document(docx_path)
    parts = []
    for child in d.element.body:
        tag = etree.QName(child).localname
        if tag == 'p':
            br = child.find('.//' + el('br'))
            if br is not None and br.get(qn('w:type')) == 'page':
                parts.append('<div style="break-after:page;height:0"></div>')
                continue
            parts.append(para_html(child))
        elif tag == 'tbl':
            parts.append(table_html(child))

    faces = ''.join(
        f'@font-face {{ font-family: {name}; src: url(file://{path}); }}'
        for name, path in FONT_FILES.items() if os.path.exists(path)
    )
    return f"""<!doctype html><html><head><meta charset="utf-8"><style>
{faces}
@page {{ size: A4; margin: 16mm 21.5mm 14mm 21.5mm; }}
{footer_css(docx_path)}
html, body {{ margin:0; padding:0; }}
body {{ font-family: NotoSerifSC; font-size:10pt; color:#1B3B33; }}
p {{ margin:0; }}
table {{ border:0; }}
td {{ padding: 0; }}
</style></head><body>{''.join(parts)}</body></html>"""


def main(docx_path, outdir='.'):
    os.makedirs(outdir, exist_ok=True)
    html = render_html(docx_path)
    html_path = os.path.join(outdir, 'preview.html')
    pdf_path = os.path.join(outdir, 'preview.pdf')
    open(html_path, 'w', encoding='utf-8').write(html)

    from weasyprint import HTML
    HTML(string=html, base_url=outdir + '/').write_pdf(pdf_path)

    for f in os.listdir(outdir):
        if f.startswith('pg-') and f.endswith('.png'):
            os.remove(os.path.join(outdir, f))
    subprocess.run(['pdftoppm', '-r', '130', '-png', pdf_path,
                    os.path.join(outdir, 'pg')], check=True)
    pages = sorted(f for f in os.listdir(outdir) if f.startswith('pg-'))
    print(f'{len(pages)} page(s): {", ".join(pages)}')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit('usage: python3 preview.py <x.docx> [outdir]')
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else 'preview')
