#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验 docx 的 OOXML 子元素顺序，防止 Word 打开时报「内容有问题」。

运行：python3 check_docx.py <x.docx>
退出码 0 = 通过；1 = 有顺序问题。
"""

import sys
import zipfile

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

ORDER = {
    'pPr': ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
            'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
            'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
            'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
            'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
            'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
            'textDirection', 'textAlignment', 'textboxTightWrap',
            'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'],
    'rPr': ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
            'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
            'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
            'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
            'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
            'eastAsianLayout', 'specVanish', 'oMath'],
    'tblPr': ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual',
              'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
              'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout',
              'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription'],
    'tcPr': ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders',
             'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign',
             'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge',
             'tcPrChange'],
    'trPr': ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore',
             'wAfter', 'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing',
             'jc', 'hidden'],
}


def check(path):
    root = etree.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    problems = []
    for tag, order in ORDER.items():
        for e in root.iter(W + tag):
            names = [etree.QName(c).localname for c in e]
            idx = [order.index(n) for n in names if n in order]
            unknown = [n for n in names if n not in order]
            if idx != sorted(idx) or unknown:
                parent = etree.QName(e.getparent()).localname
                problems.append(f'<{parent}> 下的 <{tag}> 顺序异常: {names}')
    return problems


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit('usage: python3 check_docx.py <x.docx>')
    bad = check(sys.argv[1])
    if bad:
        print(f'{len(bad)} problem(s):')
        for b in bad[:20]:
            print(' -', b)
        sys.exit(1)
    print('ok: 0 schema order problems')
