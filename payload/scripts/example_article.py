#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""示例：把一篇公众号文章排成 A4 两页打印版。

运行：python3 example_article.py [out.docx]

这份示例同时是排版的「风格说明书」：想做别的题材，复制它、换掉文案即可。
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from printdoc import PrintDoc

OUT = sys.argv[1] if len(sys.argv) > 1 else 'example.docx'

d = PrintDoc(footer="《三三原则》精华笔记　·　整理自公众号「萌萌妈妈成长日记」")
ACCENT, GREY, LIGHT, LINE = d.t['accent'], d.t['grey'], d.t['light'], d.t['line']

# ------------------------------------------------------------ 第 1 页

d.kicker("家庭学习法 · 一页读懂", after=7)
d.title("三三原则", size=38, after=4)
d.lede("每晚九个固定动作，两个孩子连续三年轮流第一", after=10)
d.rule(color=ACCENT, sz=16, width=3.4, before=0, after=14)

p = d.p(0, 11, 18.5)
d.text(p, "孩子学习出问题，只会出在三个地方——", size=11, color=GREY)
d.text(p, "作业、状态、错题", size=11, bold=True)
d.text(p, "。每个环节固定三个动作，一共九个，", size=11, color=GREY)
d.text(p, "孩子照着走就行", size=11, bold=True)
d.text(p, "。", size=11, color=GREY)


def principle(cell):
    p0 = d.cell_para(cell, first=True, after=5, line=15, left=0.42)
    d.border(p0, {'left': (20, ACCENT)})
    d.text(p0, "为什么有效：孩子缺的不是努力，是「做决定」",
           cn=d.t['sans'], size=10, bold=True)
    p1 = d.cell_para(cell, after=0, line=15, left=0.42)
    d.text(p1, "先写哪科、卡住了怎么办、错了要不要管——每个决定都在消耗他，"
              "决定做多了，人就瘫了。三三原则把所有决定提前做完，晚上只剩下执行。",
           size=9.5, color=GREY)


d.callout(principle)
d.gap(16)

d.section_head("01", "作业 · 把「开始」变简单", "WORK")
d.item("①", "固定时间、固定动作开始",
       "到家洗手、喝水、坐下，雷打不动——身体记住流程，就不用每天重新下决心。")
d.item("②", "动笔前给每科估个时间", "写在旁边。看得见终点，才跑得起速度。")
d.item("③", "卡住的题画圈跳过", "最后集中解决。死磕一道题，赔上的是整晚。")

d.gap(7)
d.section_head("02", "状态 · 让专注撑完全场", "FOCUS")
d.item("①", "每科之间休息两分钟", "喝水、望远——大脑清空缓存，下一科才装得进。")
d.item("②", "休息不碰任何屏幕", "屏幕不是休息，是把注意力拐走。")
d.item("③", "书桌只用来学习", "玩去别的地方——地方有了分工，孩子坐下就知道该干什么。")

d.gap(7)
d.section_head("03", "错题 · 不让同一个坑绊倒两次", "REVIEW")
d.item("①", "当天的错题当天弄懂", "不过夜。账不过夜，才不会滚成债。")
d.item("②", "弄懂的标准是「讲出来」", "能把关键一步讲清楚才算会；只改答案，换个问法照样错。")
d.item("③", "周末只重做错过两次以上的题", "力气要花在最常摔的那个坑上。")

# ------------------------------------------------------------ 第 2 页

d.pagebreak()
d.gap(42)

p = d.p(0, 5, 13)
d.text(p, "三年之后", size=9, bold=True, color=ACCENT, spacing=2.6)
d.text(p, "　·　效果不是立刻的，是慢慢长出来的", size=8.5, color=LIGHT, spacing=0.4)
d.rule(color=LINE, sz=8, before=10, after=30)

d.timeline([
    ("第 1 个月", ["变化只是", "「不用催了」"]),
    ("3 个月", ["作业从两个小时", "缩到一个多小时"]),
    ("1 年", ["错题本", "越翻越薄"]),
    ("第 3 年", ["俩娃轮流第一", "家长很久没进书房"]),
])

d.gap(68)


def reading(cell):
    p = d.cell_para(cell, first=True, after=5, line=13, left=0.1)
    d.text(p, "三三原则管不到的一块", size=8, bold=True, color=ACCENT, spacing=1.6)
    p = d.cell_para(cell, after=6, line=28, left=0.1)
    d.text(p, "每天 15 分钟精读，补齐「理解力」",
           cn=d.t['title'], size=14, bold=True)
    p = d.cell_para(cell, after=10, line=16.5, left=0.1)
    d.text(p, "管住了作业、状态和错题，语文仍有一块管不着：理解力。"
              "读不懂题的孩子，三年级以后连数学应用题都吃亏——"
              "它只能从阅读里长出来，没法外包，也没法突击。",
           size=9, color=GREY)
    for mark, head, body in [
        ("①", "合上书复述讲了什么", "讲不顺的地方，就是没读懂的地方"),
        ("②", "看作者先写什么、后写什么", "结构看明白，题目怎么变都有方向"),
        ("③", "挑一句最打动人的，说出好在哪", "读得出门道，笔下才有办法"),
    ]:
        p = d.cell_para(cell, after=5, line=16, left=0.6, hanging=0.6)
        d.text(p, mark + "  ", cn=d.t['sans'], size=9, bold=True, color=ACCENT)
        d.text(p, head, size=10.5, bold=True)
        d.text(p, "　" + body + "。", size=9, color=LIGHT)
    p = d.cell_para(cell, before=10, after=4, line=16.5, left=0.1)
    d.border(p, {'top': (4, LINE)})
    d.text(p, "一天 15 分钟，一年就是 90 多个小时。半年下来，老大语文没下过 90 分；"
              "老二做阅读题，每个答案都说得出从哪段找到的依据。",
           size=9, color=GREY)


d.callout(reading, pad=(300, 300, 360, 360))

d.gap(112)
d.quote("所谓自推，不过如此",
        ["把决定变成流程，把流程变成习惯，", "最后孩子自己往前走。"],
        "三三原则管的是作业，养成的却是「自己的事自己心里有数」。",
        size=15.5)

d.save(OUT,
       title="三三原则 · 学习法精华整理",
       author="整理自 萌萌妈妈成长日记",
       comments="每晚九个固定动作：作业、状态、错题；另加每天 15 分钟精读")
print("saved", OUT)
