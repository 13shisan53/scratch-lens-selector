# -*- coding: utf-8 -*-
"""生成公众号长文《0.08 mm 的划痕》Word 版。

用法：python article/build_article.py
输出：article/划痕检测-公众号长文.docx
"""

import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
IMG = os.path.join(ROOT, "scratch-lens-selector", "docs", "img")
OUT = os.path.join(HERE, "划痕检测-公众号长文.docx")

FONT = "微软雅黑"
INK = "1F2A30"        # 正文
MUTED = "6C7A82"      # 次要文字
ACCENT = "0B6E63"     # 主题色（深青）
ACCENT2 = "B26B00"    # 强调金


# ---------------------------------------------------------------- 基础样式
def set_font(run, size=11.5, bold=False, color=INK, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = FONT
    run.font.color.rgb = RGBColor.from_string(color)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


# OOXML 规定 pPr / tblPr 的子元素有固定顺序，追加到末尾会让 Word 报「需要修复」
_PPR_AFTER_PBDR = (
    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
_PPR_AFTER_SHD = _PPR_AFTER_PBDR[1:]
_TBLPR_AFTER_BORDERS = (
    "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
    "w:tblDescription", "w:tblPrChange",
)


def shade(paragraph, fill):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.insert_element_before(shd, *_PPR_AFTER_SHD)


def border(paragraph, side="left", color=ACCENT, sz=24, space=10):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    el = OxmlElement("w:" + side)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    pBdr.append(el)
    pPr.insert_element_before(pBdr, *_PPR_AFTER_PBDR)


def rich(paragraph, text, size=11.5, color=INK, bold_color=None):
    """支持 **加粗** 的简易行内解析。"""
    for i, seg in enumerate(text.split("**")):
        if not seg:
            continue
        r = paragraph.add_run(seg)
        on = i % 2 == 1
        set_font(r, size, bold=on, color=(bold_color or color) if on else color)


def body(text, size=11.5, after=9, before=0, color=INK):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.72
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rich(p, text, size=size, color=color, bold_color="0E1A20")
    return p


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after = Pt(9)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing = 1.3
    r = p.add_run("▍ " + text)
    set_font(r, 15.5, bold=True, color=ACCENT)
    return p


def kicker(text, color=MUTED):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_font(r, 9, color=color)
    return p


def quote(text, size=10.5):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.55)
    p.paragraph_format.right_indent = Cm(0.3)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(14)
    p.paragraph_format.line_spacing = 1.6
    border(p, "left", ACCENT, 24, 12)
    shade(p, "F2F7F7")
    rich(p, text, size=size, color="3A4A52", bold_color=ACCENT)
    return p


def bullet(text, size=11.5):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.75)
    p.paragraph_format.first_line_indent = Cm(-0.42)
    p.paragraph_format.line_spacing = 1.65
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run("·  ")
    set_font(r, size, bold=True, color=ACCENT)
    rich(p, text, size=size)
    return p


def divider(space=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space)
    p.paragraph_format.space_after = Pt(space)
    border(p, "bottom", "DCE3E7", 6, 1)
    return p


def cell_shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.insert_element_before(
        shd, "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
        "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
        "w:cellMerge", "w:tcPrChange")


def table(headers, rows, widths, font_size=9.5, head_fill=ACCENT):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "DCE3E7")
        borders.append(el)
    tblPr.insert_element_before(borders, *_TBLPR_AFTER_BORDERS)

    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.width = Cm(widths[i])
        cell_shade(c, head_fill)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(h)
        set_font(r, font_size, bold=True, color="FFFFFF")

    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].width = Cm(widths[i])
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.25
            rich(p, str(v), size=font_size, color="2A363D", bold_color="0E1A20")

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


def figure(name, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    path = os.path.join(IMG, name)
    if not os.path.exists(path):
        path = os.path.join(HERE, name)
    p.add_run().add_picture(path, width=Cm(width_cm))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after = Pt(14)
    r = c.add_run(caption)
    set_font(r, 9, color=MUTED)


def step(label, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r1 = p.add_run(label + "  ")
    set_font(r1, 11, bold=True, color=ACCENT2)
    r2 = p.add_run(text)
    set_font(r2, 11, bold=True, color="0E1A20")
    return p


def formula(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.5
    shade(p, "F5F7F8")
    r = p.add_run(text)
    set_font(r, 10.5, color="243139")
    r.font.name = "Consolas"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


# ---------------------------------------------------------------- 文档
doc = Document()

sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = Cm(2.2)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.6)
sec.right_margin = Cm(2.6)

normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = Pt(11.5)
normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
normal.paragraph_format.space_after = Pt(9)

# 页脚
foot = sec.footer.paragraphs[0]
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("划痕检测选型助手 · 1369 个海康真实型号 · 13shisan53.github.io/scratch-lens-selector")
set_font(fr, 8.5, color="9AA6AD")

# ---------------- 头部
kicker("工 业 视 觉 · 技 术 长 文")
tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
tp.paragraph_format.space_before = Pt(6)
tp.paragraph_format.space_after = Pt(10)
tp.paragraph_format.line_spacing = 1.32
tr = tp.add_run("0.08 mm 的划痕，\n逼我把整套选型逻辑重写了一遍")
set_font(tr, 24, bold=True, color="111B21")

sp = doc.add_paragraph()
sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
sp.paragraph_format.space_after = Pt(8)
sr = sp.add_run("从「大概选一下」到「算给你看」：3 个公式、1369 个真实型号、4 个差点翻车的坑")
set_font(sr, 11, color=MUTED)

mp = doc.add_paragraph()
mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
mp.paragraph_format.space_after = Pt(14)
mr = mp.add_run("预计阅读 9 分钟 · 文末有工具地址")
set_font(mr, 9, color="9AA6AD")

divider(6)

quote("一支 0.08 mm 宽的划痕，肉眼要凑近才看得清。\n"
      "但在产线上稳定地拍到它，只需要算清楚三件事：多少像素、多少焦距、什么角度打光。\n"
      "这篇文章把三件事全部写成公式，并用海康在售的 1369 条真实型号做了验证。", 11)

# ---------------- 一
h2("一、凌晨两点的那道划痕")

body("先说场景。")
body("铝件，精加工面，客户投诉表面有细划痕。质检拿放大镜找到了：长度不到 2 mm，**宽度 0.08 mm**。")
body("产线停着，工程师抱着相机拍了一晚上。图像里那道线，时有时无——换个角度就没了，挪一下灯又出来了。")
body("这不是相机不行。是从头到尾，没有人算过一件事：这条 0.08 mm 的线，到底需要多少像素才拍得到。")

# ---------------- 二
h2("二、这个行业最贵的三个习惯")

body("第一个：先买 500 万像素，不够再换。")
body("第二个：镜头先凑合，焦距靠试。")
body("第三个：打光越亮越好。")

body("单独看都不致命，连在一起就是灾难：")
body("相机买大了 → 像元变小 → 单像素进光量下降 → 只能把光打得更亮 → 反光糊掉划痕 → 换个角度 → 划痕又没了。")
body("于是加钱、加灯、加班。")
body("选型错误的真实成本，从来不是多花的那点钱，而是被试错占用的产线时间——停一夜，够买三台相机。")
body("真相是：**相机、镜头、光源不是三个独立零件，而是一条链。任何一环靠猜，后面全废。**")

# ---------------- 三
h2("三、能算的，就别猜")

body("把需求翻译成数字，只需要三步。")

step("STEP 1", "相机：到底需要多少像素")
formula("所需精度 = 最小划痕宽度 ÷ 缺陷占几像素\n"
        "需要像素 = 视场 ÷ 所需精度")
body("3 px 是工程上的稳妥值：2 px 勉强，3 px 推荐，5 px 严苛。为什么不是 2 px？因为缺陷边缘还要留出亚像素定位的余量，"
     "加上噪声和轻微离焦，2 px 的划痕在图像里往往只剩一条虚边。代入本例：")
formula("0.08 ÷ 3 = 0.0267 mm/px\n"
        "30 ÷ 0.0267 = 1125 px    →  长边短边都要过线\n"
        "最小可用：1280 × 1024（130 万像素）")
body("注意这里没有出现「500 万起步」这种话。像素是算出来的，不是买的。")

step("STEP 2", "镜头：到底需要多少焦距")
formula("靶面宽 = 分辨率 × 像元 = 1280 × 4.8 μm = 6.14 mm\n"
        "f ≈ 工作距离 × 靶面宽 ÷ (视场宽 + 靶面宽)\n"
        "  = 200 × 6.14 ÷ (30 + 6.14) ≈ 34 mm")
body("标准焦距里 35 mm 最接近，于是镜头定了。")
body("这里有个选型表上永远不会写、但决定你装上去能不能用的细节：35 mm 在 200 mm 处视场约 **35 mm**，比目标的 30 mm 大 17%。"
     "要么把工作距离缩到 171 mm，要么拍大一点、用 ROI 裁回来。")
quote("能算出来的误差，叫设计余量；算不出来的误差，叫现场调试。")

step("STEP 3", "光源：什么角度，而不是多亮")
body("划痕检测的主流打法是**暗场**：低角度环光掠射工件表面。")
body("平整表面把光反射出镜头，成像接近全黑；划痕是个微小的凹槽，把光散射进镜头，成像变亮。")
body("所以浅划痕的对比度，取决于**入射角**，而不是亮度。这也解释了为什么「把灯调亮」经常让划痕消失——亮度上去，背景反射也一起上去了。")
figure("principle.png", 15.2, "暗场划痕成像：低角度光掠射，平整面反射出镜头，划痕散射进镜头")

# ---------------- 四
h2("四、把官网「搬」下来：1369 条真实型号")

body("公式好写，难的是填参数。")
body("网上流传的选型表，最大的问题不是错，而是**不知道哪年的、哪来的、改没改过**。所以干脆直接抓海康机器人官网的产品页——"
     "用的是官网页面自己调用的接口，分页 + 并发，抓完再逐项核对字段。")

table(
    ["品类", "数量", "抓取的关键字段"],
    [
        ["工业面阵相机", "**247** 款", "分辨率、像元尺寸、靶面、最大帧率、数据接口、镜头接口、曝光时间"],
        ["FA 镜头", "**124** 款", "焦距、F 数、像面尺寸、接口类型"],
        ["远心 / 工业镜头", "**333** 款", "放大倍率、工作距离、像面尺寸、接口类型"],
        ["光源", "**665** 款", "类型、颜色、功率、控制器接口、外形尺寸、角度"],
        ["合计", "**1369** 条", "抓取时间 2026-10-26 22:57:55，可逐条回官网核对"],
    ],
    widths=[3.4, 2.4, 10.0],
)

body("相机那一栏还有一个隐藏工作量：官网面阵相机共 722 款，先筛出黑白、1.2–20 MP、GigE/USB3 的 247 款候选，再逐台抓详情参数。"
     "因为「像元尺寸」和「靶面」根本不在列表页里，而没有这两个数，焦距就算不出来。")

# ---------------- 五
h2("五、默认工况跑一遍")

body("参数：视场 30 × 25 mm｜最小划痕 0.08 mm｜缺陷占 3 px｜工作距离 150–250 mm｜产线速度 50 mm/s｜金属 + 线状划痕 + 平面。")

table(
    ["环节", "型号", "关键参数"],
    [
        ["相机", "**MV-CA013-20GM**", "1280 × 1024 · 4.8 μm · 1/2″ · GigE · 90 fps"],
        ["镜头", "**MVL-HF3524M-10MP**", "35 mm · F2.4 · 像面 Φ9 mm · C 口"],
        ["远心（备选）", "**MVL-MY-02-140**", "0.2× · 工作距离 140 mm · 像面 Φ11 mm"],
        ["光源", "**MV-LRDS-34-75-W**", "环形 · 白光 · 2.2 W · SMR-03V-B"],
    ],
    widths=[2.8, 5.6, 7.4],
)

body("跑完的校核长这样：")
bullet("实际精度 **0.023 mm/px**，0.08 mm 划痕约占 **3.4 px**，过 3 px 线 ✓")
bullet("带宽 **2.6 MB/s**，GigE 能力 118 MB/s，余量充足 ✓")
bullet("帧率需求 **2 fps**（50 mm/s ÷ 25 mm），相机 90 fps ✓")
bullet("镜头像面 Φ9 mm ≥ 靶面对角 Φ7.87 mm，四角不会暗 ✓")

figure("tool-hero.png", 15.5, "工具界面：左侧改参数，右侧即时给出型号与校核，下方可核对 1369 条原始数据")

# ---------------- 六
h2("六、从「能跑」到「敢用」：4 个坑")

body("写完第一版我以为结束了。真正花时间的，是下面这四件事。")

table(
    ["现象", "原因", "修正"],
    [
        ["工作距离 200 mm 显示成「2 mm」", "格式化函数把末位 0 全裁掉了", "只裁小数部分"],
        ["相机推荐出一个非 MV 系列型号", "没有品牌系列偏好", "优先海康自有 MV- 系列"],
        ["25 mm 视场配了 230 mm 外径环光", "只按类型选，没按尺寸筛", "超过视场短边 4 倍直接排除"],
        ["1/2″ 靶面配了 1.1″ 微距镜头", "只看焦距误差，不看像面富余", "像面过度富余加惩罚，改选 Φ9 mm"],
    ],
    widths=[5.4, 5.4, 5.0],
)

body("它们有一个共同点：**都不是崩溃，都是「看起来对的错答案」。**")
body("工具最危险的状态从来不是打不开，而是给你一个能直接执行的错误结论。")
quote("可信度不是在顺境里建立的，是在这四个坑里建立的。")

# ---------------- 七
h2("七、三个反常识")

body("**一、分辨率不是越高越好。**")
body("本例 130 万像素就够。硬上 500 万，价格、带宽、处理时间全涨，像元从 4.8 μm 缩到 3.45 μm，单像素进光量还下降——"
     "为了「更清楚」换来的，往往是「更暗、更慢、更贵」。选型的终点是「刚好够」，不是「越大越好」。")

body("**二、打光不是越亮越好。**")
body("暗场要的是角度。环光离工件越近、角度越贴表面，划痕越亮、背景越黑。所以工具会按视场短边去匹配环光外径："
     "给 25 mm 的视场配 230 mm 的环光，等于用大炮打蚊子，还打不准。")

body("**三、镜头不是「能装就行」。**")
body("像面必须 ≥ 靶面对角线，否则四角暗、边缘糊；但反过来，像面过度富余也是浪费——"
     "1/2″ 的靶面配 1.1″ 的微距镜头，多花的钱并不会让划痕更清楚。")

# ---------------- 八
h2("八、它不做什么")

bullet("只做几何估算。景深、畸变、相对照度、镜头 MTF、光源辐照度分布，都必须实拍验证。")
bullet("相机库只覆盖黑白、1.2–20 MP、GigE/USB3；线阵、10GigE、CoaXPress 需要扩展抓取范围。")
bullet("光源的「角度」以官网标注为准，不同系列口径可能不同，不能直接横比。")
bullet("型号与参数以海康官方最新规格书为准，工具不替代原厂选型支持。")

body("把这些写出来不是为了免责，是因为**一个不肯说边界的工具，不值得被信任**。")

# ---------------- 九
h2("九、开工前，先备齐这 5 个数")

body("这套方法能跑起来，前提是五个输入。缺一个，后面的计算就退化成猜。")

bullet("**① 视场**：不是零件尺寸，而是相机一次要覆盖的幅面（长 × 宽，mm）。宁可多留 10%，别卡边。")
bullet("**② 最小缺陷**：不是「精度要求 0.01 mm」这种话，而是「最细要拍到多细」——本例是 0.08 mm 的划痕宽度。")
bullet("**③ 工作距离**：由机构决定，先量再选。300 mm 的安装空间和 80 mm 的，是两套完全不同的方案。")
bullet("**④ 节拍 / 线速度**：决定帧率下限和曝光时间上限。速度翻倍，帧率也得翻倍。")
bullet("**⑤ 材质 + 缺陷类型 + 表面形状**：决定打光。金属拉丝面、玻璃、注塑件，没有通用解。")

body("顺序也不能反：**先相机，再镜头，最后光源。** 反过来做，等于让光学去迁就一个随便选的相机。")
quote("选型不是选零件，是解一道有约束的方程。约束没写全，答案就只是猜测。")

# ---------------- 十
h2("十、怎么用")

body("打开链接 → 改左边参数 → 右边出 BOM → 下方四个页签核对型号与参数。")
body("一个建议：**先定视场和最小缺陷，别先定相机。**这两个数决定了后面所有事。")
body("工具地址：**https://13shisan53.github.io/scratch-lens-selector/**（数据已内联，下载下来双击也能用）")

divider(14)

quote("选型的本质，不是从货架上挑一个「更好」的，\n而是先算清楚「够用」的边界在哪。\n算清楚了，钱才花得准。", 11.5)

body("如果你手上有具体工况——材质、划痕宽度、视场、节拍——把参数发我，我帮你跑一遍，把结果贴出来。", after=16)

# ---------------- 附：编辑用
divider(12)
kicker("— 编 辑 附 · 发 布 时 可 直 接 取 用 —")
h2("备选标题")

bullet("我把海康 1369 个真实型号翻了一遍，只为算清「这道划痕要多少钱的相机」")
bullet("0.08 mm 划痕拍不到？不是相机不行，是你没算过这三个数")
bullet("别再「大概选一下」了：划痕检测选型的 3 个公式和 4 个坑")

h2("摘要 / 朋友圈文案")
body("一支 0.08 mm 的划痕，让产线停了一夜。真正的问题不是相机不够贵，而是没人算过它需要多少像素。"
     "这篇把相机、镜头、光源的选型全部写成公式，用 1369 个海康在售真实型号验证，附可直接使用的在线工具。", size=10.5)

h2("配图建议")

bullet("头图：工具界面截图（screenshot.png），文案压在左侧深色区域")
bullet("文中图 1：暗场原理图（principle.png），放在「STEP 3 光源」之后")
bullet("文中图 2：1369 条数据表截图，放在第四章，强化「真实数据」")
bullet("结尾图：BOM 与校核截图，呼应「算给你看」")

doc.save(OUT)
print("saved:", OUT, os.path.getsize(OUT) // 1024, "KB")
