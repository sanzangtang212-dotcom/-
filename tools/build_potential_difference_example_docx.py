from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

OUT = Path('downloads/一题三变加高考真题_第十章_电势差_出版规范版.docx')
OUT.parent.mkdir(parents=True, exist_ok=True)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


# ============================================================
# 普通 Word 文本：正文中文宋体，标题黑体，物理量斜体，数字和单位正体
# ============================================================
def wrun(text, font='宋体', size=24, bold=False, italic=False,
         sub=False, sup=False, color=None):
    text = escape(str(text))
    props = [
        f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>',
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    ]
    if bold:
        props.append('<w:b/>')
    if italic:
        props.append('<w:i/>')
    if sub:
        props.append('<w:vertAlign w:val="subscript"/>')
    if sup:
        props.append('<w:vertAlign w:val="superscript"/>')
    if color:
        props.append(f'<w:color w:val="{color}"/>')
    return f'<w:r><w:rPr>{"".join(props)}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def cn(text, size=24, bold=False):
    return wrun(text, font='宋体', size=size, bold=bold)


def hei(text, size=28):
    return wrun(text, font='黑体', size=size, bold=True)


def roman(text, size=24, sub=False, sup=False):
    return wrun(text, font='Times New Roman', size=size, sub=sub, sup=sup)


def phys(text, size=24):
    return wrun(text, font='Times New Roman', size=size, italic=True)


def mono(text, size=22):
    return wrun(text, font='Courier New', size=size)


def Ut(subscript, size=24):
    return phys('U', size=size) + roman(subscript, size=size, sub=True)


def Wt(subscript, size=24):
    return phys('W', size=size) + roman(subscript, size=size, sub=True)


def phit(subscript, size=24):
    return phys('φ', size=size) + roman(subscript, size=size, sub=True)


def Ept(subscript=None, size=24):
    out = phys('E', size=size) + roman('p', size=size, sub=True)
    if subscript:
        out += roman(subscript, size=size, sub=True)
    return out


def qt(subscript=None, size=24):
    out = phys('q', size=size)
    if subscript:
        out += roman(subscript, size=size, sub=True)
    return out


def sci_text(coef, exponent, unit_text=None, size=24):
    out = roman(f'{coef}×10', size=size)
    out += roman(str(exponent).replace('-', '−'), size=size, sup=True)
    if unit_text:
        out += roman(' ' + unit_text, size=size)
    return out


def paragraph(inner='', align=None, before=0, after=90, line=360,
              page_break=False, keep_next=False, first_line=0):
    ppr = []
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    if page_break:
        ppr.append('<w:pageBreakBefore/>')
    if keep_next:
        ppr.append('<w:keepNext/>')
    if first_line:
        ppr.append(f'<w:ind w:firstLine="{first_line}"/>')
    ppr.append(
        f'<w:spacing w:before="{before}" w:after="{after}" '
        f'w:line="{line}" w:lineRule="auto"/>'
    )
    return f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>{inner}</w:p>'


def body(*runs, indent=True, after=70):
    return paragraph(''.join(runs), first_line=480 if indent else 0, after=after)


def plain(text, indent=True, after=70):
    return body(cn(text), indent=indent, after=after)


def label(text):
    return paragraph(hei(text, size=24), before=110, after=45, line=320, keep_next=True)


def heading(text, page_break=False):
    return paragraph(hei(text, size=28), before=180, after=100, line=340,
                     page_break=page_break, keep_next=True)


# ============================================================
# OMML 公式：所有分式、上下角标均为 Word 可编辑结构
# ============================================================
def mr(text, upright=False, font='Times New Roman', preserve=False):
    text = escape(str(text))
    sty = 'p' if upright else 'i'
    space = ' xml:space="preserve"' if preserve else ''
    return (
        '<m:r>'
        f'<m:rPr><m:sty m:val="{sty}"/></m:rPr>'
        f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        f'<m:t{space}>{text}</m:t>'
        '</m:r>'
    )


def mv(text):
    return mr(text, upright=False)


def mu(text, preserve=False):
    return mr(text, upright=True, preserve=preserve)


def mn(text):
    return mr(text, upright=True)


def mop(text):
    return mr(text, upright=True)


def msub(base, subscript):
    return f'<m:sSub><m:sSubPr/><m:e>{base}</m:e><m:sub>{subscript}</m:sub></m:sSub>'


def msup(base, exponent):
    return f'<m:sSup><m:sSupPr/><m:e>{base}</m:e><m:sup>{exponent}</m:sup></m:sSup>'


def mfrac(numerator, denominator):
    return f'<m:f><m:fPr/><m:num>{numerator}</m:num><m:den>{denominator}</m:den></m:f>'


def mparen(expr):
    return f'<m:d><m:dPr/><m:e>{expr}</m:e></m:d>'


def eq(expr, after=80):
    return paragraph(
        f'<m:oMathPara><m:oMath>{expr}</m:oMath></m:oMathPara>',
        align='center', before=35, after=after, line=300
    )


def U(sub):
    return msub(mv('U'), mu(sub))


def WW(sub):
    return msub(mv('W'), mu(sub))


def phi(sub):
    return msub(mv('φ'), mu(sub))


def q(sub=None):
    return msub(mv('q'), mu(sub)) if sub else mv('q')


def Ep(sub=None):
    base = msub(mv('E'), mu('p'))
    return msub(base, mu(sub)) if sub else base


def power10(exponent):
    return msup(mn('10'), mu(str(exponent).replace('-', '−')))


def sci(coef, exponent):
    return mn(coef) + mop('×') + power10(exponent)


def ratio(numerator, denominator):
    return mfrac(numerator, denominator)


def unit(text):
    return mu(' ' + text, preserve=True)


def abs_expr(expr):
    return mop('|') + expr + mop('|')


# ============================================================
# 真数据表格
# ============================================================
def table_cell(inner, width, header=False):
    shade = '<w:shd w:val="clear" w:fill="DCE6F1"/>' if header else ''
    tcpr = f'<w:tcW w:w="{width}" w:type="dxa"/>{shade}<w:vAlign w:val="center"/>'
    return f'<w:tc><w:tcPr>{tcpr}</w:tcPr>{paragraph(inner, align="center", after=0, line=300)}</w:tc>'


def table(rows, widths):
    borders = (
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="8" w:color="666666"/>'
        '<w:left w:val="single" w:sz="8" w:color="666666"/>'
        '<w:bottom w:val="single" w:sz="8" w:color="666666"/>'
        '<w:right w:val="single" w:sz="8" w:color="666666"/>'
        '<w:insideH w:val="single" w:sz="6" w:color="999999"/>'
        '<w:insideV w:val="single" w:sz="6" w:color="999999"/>'
        '</w:tblBorders>'
    )
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    trs = []
    for i, row in enumerate(rows):
        cells = [table_cell(row[j], widths[j], header=(i == 0)) for j in range(len(row))]
        trs.append(f'<w:tr>{"".join(cells)}</w:tr>')
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        f'<w:jc w:val="center"/>{borders}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>'
        f'{"".join(trs)}</w:tbl>'
    )


parts = []
parts.append(paragraph(hei('第十章  静电场中的能量', size=34), align='center', after=55))
parts.append(paragraph(hei('2. 电势差', size=30), align='center', after=75))
parts.append(paragraph(hei('例题·一题三变＋高考真题', size=28), align='center', after=200))

# ------------------------------------------------------------
# 原例题解析
# ------------------------------------------------------------
parts.append(heading('原例题解析'))
parts.append(body(
    cn('已知正点电荷电荷量 '), qt(), roman('='), sci_text('2.0', '-9', 'C'),
    cn('，从 '), roman('A'), cn(' 点移动到 '), roman('B'), cn(' 点时静电力做功 '),
    Wt('AB'), roman('='), sci_text('1.6', '-7', 'J'), cn('；从 '), roman('B'),
    cn(' 点移动到 '), roman('C'), cn(' 点时静电力做功 '), Wt('BC'), roman('='),
    sci_text('−4.0', '-7', 'J'), cn('。')
))
parts.append(plain('根据静电力做功与电势差的关系可得'))
parts.append(eq(U('AB') + mop('=') + ratio(WW('AB'), q()) + mop('=') +
                ratio(sci('1.6', '-7'), sci('2.0', '-9')) + unit('V') +
                mop('=') + mn('80') + unit('V')))
parts.append(eq(U('BC') + mop('=') + ratio(WW('BC'), q()) + mop('=') +
                ratio(mop('−') + sci('4.0', '-7'), sci('2.0', '-9')) + unit('V') +
                mop('=') + mop('−') + mn('200') + unit('V')))
parts.append(plain('由UAB=φA−φB=80 V可知φA>φB；由UBC=φB−φC=−200 V可知φC>φB。取φB=0，可得'))
parts.append(eq(phi('A') + mop('=') + mn('80') + unit('V') + mop('，') +
                phi('B') + mop('=') + mn('0') + unit('V') + mop('，') +
                phi('C') + mop('=') + mn('200') + unit('V')))
parts.append(plain('因此，C点电势最高，B点电势最低，三点电势关系为φC>φA>φB。'))
parts.append(plain('根据电势差的定义可得'))
parts.append(eq(U('AC') + mop('=') + phi('A') + mop('−') + phi('C') +
                mop('=') + mop('−') + mn('120') + unit('V')))
parts.append(plain('把电荷量为1.5×10的−9次方C的正点电荷从A点移动到C点，根据W=qU可得'))
parts.append(eq(WW('AC') + mop('=') + q() + U('AC') + mop('=') +
                mparen(sci('1.5', '-9') + mop('×') + mparen(mop('−') + mn('120'))) +
                unit('J') + mop('=') + mop('−') + sci('1.8', '-7') + unit('J')))
parts.append(plain('在匀强电场中，沿电场方向电势降低。可将C、A、B三点置于同一条电场线上，电场方向由C指向B。'))
parts.append(paragraph(mono('C（200 V）────────A（80 V）─────B（0 V）       E →'), align='center', after=45, line=300))
parts.append(plain('由于匀强电场中电势差与沿电场方向的距离成正比，所以CA∶AB=120∶80=3∶2。该图只表示一种可能的位置关系。'))

# ------------------------------------------------------------
# 基础变式
# ------------------------------------------------------------
parts.append(heading('一、基础变式——同模型巩固', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('在匀强电场中，把电荷量为 '), sci_text('4.0', '-9', 'C'),
    cn(' 的正点电荷从 '), roman('A'), cn(' 点移动到 '), roman('B'),
    cn(' 点，静电力做功为 '), sci_text('1.2', '-7', 'J'), cn('；再从 '),
    roman('B'), cn(' 点移动到 '), roman('C'), cn(' 点，静电力做功为 '),
    sci_text('−2.0', '-7', 'J'), cn('。')
))
parts.append(plain('（1）比较A、B、C三点电势的高低。'))
parts.append(plain('（2）求UAB、UBC和UAC。'))
parts.append(body(cn('（3）把电荷量为 '), sci_text('3.0', '-9', 'C'), cn(' 的正点电荷从A点移动到C点，求静电力做的功。'))
parts.append(plain('（4）若A、B、C三点位于同一条电场线上，画出一种可能的位置关系，并求CA与AB的距离之比。'))
parts.append(label('【答案】'))
parts.append(plain('（1）φC>φA>φB；（2）UAB=30 V，UBC=−50 V，UAC=−20 V；（3）−6.0×10的−8次方J；（4）电场方向由C指向B，CA∶AB=2∶3。'))
parts.append(label('【解析】'))
parts.append(plain('根据静电力做功与电势差的关系可得'))
parts.append(eq(U('AB') + mop('=') + ratio(sci('1.2', '-7'), sci('4.0', '-9')) + unit('V') +
                mop('=') + mn('30') + unit('V')))
parts.append(eq(U('BC') + mop('=') + ratio(mop('−') + sci('2.0', '-7'), sci('4.0', '-9')) + unit('V') +
                mop('=') + mop('−') + mn('50') + unit('V')))
parts.append(plain('取φB=0，可得φA=30 V、φC=50 V，因此φC>φA>φB。'))
parts.append(plain('根据电势差的可加性可得'))
parts.append(eq(U('AC') + mop('=') + U('AB') + mop('+') + U('BC') +
                mop('=') + mparen(mn('30') + mop('−') + mn('50')) + unit('V') +
                mop('=') + mop('−') + mn('20') + unit('V')))
parts.append(plain('根据W=qU可得'))
parts.append(eq(WW('AC') + mop('=') + mparen(sci('3.0', '-9') + mop('×') + mparen(mop('−') + mn('20'))) +
                unit('J') + mop('=') + mop('−') + sci('6.0', '-8') + unit('J')))
parts.append(plain('在匀强电场中沿电场方向电势降低，一种可能的位置关系为C—A—B，电场方向由C指向B。'))
parts.append(paragraph(mono('C（50 V）────A（30 V）──────B（0 V）       E →'), align='center', after=45, line=300))
parts.append(plain('由于φC−φA=20 V，φA−φB=30 V，所以CA∶AB=20∶30=2∶3。'))

# ------------------------------------------------------------
# 情境变式：负电荷符号迁移
# ------------------------------------------------------------
parts.append(heading('二、情境变式——电子束标定中的符号判断', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('某电子束标定装置中，一个电子的电荷量为 '), roman('−'), sci_text('1.60', '-19', 'C'),
    cn('。电子从 '), roman('A'), cn(' 点移动到 '), roman('B'), cn(' 点时，静电力做功为 '),
    sci_text('3.2', '-18', 'J'), cn('；从 '), roman('B'), cn(' 点移动到 '), roman('C'),
    cn(' 点时，静电力做功为 '), sci_text('−1.6', '-18', 'J'), cn('。')
))
parts.append(plain('（1）求UAB、UBC和UAC。'))
parts.append(plain('（2）比较A、B、C三点电势的高低。'))
parts.append(plain('（3）求电子从A点移动到C点时静电力做的功，并说明能否直接把两段功相加。'))
parts.append(plain('（4）若把电子换成电荷量为+1.60×10的−19次方C的质子，从A点移动到C点，静电力做功为多少？'))
parts.append(label('【答案】'))
parts.append(plain('（1）UAB=−20 V，UBC=10 V，UAC=−10 V；（2）φB>φC>φA；（3）1.6×10的−18次方J，可以直接把同一电荷沿连续路径两段功相加；（4）−1.6×10的−18次方J。'))
parts.append(label('【解析】'))
parts.append(plain('电子带负电，应用U=W/q时必须把电荷量的负号代入。'))
parts.append(eq(U('AB') + mop('=') + ratio(sci('3.2', '-18'), mop('−') + sci('1.60', '-19')) + unit('V') +
                mop('=') + mop('−') + mn('20') + unit('V')))
parts.append(eq(U('BC') + mop('=') + ratio(mop('−') + sci('1.6', '-18'), mop('−') + sci('1.60', '-19')) + unit('V') +
                mop('=') + mn('10') + unit('V')))
parts.append(eq(U('AC') + mop('=') + U('AB') + mop('+') + U('BC') +
                mop('=') + mop('−') + mn('10') + unit('V')))
parts.append(plain('取φA=0，可得φB=20 V、φC=10 V，因此φB>φC>φA。'))
parts.append(plain('同一电子沿A→B→C连续移动，静电力总功等于两段功的代数和。'))
parts.append(eq(WW('AC') + mop('=') + WW('AB') + mop('+') + WW('BC') +
                mop('=') + mparen(sci('3.2', '-18') + mop('−') + sci('1.6', '-18')) + unit('J') +
                mop('=') + sci('1.6', '-18') + unit('J')))
parts.append(plain('换成正电荷后，A、C两点间的电势差不变，但电荷量符号改变，因此静电力做功变号。'))
parts.append(eq(WW('AC') + mop('=') + mparen(sci('1.60', '-19') + mop('×') + mparen(mop('−') + mn('10'))) +
                unit('J') + mop('=') + mop('−') + sci('1.6', '-18') + unit('J')))

# ------------------------------------------------------------
# 素养提升：数据表、图像、匀强电场一致性
# ------------------------------------------------------------
parts.append(heading('三、素养提升——实验数据与电势—位置图像', page_break=True))
parts.append(label('【题目】'))
parts.append(plain('某实验小组在同一匀强电场中，分别选取不同试探电荷测量A、B、C三点间的静电力做功，数据如下表。A、B、C位于同一条电场线上，AB=0.30 m，BC=0.20 m。'))
parts.append(table([
    [hei('移动过程', size=22), hei('试探电荷量 q', size=22), hei('静电力做功 W', size=22), hei('电势差 U', size=22)],
    [roman('A→B', size=22), roman('+2.0×10', size=22) + roman('−9', size=22, sup=True) + roman(' C', size=22), roman('+1.2×10', size=22) + roman('−7', size=22, sup=True) + roman(' J', size=22), cn('待求', size=22)],
    [roman('B→C', size=22), roman('−4.0×10', size=22) + roman('−9', size=22, sup=True) + roman(' C', size=22), roman('−1.6×10', size=22) + roman('−7', size=22, sup=True) + roman(' J', size=22), cn('待求', size=22)],
    [roman('A→C', size=22), roman('+3.0×10', size=22) + roman('−9', size=22, sup=True) + roman(' C', size=22), roman('+3.0×10', size=22) + roman('−7', size=22, sup=True) + roman(' J', size=22), cn('待求', size=22)],
], [1900, 2600, 2600, 1600]))
parts.append(plain('（1）完成表格中的电势差，并验证UAC=UAB+UBC。'))
parts.append(plain('（2）取C点电势为0，求A、B两点的电势。'))
parts.append(plain('（3）分别利用AB段和BC段数据计算电场强度，判断实验数据是否与匀强电场相符。'))
parts.append(plain('（4）以A点为x=0、沿A→C方向建立x轴，定性画出φ-x图像，并写出A、B、C三点坐标。'))
parts.append(body(cn('（5）预测电荷量为 '), sci_text('5.0', '-9', 'C'), cn(' 的正点电荷从A点移动到C点时静电力做的功。'))
parts.append(label('【答案】'))
parts.append(plain('（1）UAB=60 V，UBC=40 V，UAC=100 V，且100 V=60 V+40 V；（2）φA=100 V，φB=40 V；（3）两段均得E=200 V/m，数据与匀强电场相符；（4）图像为从(0，100 V)到(0.50 m，0)的下降直线，B点坐标为(0.30 m，40 V)；（5）5.0×10的−7次方J。'))
parts.append(label('【解析】'))
parts.append(plain('根据U=W/q，逐行计算可得'))
parts.append(eq(U('AB') + mop('=') + ratio(sci('1.2', '-7'), sci('2.0', '-9')) + unit('V') + mop('=') + mn('60') + unit('V')))
parts.append(eq(U('BC') + mop('=') + ratio(mop('−') + sci('1.6', '-7'), mop('−') + sci('4.0', '-9')) + unit('V') + mop('=') + mn('40') + unit('V')))
parts.append(eq(U('AC') + mop('=') + ratio(sci('3.0', '-7'), sci('3.0', '-9')) + unit('V') + mop('=') + mn('100') + unit('V')))
parts.append(plain('三点间电势差满足可加性：UAC=UAB+UBC=100 V。'))
parts.append(plain('取φC=0，由UBC=φB−φC=40 V可得φB=40 V；再由UAB=φA−φB=60 V可得φA=100 V。'))
parts.append(plain('根据匀强电场中U=Ed可得'))
parts.append(eq(mv('E') + mop('=') + ratio(U('AB'), mn('0.30') + unit('m')) +
                mop('=') + mn('200') + unit('V/m')))
parts.append(eq(mv('E') + mop('=') + ratio(U('BC'), mn('0.20') + unit('m')) +
                mop('=') + mn('200') + unit('V/m')))
parts.append(plain('两段计算得到的电场强度相同，因此实验数据与匀强电场模型相符。'))
parts.append(paragraph(mono('φ/V'), align='center', after=10, line=260))
parts.append(paragraph(mono('100 ● A'), align='center', after=0, line=260))
parts.append(paragraph(mono('     ╲'), align='center', after=0, line=260))
parts.append(paragraph(mono(' 40      ● B'), align='center', after=0, line=260))
parts.append(paragraph(mono('          ╲'), align='center', after=0, line=260))
parts.append(paragraph(mono('  0            ● C──────── x/m'), align='center', after=45, line=260))
parts.append(plain('对应坐标为A(0，100 V)、B(0.30 m，40 V)、C(0.50 m，0)。'))
parts.append(plain('预测新电荷的静电力做功时，只需使用由电场本身决定的UAC=100 V。'))
parts.append(eq(WW('AC') + mop('=') + mparen(sci('5.0', '-9') + mop('×') + mn('100')) + unit('J') +
                mop('=') + sci('5.0', '-7') + unit('J')))

# ------------------------------------------------------------
# 高考真题
# ------------------------------------------------------------
parts.append(heading('四、高考真题——电势、电势能与静电力做功', page_break=True))
parts.append(label('【真题来源】'))
parts.append(plain('2020年普通高等学校招生全国统一考试江苏卷物理第9题。为保证脱离原图后仍可独立作答，以下将图示位置关系用文字完整表述。'))
parts.append(label('【题目】'))
parts.append(plain('一根绝缘轻杆的两端固定带有等量异号电荷的小球，不计重力，轻杆可绕中点O转动。匀强电场方向水平向右。初态时，带正电小球位于O点右下方的A点，带负电小球位于O点左上方的B点；在静电力作用下，轻杆逆时针转到水平位置，末态为负电小球在O点左侧、正电小球在O点右侧。取O点电势为0。'))
parts.append(paragraph(mono('初态：B（−q）●'), align='center', after=0, line=280))
parts.append(paragraph(mono('               ╲'), align='center', after=0, line=280))
parts.append(paragraph(mono('                O────────→ E'), align='center', after=0, line=280))
parts.append(paragraph(mono('                 ╲'), align='center', after=0, line=280))
parts.append(paragraph(mono('                  ● A（+q）'), align='center', after=20, line=280))
parts.append(paragraph(mono('末态：B（−q）●────O────● A（+q）'), align='center', after=45, line=280))
parts.append(plain('下列说法正确的有（　　）'))
parts.append(plain('A．电场中A点电势低于B点'))
parts.append(plain('B．转动中两小球的电势能始终相等'))
parts.append(plain('C．该过程静电力对两小球均做负功'))
parts.append(plain('D．该过程两小球的总电势能增加'))
parts.append(label('【答案】'))
parts.append(plain('A、B。'))
parts.append(label('【解析】'))
parts.append(plain('沿匀强电场方向电势降低。由于A点位于B点的右侧，因此φA<φB，A项正确。'))
parts.append(plain('设转动过程中正电小球相对O点的水平坐标为x，则负电小球的水平坐标为−x。取O点电势为0，可得'))
parts.append(eq(phi('A') + mop('=') + mop('−') + mv('E') + mv('x') + mop('，') +
                phi('B') + mop('=') + mv('E') + mv('x')))
parts.append(plain('设正电小球带电荷量+q，负电小球带电荷量−q。根据电势能与电势的关系可得'))
parts.append(eq(Ep('A') + mop('=') + q() + phi('A') + mop('=') + mop('−') + q() + mv('E') + mv('x')))
parts.append(eq(Ep('B') + mop('=') + mparen(mop('−') + q()) + phi('B') + mop('=') + mop('−') + q() + mv('E') + mv('x')))
parts.append(plain('所以转动过程中两小球的电势能始终相等，B项正确。'))
parts.append(plain('轻杆转到水平位置的过程中，x逐渐增大，两小球的电势能均减小。根据静电力做功与电势能变化的关系，静电力对两小球均做正功，两小球的总电势能减小，因此C、D项错误。'))
parts.append(label('【与原例题的关联】'))
parts.append(plain('原例题通过W=qU由静电力做功反推电势差和电势高低；本题则由匀强电场方向和空间位置判断电势，再利用Ep=qφ及W=−ΔEp判断电势能和静电力做功。两题的核心均是“电势差—电势能—静电力做功”之间的相互转换。'))

# ============================================================
# OOXML 打包
# ============================================================
sect = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1247" w:right="1191" w:bottom="1247" w:left="1191" '
    'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

document_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:r="{R}">'
    f'<w:body>{"".join(parts)}{sect}</w:body></w:document>'
)

styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>'''

settings_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W}">
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
<w:defaultTabStop w:val="420"/>
</w:settings>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

with ZipFile(OUT, 'w', ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', root_rels)
    z.writestr('word/document.xml', document_xml)
    z.writestr('word/styles.xml', styles_xml)
    z.writestr('word/settings.xml', settings_xml)
    z.writestr('word/_rels/document.xml.rels', doc_rels)


# ============================================================
# 强制自检：数值、结构、公式、角标、内容一致性
# ============================================================
# 原例题
assert abs(1.6e-7 / 2.0e-9 - 80.0) < 1e-12
assert abs(-4.0e-7 / 2.0e-9 + 200.0) < 1e-12
assert abs((80.0 - 200.0) + 120.0) < 1e-12
assert abs(1.5e-9 * (-120.0) + 1.8e-7) < 1e-20
# 基础变式
assert abs(1.2e-7 / 4.0e-9 - 30.0) < 1e-12
assert abs(-2.0e-7 / 4.0e-9 + 50.0) < 1e-12
assert abs(3.0e-9 * (-20.0) + 6.0e-8) < 1e-20
# 情境变式
assert abs(3.2e-18 / (-1.60e-19) + 20.0) < 1e-12
assert abs((-1.6e-18) / (-1.60e-19) - 10.0) < 1e-12
assert abs(3.2e-18 - 1.6e-18 - 1.6e-18) < 1e-30
# 素养提升
assert abs(1.2e-7 / 2.0e-9 - 60.0) < 1e-12
assert abs((-1.6e-7) / (-4.0e-9) - 40.0) < 1e-12
assert abs(3.0e-7 / 3.0e-9 - 100.0) < 1e-12
assert abs(60.0 / 0.30 - 200.0) < 1e-12
assert abs(40.0 / 0.20 - 200.0) < 1e-12
assert abs(5.0e-9 * 100.0 - 5.0e-7) < 1e-20

assert OUT.exists() and OUT.stat().st_size > 7000
with ZipFile(OUT) as z:
    required = {
        '[Content_Types].xml', '_rels/.rels', 'word/document.xml',
        'word/styles.xml', 'word/settings.xml', 'word/_rels/document.xml.rels'
    }
    assert required.issubset(set(z.namelist()))
    xml = z.read('word/document.xml').decode('utf-8')
    ET.fromstring(xml)
    ET.fromstring(z.read('word/styles.xml'))
    ET.fromstring(z.read('word/settings.xml'))

    # 禁止 Unicode 伪上下标，避免角标间距异常
    forbidden = '⁰¹²³⁴⁵⁶⁷⁸⁹⁻₀₁₂₃₄₅₆₇₈₉ₑₚ'
    assert not any(ch in xml for ch in forbidden)

    # 正文真上下标与公式真结构
    assert '<w:vertAlign w:val="superscript"/>' in xml
    assert '<w:vertAlign w:val="subscript"/>' in xml
    assert '<m:sSub>' in xml
    assert '<m:sSup>' in xml
    assert '<m:f>' in xml
    assert '<m:oMathPara>' in xml
    assert '<w:tbl>' in xml

    # 固定结构与来源
    assert xml.count('【题目】') == 4
    assert xml.count('【答案】') == 4
    assert xml.count('【解析】') == 4
    assert '2020年普通高等学校招生全国统一考试江苏卷物理第9题' in xml
    assert 'A、B。' in xml

    # 关键答案一致性
    for text in ['φC>φA>φB', 'UAB=30 V', 'UAB=−20 V', 'E=200 V/m', 'CA∶AB=3∶2']:
        assert text in xml

    # 不得存在占位符或乱码标记
    for bad in ['TODO', '待补充', 'XXXX', '�']:
        assert bad not in xml

print(OUT)
