from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape
from math import sqrt
import xml.etree.ElementTree as ET

OUT = Path('downloads/一题三变加高考真题_第十章_电势差与电场强度的关系_出版规范版.docx')
OUT.parent.mkdir(parents=True, exist_ok=True)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


# ============================================================
# 普通 Word 文本
# ============================================================
def wrun(text, font='宋体', size=24, bold=False, italic=False,
         sub=False, sup=False):
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
    return f'<w:r><w:rPr>{"".join(props)}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def cn(text, size=24, bold=False):
    return wrun(text, font='宋体', size=size, bold=bold)


def hei(text, size=28):
    return wrun(text, font='黑体', size=size, bold=True)


def roman(text, size=24, sub=False, sup=False):
    return wrun(text, font='Times New Roman', size=size, sub=sub, sup=sup)


def phys(text, size=24):
    return wrun(text, font='Times New Roman', size=size, italic=True)


def velocity(text='v', size=24):
    return wrun(text, font='Book Antiqua', size=size, italic=True)


def sci_text(coef, exponent, unit_text=None, size=24):
    out = roman(f'{coef}×10', size=size)
    out += roman(str(exponent).replace('-', '−'), size=size, sup=True)
    if unit_text:
        out += roman(' ' + unit_text, size=size)
    return out


def variable_with_sub(symbol, subscript, font='Times New Roman', size=24):
    italic_font = 'Book Antiqua' if symbol == 'v' else font
    return wrun(symbol, font=italic_font, size=size, italic=True) + roman(subscript, size=size, sub=True)


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


def p_text(text, indent=True, after=70):
    return paragraph(cn(text), first_line=480 if indent else 0, after=after)


def p_runs(runs, indent=True, after=70, align=None):
    return paragraph(''.join(runs), align=align,
                     first_line=480 if indent else 0, after=after)


def label(text):
    return paragraph(hei(text, size=24), before=110, after=45,
                     line=320, keep_next=True)


def heading(text, page_break=False):
    return paragraph(hei(text, size=28), before=180, after=100,
                     line=340, page_break=page_break, keep_next=True)


# ============================================================
# OMML 数学公式
# ============================================================
def mr(text, upright=False, font='Times New Roman', preserve=False):
    text = escape(str(text))
    style = 'p' if upright else 'i'
    space = ' xml:space="preserve"' if preserve else ''
    return (
        '<m:r>'
        f'<m:rPr><m:sty m:val="{style}"/></m:rPr>'
        f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        f'<m:t{space}>{text}</m:t>'
        '</m:r>'
    )


def mv(text, font='Times New Roman'):
    return mr(text, upright=False, font=font)


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


def mrad(expr):
    return '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>' + expr + '</m:e></m:rad>'


def mparen(expr):
    return f'<m:d><m:dPr/><m:e>{expr}</m:e></m:d>'


def eq(expr, after=80):
    return paragraph(
        f'<m:oMathPara><m:oMath>{expr}</m:oMath></m:oMathPara>',
        align='center', before=35, after=after, line=300
    )


def ratio(numerator, denominator):
    return mfrac(numerator, denominator)


def sq(expr):
    return msup(expr, mn('2'))


def power10(exponent):
    return msup(mn('10'), mu(str(exponent).replace('-', '−')))


def sci(coef, exponent):
    return mn(coef) + mop('×') + power10(exponent)


def unit(text):
    return mu(' ' + text, preserve=True)


def v(sub=None):
    base = mv('v', font='Book Antiqua')
    return msub(base, mu(sub)) if sub is not None else base


def sym(symbol, sub=None):
    base = mv(symbol)
    return msub(base, mu(sub)) if sub is not None else base


# ============================================================
# 数据表格
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
parts.append(paragraph(hei('3. 电势差与电场强度的关系', size=30), align='center', after=75))
parts.append(paragraph(hei('例题·一题三变＋高考真题', size=28), align='center', after=200))

# ============================================================
# 原例题解析
# ============================================================
parts.append(heading('原例题解析'))
parts.append(p_runs([
    cn('真空中平行金属板 '), roman('M'), cn('、'), roman('N'), cn(' 之间的距离 '),
    phys('d'), roman('='), roman('0.04 m'), cn('，带电粒子的质量 '), phys('m'), roman('='),
    sci_text('2.0', '-15', 'kg'), cn('，电荷量 '), phys('q'), roman('='),
    sci_text('8.0', '-15', 'C'), cn('，两板间电压 '), phys('U'), roman('='), roman('200 V'), cn('。')
]))
parts.append(p_text('（1）两板间为匀强电场。根据电势差与电场强度的关系可得'))
parts.append(eq(sym('E') + mop('=') + ratio(sym('U'), sym('d')) + mop('=') +
                ratio(mn('200'), mn('0.04')) + unit('V/m') + mop('=') +
                sci('5.0', '3') + unit('V/m')))
parts.append(p_text('根据静电力公式可得'))
parts.append(eq(sym('F') + mop('=') + sym('q') + sym('E') + mop('=') +
                mparen(sci('8.0', '-15') + mop('×') + sci('5.0', '3')) + unit('N') +
                mop('=') + sci('4.0', '-11') + unit('N')))
parts.append(p_text('粒子所受重力约为2.0×10的−14次方N，远小于静电力，因此以下忽略重力。'))
parts.append(p_text('（2）粒子从M板旁由静止运动到N板。根据动能定理可得'))
parts.append(eq(sym('q') + sym('U') + mop('=') + ratio(mn('1'), mn('2')) + sym('m') + sq(v())))
parts.append(eq(v() + mop('=') + mrad(ratio(mn('2') + sym('q') + sym('U'), sym('m'))) +
                mop('=') + mrad(ratio(mn('2') + mop('×') + sci('8.0', '-15') + mop('×') + mn('200'),
                                       sci('2.0', '-15'))) + unit('m/s') +
                mop('=') + mn('40') + unit('m/s')))
parts.append(p_text('（3）当极板间距增大为原来的2倍而电压保持不变时，场强和静电力均减小为原来的1/2。'))
parts.append(eq(sym('E', '1') + mop('=') + ratio(sym('U'), mn('2') + sym('d')) +
                mop('=') + ratio(sym('E'), mn('2')) + mop('，') +
                sym('F', '1') + mop('=') + ratio(sym('F'), mn('2')) +
                mop('=') + sci('2.0', '-11') + unit('N')))
parts.append(p_text('由于粒子跨越两板时静电力做功仍为qU，故到达N板时的速度不变。'))
parts.append(eq(v('1') + mop('=') + v() + mop('=') + mn('40') + unit('m/s')))
parts.append(p_text('核心结论：电压一定时，F=qU/d随板间距改变；粒子由静止跨越两板的末速度由qU决定，与板间距无关。'))

# ============================================================
# 基础变式
# ============================================================
parts.append(heading('一、基础变式——同模型巩固', page_break=True))
parts.append(label('【题目】'))
parts.append(p_runs([
    cn('真空中两平行金属板间距为 '), roman('0.030 m'), cn('，两板间加 '), roman('225 V'),
    cn(' 直流电压。一个质量为 '), sci_text('2.0', '-15', 'kg'), cn('、电荷量为 '),
    sci_text('4.0', '-15', 'C'), cn(' 的正粒子位于正极板旁，由静止释放，不计重力。')
]))
parts.append(p_text('（1）求两板间的电场强度。'))
parts.append(p_text('（2）求粒子所受静电力的大小。'))
parts.append(p_text('（3）求粒子到达负极板时的速度。'))
parts.append(p_text('（4）若仅将板间距增大为原来的2倍，求静电力和末速度。'))
parts.append(label('【答案】'))
parts.append(p_runs([
    cn('（1）'), sci_text('7.5', '3', 'V/m'), cn('；（2）'), sci_text('3.0', '-11', 'N'),
    cn('；（3）'), roman('30 m/s'), cn('；（4）'), sci_text('1.5', '-11', 'N'),
    cn('，末速度仍为 '), roman('30 m/s'), cn('。')
]))
parts.append(label('【解析】'))
parts.append(p_text('根据匀强电场中电势差与电场强度的关系可得'))
parts.append(eq(sym('E') + mop('=') + ratio(mn('225'), mn('0.030')) + unit('V/m') +
                mop('=') + sci('7.5', '3') + unit('V/m')))
parts.append(p_text('根据静电力公式可得'))
parts.append(eq(sym('F') + mop('=') + mparen(sci('4.0', '-15') + mop('×') + sci('7.5', '3')) + unit('N') +
                mop('=') + sci('3.0', '-11') + unit('N')))
parts.append(p_text('根据动能定理可得'))
parts.append(eq(v() + mop('=') + mrad(ratio(mn('2') + mop('×') + sci('4.0', '-15') + mop('×') + mn('225'),
                                       sci('2.0', '-15'))) + unit('m/s') +
                mop('=') + mn('30') + unit('m/s')))
parts.append(p_text('板间距增大为原来的2倍时，电压不变，故场强和静电力均变为原来的1/2，而qU不变，末速度不变。'))

# ============================================================
# 情境变式
# ============================================================
parts.append(heading('二、情境变式——微粒直线加速器的两种调节方式', page_break=True))
parts.append(label('【题目】'))
parts.append(p_runs([
    cn('某微粒直线加速器由两块平行金属板构成。初始板间距为 '), roman('0.040 m'),
    cn('，板间电压为 '), roman('120 V'), cn('。一个质量为 '), sci_text('3.0', '-15', 'kg'),
    cn('、电荷量为 '), sci_text('5.0', '-15', 'C'), cn(' 的正粒子从正极板旁由静止释放，不计重力。')
]))
parts.append(p_text('（1）求初始状态下粒子所受静电力和到达负极板时的速度。'))
parts.append(p_text('（2）将板间距增大为原来的2倍，若电压保持不变，求静电力和末速度。'))
parts.append(p_text('（3）将板间距增大为原来的2倍，若通过调节电源使电场强度保持不变，求此时板间电压、静电力和末速度。'))
parts.append(label('【答案】'))
parts.append(p_runs([
    cn('（1）'), sci_text('1.5', '-11', 'N'), cn('，'), roman('20 m/s'),
    cn('；（2）'), sci_text('7.5', '-12', 'N'), cn('，'), roman('20 m/s'),
    cn('；（3）'), roman('240 V'), cn('，'), sci_text('1.5', '-11', 'N'),
    cn('，'), roman('20√2 m/s'), cn('，约为 '), roman('28.3 m/s'), cn('。')
]))
parts.append(label('【解析】'))
parts.append(p_text('初始状态下，根据E=U/d和F=qE可得'))
parts.append(eq(sym('E') + mop('=') + ratio(mn('120'), mn('0.040')) + unit('V/m') +
                mop('=') + sci('3.0', '3') + unit('V/m')))
parts.append(eq(sym('F') + mop('=') + mparen(sci('5.0', '-15') + mop('×') + sci('3.0', '3')) + unit('N') +
                mop('=') + sci('1.5', '-11') + unit('N')))
parts.append(eq(v() + mop('=') + mrad(ratio(mn('2') + mop('×') + sci('5.0', '-15') + mop('×') + mn('120'),
                                       sci('3.0', '-15'))) + unit('m/s') +
                mop('=') + mn('20') + unit('m/s')))
parts.append(p_text('板间距加倍而电压不变时，场强和静电力减半；粒子跨越两板的电势差不变，故末速度不变。'))
parts.append(p_text('板间距加倍而场强保持不变时，根据U=Ed可知电压加倍。静电力保持不变，静电力做功加倍。'))
parts.append(eq(sym('U', '2') + mop('=') + sym('E') + mparen(mn('2') + sym('d')) +
                mop('=') + mn('2') + sym('U') + mop('=') + mn('240') + unit('V')))
parts.append(eq(v('2') + mop('=') + mrad(ratio(mn('2') + sym('q') + mparen(mn('2') + sym('U')), sym('m'))) +
                mop('=') + mrad(mn('2')) + v() + mop('=') + mn('20') + mrad(mn('2')) + unit('m/s')))
parts.append(p_text('本题关键是先判断保持不变的是电压还是电场强度，再确定静电力和静电力做功。'))

# ============================================================
# 素养提升
# ============================================================
parts.append(heading('三、素养提升——控制变量与多方案比较', page_break=True))
parts.append(label('【题目】'))
parts.append(p_text('同一带正电粒子均从正极板旁由静止释放，不计重力。以方案A作为参照，比较其他方案中板间电压、板间距、电场强度、静电力和末速度相对于方案A的倍数。'))
parts.append(table([
    [hei('方案', size=22), hei('电压倍数', size=22), hei('距离倍数', size=22), hei('场强倍数', size=22), hei('静电力倍数', size=22), hei('速度倍数', size=22)],
    [roman('A', size=22), roman('1', size=22), roman('1', size=22), cn('待填', size=22), cn('待填', size=22), cn('待填', size=22)],
    [roman('B', size=22), roman('1', size=22), roman('2', size=22), cn('待填', size=22), cn('待填', size=22), cn('待填', size=22)],
    [roman('C', size=22), roman('2', size=22), roman('2', size=22), cn('待填', size=22), cn('待填', size=22), cn('待填', size=22)],
    [roman('D', size=22), roman('1/2', size=22), roman('1/2', size=22), cn('待填', size=22), cn('待填', size=22), cn('待填', size=22)],
], [1000, 1350, 1350, 1350, 1350, 1600]))
parts.append(p_text('（1）推导场强倍数、静电力倍数和速度倍数与电压倍数、距离倍数的关系。'))
parts.append(p_text('（2）完成表格。'))
parts.append(p_text('（3）哪些方案中粒子所受静电力相同？哪些方案中末速度相同？'))
parts.append(p_text('（4）概括板间电压和板间距分别影响静电力与末速度的规律。'))
parts.append(label('【答案】'))
parts.append(p_text('（1）场强倍数等于电压倍数与距离倍数之比，静电力倍数等于场强倍数，速度倍数等于电压倍数的平方根。'))
parts.append(p_text('（2）A：1、1、1；B：1/2、1/2、1；C：1、1、√2；D：1、1、1/√2。'))
parts.append(p_text('（3）A、C、D三方案静电力相同；A、B两方案末速度相同。'))
parts.append(p_text('（4）静电力由U/d决定；由静止跨越两板的末速度由电压U决定，与板间距d无关。'))
parts.append(label('【解析】'))
parts.append(p_text('根据匀强电场中电势差与电场强度的关系可得'))
parts.append(eq(ratio(sym('E'), sym('E', '0')) + mop('=') +
                ratio(ratio(sym('U'), sym('U', '0')), ratio(sym('d'), sym('d', '0')))))
parts.append(p_text('同一粒子的电荷量不变，根据F=qE可得'))
parts.append(eq(ratio(sym('F'), sym('F', '0')) + mop('=') + ratio(sym('E'), sym('E', '0'))))
parts.append(p_text('根据动能定理可得'))
parts.append(eq(ratio(v(), v('0')) + mop('=') + mrad(ratio(sym('U'), sym('U', '0')))))
parts.append(p_text('代入各方案的U/U0和d/d0，即可得到表中各比值。该结果表明，改变板间距会改变加速度和运动时间，但在电压一定时不会改变末动能。'))

# ============================================================
# 高考真题
# ============================================================
parts.append(heading('四、高考真题——直线加速与类平抛运动综合', page_break=True))
parts.append(label('【真题来源】'))
parts.append(p_text('2022年浙江省6月普通高校招生选考科目考试物理第9题。原题图示位置关系已用文字完整表述。'))
parts.append(label('【题目】'))
parts.append(p_text('带等量异种电荷的两正对平行金属板M、N间存在匀强电场，板长为L，不考虑边界效应。M板在左，N板在右，粒子源位于M板中点。'))
parts.append(p_runs([
    cn('在 '), phys('t'), roman('=0'), cn(' 时，粒子源同时发射两个速度大小均为 '), variable_with_sub('v', '0'),
    cn(' 的相同粒子。第一个粒子的初速度垂直M板向右，到达N板时速度大小为 '),
    roman('√2'), variable_with_sub('v', '0'), cn('；第二个粒子的初速度平行M板竖直向下，刚好从N板下端射出。不计重力和粒子间的相互作用。')
]))
parts.append(p_text('下列说法正确的是（　　）'))
parts.append(p_text('A．M板电势高于N板电势'))
parts.append(p_text('B．两个粒子的电势能都增加'))
parts.append(p_runs([cn('C．粒子在两板间的加速度为 '), phys('a'), roman('='), roman('2'), variable_with_sub('v', '0'), roman('2', sup=True), roman('/'), phys('L')]))
parts.append(p_runs([cn('D．粒子从N板下端射出的时间为 '), phys('t'), roman('='), roman('(√2−1)'), phys('L'), roman('/(2'), variable_with_sub('v', '0'), roman(')')]))
parts.append(label('【答案】'))
parts.append(p_text('C。'))
parts.append(label('【解析】'))
parts.append(p_text('第一个粒子由M板运动到N板时速度增大，说明静电力做正功，粒子的电势能减小。两个粒子从M板运动到N板，电势能变化相同，因此B项错误。由于粒子的电性未知，不能确定电场方向，也不能确定M、N两板电势的高低，A项错误。'))
parts.append(p_text('设两板间距为d，粒子在板间的加速度大小为a。对于垂直M板向右的粒子，根据匀变速直线运动规律可得'))
parts.append(eq(sq(mn('√2') + v('0')) + mop('−') + sq(v('0')) + mop('=') + mn('2') + sym('a') + sym('d')))
parts.append(eq(sym('d') + mop('=') + ratio(sq(v('0')), mn('2') + sym('a'))))
parts.append(p_text('对于平行M板向下的粒子，竖直方向做匀速直线运动，从M板中点运动到N板下端的竖直位移为L/2，因此'))
parts.append(eq(ratio(sym('L'), mn('2')) + mop('=') + v('0') + sym('t')))
parts.append(eq(sym('t') + mop('=') + ratio(sym('L'), mn('2') + v('0'))))
parts.append(p_text('该粒子水平方向做初速度为零的匀加速直线运动，水平方向位移等于板间距，故'))
parts.append(eq(sym('d') + mop('=') + ratio(mn('1'), mn('2')) + sym('a') + sq(sym('t')) +
                mop('=') + ratio(sym('a') + sq(sym('L')), mn('8') + sq(v('0')))))
parts.append(p_text('联立两式可得'))
parts.append(eq(sym('a') + mop('=') + ratio(mn('2') + sq(v('0')), sym('L'))))
parts.append(p_text('因此C项正确。第二个粒子从N板下端射出的实际时间为L/(2v0)，D项所给时间是第一个粒子在水平方向加速所需的时间，故D项错误。'))
parts.append(label('【与原例题的关联】'))
parts.append(p_text('原例题研究粒子在平行板匀强电场中的直线加速，核心关系为E=U/d、F=qE以及静电力做功改变动能。本题仍使用同一匀强电场和同一加速度，但把直线加速与类平抛运动组合起来，通过两个粒子的运动信息反求加速度，实现了由单一模型到综合模型的迁移。'))

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
# 强制自检
# ============================================================
# 原例题
assert abs(200 / 0.04 - 5.0e3) < 1e-10
assert abs(8.0e-15 * 5.0e3 - 4.0e-11) < 1e-24
assert abs(sqrt(2 * 8.0e-15 * 200 / 2.0e-15) - 40.0) < 1e-12
assert abs(2.0e-14 / 4.0e-11 - 5.0e-4) < 1e-12
# 基础变式
assert abs(225 / 0.030 - 7.5e3) < 1e-10
assert abs(4.0e-15 * 7.5e3 - 3.0e-11) < 1e-24
assert abs(sqrt(2 * 4.0e-15 * 225 / 2.0e-15) - 30.0) < 1e-12
# 情境变式
assert abs(120 / 0.040 - 3.0e3) < 1e-10
assert abs(5.0e-15 * 3.0e3 - 1.5e-11) < 1e-24
assert abs(sqrt(2 * 5.0e-15 * 120 / 3.0e-15) - 20.0) < 1e-12
assert abs(sqrt(2 * 5.0e-15 * 240 / 3.0e-15) - 20 * sqrt(2)) < 1e-12
# 高考真题
# 令v0=L=1，可得a=2，t=1/2，d=1/4；两种求d方式一致
v0_num = 1.0
L_num = 1.0
a_num = 2.0 * v0_num ** 2 / L_num
t_num = L_num / (2.0 * v0_num)
d1 = v0_num ** 2 / (2.0 * a_num)
d2 = 0.5 * a_num * t_num ** 2
assert abs(d1 - d2) < 1e-12
assert abs(t_num - 0.5) < 1e-12

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

    # 禁止 Unicode 伪上下标
    forbidden = '⁰¹²³⁴⁵⁶⁷⁸⁹⁻₀₁₂₃₄₅₆₇₈₉ₑₚ'
    assert not any(ch in xml for ch in forbidden)

    # 真上下标、分式、根式和可编辑公式
    assert '<w:vertAlign w:val="superscript"/>' in xml
    assert '<w:vertAlign w:val="subscript"/>' in xml
    assert '<m:sSub>' in xml
    assert '<m:sSup>' in xml
    assert '<m:f>' in xml
    assert '<m:rad>' in xml
    assert '<m:oMathPara>' in xml
    assert '<w:tbl>' in xml
    assert 'Book Antiqua' in xml
    assert 'Times New Roman' in xml

    # 固定结构和高考来源
    assert xml.count('【题目】') == 4
    assert xml.count('【答案】') == 4
    assert xml.count('【解析】') == 4
    assert '2022年浙江省6月普通高校招生选考科目考试物理第9题' in xml
    assert 'C。' in xml

    # 关键结果一致性
    for text in ['末速度不变', '30 m/s', '28.3 m/s', 'A、C、D三方案静电力相同', '实际时间为']:
        assert text in xml

    # 无占位符、乱码
    for bad in ['TODO', '待补充', 'XXXX', '�']:
        assert bad not in xml

print(OUT)
