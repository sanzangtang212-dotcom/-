from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

OUT = Path('downloads/一题三变加高考真题_第十章_电势差_重新生成自检版.docx')
OUT.parent.mkdir(parents=True, exist_ok=True)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ---------- 普通文本 ----------
def run(text, font='宋体', size=24, bold=False, italic=False, sub=False, sup=False):
    t = escape(str(text))
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
    return f'<w:r><w:rPr>{"".join(props)}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'


def cn(text, size=24, bold=False):
    return run(text, '宋体', size, bold=bold)


def hei(text, size=28):
    return run(text, '黑体', size, bold=True)


def roman(text, size=24, sub=False, sup=False):
    return run(text, 'Times New Roman', size, sub=sub, sup=sup)


def phys(text, size=24):
    return run(text, 'Times New Roman', size, italic=True)


def mono(text, size=22):
    return run(text, 'Courier New', size)


def sci_text(coef, exponent, unit_text=None, size=24):
    out = roman(f'{coef}×10', size=size)
    out += roman(str(exponent).replace('-', '−'), size=size, sup=True)
    if unit_text:
        out += roman(' ' + unit_text, size=size)
    return out


def inline_sub(symbol, subscript, size=24):
    return phys(symbol, size=size) + roman(subscript, size=size, sub=True)


def para(inner='', align=None, before=0, after=90, line=360,
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
    ppr.append(f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>')
    return f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>{inner}</w:p>'


def body(*runs, indent=True, after=70):
    return para(''.join(runs), first_line=480 if indent else 0, after=after)


def plain(text, indent=True, after=70):
    return body(cn(text), indent=indent, after=after)


def heading(text, page_break=False):
    return para(hei(text, 28), before=180, after=100, line=340, page_break=page_break, keep_next=True)


def label(text):
    return para(hei(text, 24), before=110, after=45, line=320, keep_next=True)

# ---------- OMML 数学 ----------
def mr(text, upright=False, font='Times New Roman', preserve=False):
    t = escape(str(text))
    style = 'p' if upright else 'i'
    space = ' xml:space="preserve"' if preserve else ''
    return (
        '<m:r>'
        f'<m:rPr><m:sty m:val="{style}"/></m:rPr>'
        f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        f'<m:t{space}>{t}</m:t>'
        '</m:r>'
    )


def mv(text): return mr(text, upright=False)
def mu(text, preserve=False): return mr(text, upright=True, preserve=preserve)
def mn(text): return mr(text, upright=True)
def mop(text): return mr(text, upright=True)
def msub(base, sub): return f'<m:sSub><m:sSubPr/><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>'
def msup(base, sup): return f'<m:sSup><m:sSupPr/><m:e>{base}</m:e><m:sup>{sup}</m:sup></m:sSup>'
def mfrac(num, den): return f'<m:f><m:fPr/><m:num>{num}</m:num><m:den>{den}</m:den></m:f>'
def mparen(expr): return f'<m:d><m:dPr/><m:e>{expr}</m:e></m:d>'


def eq(expr, after=80):
    return para(f'<m:oMathPara><m:oMath>{expr}</m:oMath></m:oMathPara>', align='center', before=35, after=after, line=300)


def U(sub): return msub(mv('U'), mu(sub))
def WW(sub): return msub(mv('W'), mu(sub))
def PHI(sub): return msub(mv('φ'), mu(sub))
def EP(sub): return msub(mv('E'), mu('p' + sub))
def Q(): return mv('q')
def SCI(c, e): return mn(c) + mop('×') + msup(mn('10'), mu(str(e).replace('-', '−')))
def FRAC(a, b): return mfrac(a, b)
def UNIT(t): return mu(' ' + t, preserve=True)

# ---------- 数据表 ----------
def table_cell(inner, width, header=False):
    shade = '<w:shd w:val="clear" w:fill="DCE6F1"/>' if header else ''
    tcpr = f'<w:tcW w:w="{width}" w:type="dxa"/>{shade}<w:vAlign w:val="center"/>'
    return f'<w:tc><w:tcPr>{tcpr}</w:tcPr>{para(inner, align="center", after=0, line=300)}</w:tc>'


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
        trs.append('<w:tr>' + ''.join(table_cell(row[j], widths[j], i == 0) for j in range(len(row))) + '</w:tr>')
    return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/>' + borders + '</w:tblPr><w:tblGrid>' + grid + '</w:tblGrid>' + ''.join(trs) + '</w:tbl>'

parts = []
parts.append(para(hei('第十章  静电场中的能量', 34), align='center', after=55))
parts.append(para(hei('2. 电势差', 30), align='center', after=75))
parts.append(para(hei('课本例题·一题三变＋高考真题', 28), align='center', after=190))

# ---------- 课本例题 ----------
parts.append(heading('课本例题原型'))
parts.append(body(
    cn('在匀强电场中，把电荷量为 '), sci_text('2.0', '-9', 'C'), cn(' 的正点电荷从 '), roman('A'),
    cn(' 点移动到 '), roman('B'), cn(' 点，静电力做功为 '), sci_text('1.6', '-7', 'J'),
    cn('；再把这个电荷从 '), roman('B'), cn(' 点移动到 '), roman('C'),
    cn(' 点，静电力做功为 '), sci_text('−4.0', '-7', 'J'), cn('。')
))
parts.append(plain('（1）比较A、B、C三点电势的高低；（2）求三段电势差；（3）求另一正电荷从A点移到C点时静电力做的功；（4）画出一种可能的匀强电场分布。'))
parts.append(heading('原题答案与解析'))
parts.append(plain('根据静电力做功与电势差的关系可得'))
parts.append(eq(U('AB') + mop('=') + FRAC(WW('AB'), Q()) + mop('=') + FRAC(SCI('1.6', '-7'), SCI('2.0', '-9')) + UNIT('V') + mop('=') + mn('80') + UNIT('V')))
parts.append(eq(U('BC') + mop('=') + FRAC(WW('BC'), Q()) + mop('=') + FRAC(mop('−') + SCI('4.0', '-7'), SCI('2.0', '-9')) + UNIT('V') + mop('=') + mop('−') + mn('200') + UNIT('V')))
parts.append(plain('由两段电势差可知，C点电势最高，B点电势最低。取B点电势为0，可得'))
parts.append(eq(PHI('A') + mop('=') + mn('80') + UNIT('V') + mop('，') + PHI('B') + mop('=') + mn('0') + UNIT('V') + mop('，') + PHI('C') + mop('=') + mn('200') + UNIT('V')))
parts.append(eq(U('AC') + mop('=') + U('AB') + mop('+') + U('BC') + mop('=') + mop('−') + mn('120') + UNIT('V')))
parts.append(plain('把电荷量为1.5×10的−9次方C的正点电荷从A点移动到C点，根据静电力做功公式可得'))
parts.append(eq(WW('AC') + mop('=') + Q() + U('AC') + mop('=') + mparen(SCI('1.5', '-9') + mop('×') + mparen(mop('−') + mn('120'))) + UNIT('J') + mop('=') + mop('−') + SCI('1.8', '-7') + UNIT('J')))
parts.append(plain('在匀强电场中，沿电场方向电势降低。一种可能的位置关系为：'))
parts.append(para(mono('C（200 V）────────A（80 V）─────B（0 V）      E →'), align='center', after=45, line=300))
parts.append(plain('电势差与沿电场方向的距离成正比，因此CA∶AB=120∶80=3∶2。'))

# ---------- 基础变式 ----------
parts.append(heading('一、基础变式——同模型巩固', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('在匀强电场中，把电荷量为 '), sci_text('4.0', '-9', 'C'), cn(' 的正点电荷从 '), roman('A'),
    cn(' 点移动到 '), roman('B'), cn(' 点，静电力做功为 '), sci_text('1.2', '-7', 'J'),
    cn('；再从 '), roman('B'), cn(' 点移动到 '), roman('C'), cn(' 点，静电力做功为 '), sci_text('−2.0', '-7', 'J'), cn('。')
))
parts.append(plain('（1）比较A、B、C三点电势的高低。'))
parts.append(plain('（2）求A、B间，B、C间和A、C间的电势差。'))
parts.append(body(cn('（3）把电荷量为 '), sci_text('3.0', '-9', 'C'), cn(' 的正点电荷从A点移动到C点，求静电力做的功。'))
parts.append(plain('（4）若三点位于同一条电场线上，画出一种可能的位置关系，并求CA与AB的距离之比。'))
parts.append(label('【答案】'))
parts.append(body(cn('（1）'), inline_sub('φ', 'C'), roman('>'), inline_sub('φ', 'A'), roman('>'), inline_sub('φ', 'B'), cn('。'))
parts.append(body(cn('（2）'), inline_sub('U', 'AB'), roman('=30 V，'), inline_sub('U', 'BC'), roman('=−50 V，'), inline_sub('U', 'AC'), roman('=−20 V。'))
parts.append(body(cn('（3）'), sci_text('−6.0', '-8', 'J'), cn('。 （4）电场方向由C指向B，CA∶AB=2∶3。'))
parts.append(label('【解析】'))
parts.append(plain('根据静电力做功与电势差的关系可得'))
parts.append(eq(U('AB') + mop('=') + FRAC(SCI('1.2', '-7'), SCI('4.0', '-9')) + UNIT('V') + mop('=') + mn('30') + UNIT('V')))
parts.append(eq(U('BC') + mop('=') + FRAC(mop('−') + SCI('2.0', '-7'), SCI('4.0', '-9')) + UNIT('V') + mop('=') + mop('−') + mn('50') + UNIT('V')))
parts.append(eq(U('AC') + mop('=') + U('AB') + mop('+') + U('BC') + mop('=') + mop('−') + mn('20') + UNIT('V')))
parts.append(plain('取B点电势为0，可得A点电势为30 V，C点电势为50 V。'))
parts.append(eq(WW('AC') + mop('=') + mparen(SCI('3.0', '-9') + mop('×') + mparen(mop('−') + mn('20'))) + UNIT('J') + mop('=') + mop('−') + SCI('6.0', '-8') + UNIT('J')))
parts.append(para(mono('C（50 V）────A（30 V）──────B（0 V）      E →'), align='center', after=45, line=300))
parts.append(plain('C、A间电势差大小为20 V，A、B间电势差大小为30 V，因此CA∶AB=2∶3。'))

# ---------- 情境变式 ----------
parts.append(heading('二、情境变式——负电荷的符号迁移', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('某电子束标定装置中，一个电子的电荷量为 '), sci_text('−1.60', '-19', 'C'),
    cn('。电子从A点移动到B点时，静电力做功为 '), sci_text('3.2', '-18', 'J'),
    cn('；从B点移动到C点时，静电力做功为 '), sci_text('−1.6', '-18', 'J'), cn('。')
))
parts.append(plain('（1）求三段电势差。'))
parts.append(plain('（2）比较A、B、C三点电势的高低。'))
parts.append(plain('（3）求电子从A点移动到C点时静电力做的功。'))
parts.append(body(cn('（4）若换成电荷量为 '), sci_text('1.60', '-19', 'C'), cn(' 的正电荷，从A点移动到C点，静电力做功为多少？'))
parts.append(label('【答案】'))
parts.append(body(cn('（1）'), inline_sub('U', 'AB'), roman('=−20 V，'), inline_sub('U', 'BC'), roman('=10 V，'), inline_sub('U', 'AC'), roman('=−10 V。'))
parts.append(body(cn('（2）'), inline_sub('φ', 'B'), roman('>'), inline_sub('φ', 'C'), roman('>'), inline_sub('φ', 'A'), cn('。'))
parts.append(body(cn('（3）'), sci_text('1.6', '-18', 'J'), cn('。 （4）'), sci_text('−1.6', '-18', 'J'), cn('。'))
parts.append(label('【解析】'))
parts.append(plain('电子带负电，应用电势差定义式时必须把电荷量的负号代入。'))
parts.append(eq(U('AB') + mop('=') + FRAC(SCI('3.2', '-18'), mop('−') + SCI('1.60', '-19')) + UNIT('V') + mop('=') + mop('−') + mn('20') + UNIT('V')))
parts.append(eq(U('BC') + mop('=') + FRAC(mop('−') + SCI('1.6', '-18'), mop('−') + SCI('1.60', '-19')) + UNIT('V') + mop('=') + mn('10') + UNIT('V')))
parts.append(eq(U('AC') + mop('=') + U('AB') + mop('+') + U('BC') + mop('=') + mop('−') + mn('10') + UNIT('V')))
parts.append(plain('取A点电势为0，可得B点电势为20 V，C点电势为10 V。'))
parts.append(eq(WW('AC') + mop('=') + WW('AB') + mop('+') + WW('BC') + mop('=') + SCI('1.6', '-18') + UNIT('J')))
parts.append(plain('A、C两点间的电势差由电场本身决定。换成正电荷后，电荷量符号改变，静电力做功随之变号。'))

# ---------- 素养提升 ----------
parts.append(heading('三、素养提升——数据表与匀强电场检验', page_break=True))
parts.append(label('【题目】'))
parts.append(plain('某实验小组在同一匀强电场中选取不同试探电荷，测得A、B、C三点间的静电力做功。A、B、C位于同一条电场线上，AB=0.30 m，BC=0.20 m。'))
parts.append(table([
    [hei('移动过程', 22), hei('试探电荷量', 22), hei('静电力做功', 22), hei('电势差', 22)],
    [roman('A→B', 22), sci_text('+2.0', '-9', 'C', 22), sci_text('+1.2', '-7', 'J', 22), cn('计算', 22)],
    [roman('B→C', 22), sci_text('−4.0', '-9', 'C', 22), sci_text('−1.6', '-7', 'J', 22), cn('计算', 22)],
    [roman('A→C', 22), sci_text('+3.0', '-9', 'C', 22), sci_text('+3.0', '-7', 'J', 22), cn('计算', 22)],
], [1900, 2600, 2600, 1600]))
parts.append(plain('（1）计算三段电势差，并验证电势差的可加性。'))
parts.append(plain('（2）取C点电势为0，求A、B两点的电势。'))
parts.append(plain('（3）分别利用AB段和BC段数据计算电场强度，判断数据是否符合匀强电场模型。'))
parts.append(plain('（4）以A点为原点、沿A到C方向建立位置坐标，画出电势随位置变化的图像。'))
parts.append(body(cn('（5）预测电荷量为 '), sci_text('5.0', '-9', 'C'), cn(' 的正点电荷从A点移动到C点时静电力做的功。'))
parts.append(label('【答案】'))
parts.append(body(cn('（1）'), inline_sub('U', 'AB'), roman('=60 V，'), inline_sub('U', 'BC'), roman('=40 V，'), inline_sub('U', 'AC'), roman('=100 V。'))
parts.append(body(cn('（2）'), inline_sub('φ', 'A'), roman('=100 V，'), inline_sub('φ', 'B'), roman('=40 V。'))
parts.append(body(cn('（3）两段均得 '), phys('E'), roman('=200 V/m'), cn('，数据符合匀强电场模型。'))
parts.append(body(cn('（4）图像为从（0，100 V）到（0.50 m，0）的下降直线。 （5）'), sci_text('5.0', '-7', 'J'), cn('。'))
parts.append(label('【解析】'))
parts.append(plain('根据静电力做功与电势差的关系，逐行计算可得'))
parts.append(eq(U('AB') + mop('=') + FRAC(SCI('1.2', '-7'), SCI('2.0', '-9')) + UNIT('V') + mop('=') + mn('60') + UNIT('V')))
parts.append(eq(U('BC') + mop('=') + FRAC(mop('−') + SCI('1.6', '-7'), mop('−') + SCI('4.0', '-9')) + UNIT('V') + mop('=') + mn('40') + UNIT('V')))
parts.append(eq(U('AC') + mop('=') + FRAC(SCI('3.0', '-7'), SCI('3.0', '-9')) + UNIT('V') + mop('=') + mn('100') + UNIT('V')))
parts.append(plain('三段数据满足100 V=60 V+40 V。取C点电势为0，可得B点电势为40 V，A点电势为100 V。'))
parts.append(eq(mv('E') + mop('=') + FRAC(U('AB'), mn('0.30') + UNIT('m')) + mop('=') + mn('200') + UNIT('V/m')))
parts.append(eq(mv('E') + mop('=') + FRAC(U('BC'), mn('0.20') + UNIT('m')) + mop('=') + mn('200') + UNIT('V/m')))
parts.append(plain('两段电场强度相同，因此数据与匀强电场模型一致。'))
parts.append(para(mono('φ/V'), align='center', after=0, line=260))
parts.append(para(mono('100 ● A'), align='center', after=0, line=260))
parts.append(para(mono('       ╲'), align='center', after=0, line=260))
parts.append(para(mono(' 40        ● B'), align='center', after=0, line=260))
parts.append(para(mono('             ╲'), align='center', after=0, line=260))
parts.append(para(mono('  0               ● C──── x/m'), align='center', after=45, line=260))
parts.append(eq(WW('AC') + mop('=') + mparen(SCI('5.0', '-9') + mop('×') + mn('100')) + UNIT('J') + mop('=') + SCI('5.0', '-7') + UNIT('J')))

# ---------- 高考真题 ----------
parts.append(heading('四、高考真题——电势、电势能与静电力做功', page_break=True))
parts.append(label('【真题来源】'))
parts.append(plain('2020年普通高等学校招生全国统一考试江苏卷物理第9题。'))
parts.append(label('【题目】'))
parts.append(plain('绝缘轻杆两端固定带有等量异号电荷的小球，不计重力。开始时，两小球分别静止在A、B位置。现外加一匀强电场，在静电力作用下，小球绕轻杆中点O转到水平位置。取O点电势为0。'))
parts.append(para(mono('B（−q）●'), align='center', after=0, line=280))
parts.append(para(mono('          ╲'), align='center', after=0, line=280))
parts.append(para(mono('           O────────→ E'), align='center', after=0, line=280))
parts.append(para(mono('            ╲'), align='center', after=0, line=280))
parts.append(para(mono('             ● A（+q）'), align='center', after=30, line=280))
parts.append(plain('下列说法正确的有（　　）'))
parts.append(plain('A．电场中A点电势低于B点'))
parts.append(plain('B．转动中两小球的电势能始终相等'))
parts.append(plain('C．该过程静电力对两小球均做负功'))
parts.append(plain('D．该过程两小球的总电势能增加'))
parts.append(label('【答案】'))
parts.append(plain('A、B。'))
parts.append(label('【解析】'))
parts.append(plain('沿电场方向电势降低。A点位于B点右侧，因此A点电势低于B点，A项正确。'))
parts.append(plain('设转动过程中正电小球相对O点的水平坐标为x，则负电小球的水平坐标为−x。取O点电势为0，可得'))
parts.append(eq(PHI('A') + mop('=') + mop('−') + mv('E') + mv('x') + mop('，') + PHI('B') + mop('=') + mv('E') + mv('x')))
parts.append(plain('根据电势能与电势的关系可得'))
parts.append(eq(EP('A') + mop('=') + Q() + PHI('A') + mop('=') + mop('−') + Q() + mv('E') + mv('x')))
parts.append(eq(EP('B') + mop('=') + mparen(mop('−') + Q()) + PHI('B') + mop('=') + mop('−') + Q() + mv('E') + mv('x')))
parts.append(plain('两小球的电势能始终相等，B项正确。转到水平位置的过程中，x增大，两小球的电势能均减小，因此静电力均做正功，总电势能减小，C、D项错误。'))
parts.append(label('【与课本例题的关联】'))
parts.append(plain('课本例题由静电力做功反求电势差和电势高低；本题由空间位置判断电势，再利用电势能变化判断静电力做功。两题共同考查电势差、电势能和静电力做功之间的转换。'))

# ---------- 打包 ----------
sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1247" w:right="1191" w:bottom="1247" w:left="1191" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
document_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + f'<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:r="{R}"><w:body>{"".join(parts)}{sect}</w:body></w:document>'
styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>'''
settings_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:defaultTabStop w:val="420"/></w:settings>'''
content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''

with ZipFile(OUT, 'w', ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', root_rels)
    z.writestr('word/document.xml', document_xml)
    z.writestr('word/styles.xml', styles_xml)
    z.writestr('word/settings.xml', settings_xml)
    z.writestr('word/_rels/document.xml.rels', doc_rels)

# ---------- 自检 ----------
assert abs(1.6e-7 / 2.0e-9 - 80) < 1e-12
assert abs(-4.0e-7 / 2.0e-9 + 200) < 1e-12
assert abs(1.5e-9 * (-120) + 1.8e-7) < 1e-20
assert abs(1.2e-7 / 4.0e-9 - 30) < 1e-12
assert abs(-2.0e-7 / 4.0e-9 + 50) < 1e-12
assert abs(3.2e-18 / (-1.60e-19) + 20) < 1e-12
assert abs((-1.6e-18) / (-1.60e-19) - 10) < 1e-12
assert abs(60 / 0.30 - 200) < 1e-12
assert abs(40 / 0.20 - 200) < 1e-12

assert OUT.exists() and OUT.stat().st_size > 7000
with ZipFile(OUT) as z:
    names = set(z.namelist())
    required = {'[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml', 'word/settings.xml', 'word/_rels/document.xml.rels'}
    assert required.issubset(names)
    xml = z.read('word/document.xml').decode('utf-8')
    ET.fromstring(xml)
    ET.fromstring(z.read('word/styles.xml'))
    ET.fromstring(z.read('word/settings.xml'))
    forbidden = '⁰¹²³⁴⁵⁶⁷⁸⁹⁻₀₁₂₃₄₅₆₇₈₉ₑₚ'
    assert not any(ch in xml for ch in forbidden)
    for token in ['UAB', 'UBC', 'UAC', 'WAB', 'WBC', 'WAC', 'φA', 'φB', 'φC', 'EpA', 'EpB']:
        assert token not in xml
    assert '<w:vertAlign w:val="superscript"/>' in xml
    assert '<w:vertAlign w:val="subscript"/>' in xml
    assert '<m:sSub>' in xml and '<m:sSup>' in xml and '<m:f>' in xml
    assert '<m:oMathPara>' in xml and '<w:tbl>' in xml
    assert xml.count('【题目】') == 4
    assert xml.count('【答案】') == 4
    assert xml.count('【解析】') == 4
    assert '2020年普通高等学校招生全国统一考试江苏卷物理第9题' in xml
    assert 'A、B。' in xml
    for value in ['−1.8', '−6.0', '200 V/m', '3∶2']:
        assert value in xml
    for bad in ['TODO', '待补充', 'XXXX', '�']:
        assert bad not in xml

print(OUT)
