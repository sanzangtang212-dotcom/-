from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape

OUT = Path('downloads/一题三变加高考真题_第九章_库仑定律_出版规范版.docx')
OUT.parent.mkdir(parents=True, exist_ok=True)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


# ------------------------------
# Word text runs
# ------------------------------
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


def cn(text, bold=False, size=24):
    return wrun(text, font='宋体', size=size, bold=bold)


def hei(text, bold=True, size=28):
    return wrun(text, font='黑体', size=size, bold=bold)


def sym(text, sub=False, sup=False, font='Times New Roman', size=24):
    return wrun(text, font=font, size=size, italic=not (sub or sup), sub=sub, sup=sup)


def roman(text, sub=False, sup=False, size=24):
    return wrun(text, font='Times New Roman', size=size, italic=False, sub=sub, sup=sup)


def velocity(text='v', sub=None, size=24):
    runs = [wrun(text, font='Book Antiqua', size=size, italic=True)]
    if sub is not None:
        runs.append(roman(sub, sub=True, size=size))
    return ''.join(runs)


def Ftxt(subscript, size=24):
    return sym('F', size=size) + roman(subscript, sub=True, size=size)


def mtxt(subscript, size=24):
    return sym('m', size=size) + roman(subscript, sub=True, size=size)


def Rtxt(subscript, size=24):
    return sym('R', size=size) + roman(subscript, sub=True, size=size)


def sci_txt(coef, exponent, unit_text=None, size=24):
    runs = [roman(coef + '×10', size=size), roman(str(exponent).replace('-', '−'), sup=True, size=size)]
    if unit_text:
        runs.append(roman(' ' + unit_text, size=size))
    return ''.join(runs)


def unit_nm2_per_c2(size=24):
    return ''.join([
        roman(' N·m', size=size), roman('2', sup=True, size=size),
        roman('/C', size=size), roman('2', sup=True, size=size)
    ])


def unit_nm2_per_kg2(size=24):
    return ''.join([
        roman(' N·m', size=size), roman('2', sup=True, size=size),
        roman('/kg', size=size), roman('2', sup=True, size=size)
    ])


def para(inner='', align=None, before=0, after=100, line=360,
         page_break=False, keep_next=False, indent=0):
    ppr = []
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    if page_break:
        ppr.append('<w:pageBreakBefore/>')
    if keep_next:
        ppr.append('<w:keepNext/>')
    if indent:
        ppr.append(f'<w:ind w:firstLine="{indent}"/>')
    ppr.append(
        f'<w:spacing w:before="{before}" w:after="{after}" '
        f'w:line="{line}" w:lineRule="auto"/>'
    )
    return f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>{inner}</w:p>'


def body(*runs, after=70, first_line=True):
    return para(''.join(runs), after=after, line=360, indent=480 if first_line else 0)


def body_plain(text, after=70, first_line=True):
    return body(cn(text), after=after, first_line=first_line)


def label(text):
    return para(hei(text, size=24), before=120, after=50, line=320, keep_next=True)


def heading(text, page_break=False):
    return para(hei(text, size=28), before=180, after=100, line=340,
                page_break=page_break, keep_next=True)


# ------------------------------
# OMML equations
# ------------------------------
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


def mfrac(num, den):
    return f'<m:f><m:fPr/><m:num>{num}</m:num><m:den>{den}</m:den></m:f>'


def mparen(expr):
    return f'<m:d><m:dPr/><m:e>{expr}</m:e></m:d>'


def eq(expr, after=80):
    return para(f'<m:oMathPara><m:oMath>{expr}</m:oMath></m:oMathPara>',
                align='center', before=40, after=after, line=300)


def F(sub):
    return msub(mv('F'), mu(sub))


def mass(sub):
    return msub(mv('m'), mu(sub))


def Rsym(sub):
    return msub(mv('R'), mu(sub))


def v0():
    return msub(mv('v', font='Book Antiqua'), mn('0'))


def va():
    return msub(mv('v', font='Book Antiqua'), mu('a'))


def vb():
    return msub(mv('v', font='Book Antiqua'), mu('b'))


def r0():
    return msub(mv('r'), mn('0'))


def sq(expr):
    return msup(expr, mn('2'))


def cube(expr):
    return msup(expr, mn('3'))


def power10(exp):
    return msup(mn('10'), mu(str(exp).replace('-', '−')))


def sci(coef, exp):
    return mn(coef) + mop('×') + power10(exp)


def ratio(num, den):
    return mfrac(num, den)


def munit(text):
    return mu(' ' + text, preserve=True)


# ------------------------------
# Word tables
# ------------------------------
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
    tr_xml = []
    for i, row in enumerate(rows):
        cells = [table_cell(cell, widths[j], header=(i == 0)) for j, cell in enumerate(row)]
        tr_xml.append(f'<w:tr>{"".join(cells)}</w:tr>')
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        f'<w:jc w:val="center"/>{borders}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>'
        f'{"".join(tr_xml)}</w:tbl>'
    )


# ------------------------------
# Document body
# ------------------------------
parts = []
parts.append(para(hei('第九章  静电场及其应用', size=34), align='center', after=60))
parts.append(para(hei('2. 库仑定律', size=30), align='center', after=80))
parts.append(para(hei('一题三变＋高考真题｜微观粒子间静电力与万有引力的比较', size=28), align='center', after=220))
parts.append(body_plain('核心模型：微观带电粒子之间同时存在静电力和万有引力。'))
parts.append(body_plain('核心方法：分别应用库仑定律和万有引力定律，再通过比值法消去共同因素。'))
parts.append(body_plain('训练路径：直接计算 → 比例迁移 → 一般化推导 → 高考真题迁移。'))
parts.append(body_plain('排版说明：正文中的科学计数法采用真正的 Word 上标；物理量采用规范斜体；说明性下标采用正体；独立公式采用可编辑 OMML 结构。'))

# 一、基础变式
parts.append(heading('一、基础变式——同模型巩固'))
parts.append(label('【题目】'))
parts.append(body(
    cn('在氢原子中，氢原子核可视为质子。质子与电子之间的距离为 '),
    sci_txt('5.3', '-11', 'm'), cn('。已知元电荷 '), sym('e'), roman('='),
    sci_txt('1.60', '-19', 'C'), cn('，质子质量 '), mtxt('p'), roman('='),
    sci_txt('1.67', '-27', 'kg'), cn('，电子质量 '), mtxt('e'), roman('='),
    sci_txt('9.11', '-31', 'kg'), cn('，静电力常量 '), sym('k'), roman('='),
    sci_txt('9.0', '9'), unit_nm2_per_c2(), cn('，万有引力常量 '), sym('G'), roman('='),
    sci_txt('6.67', '-11'), unit_nm2_per_kg2(), cn('。')
))
parts.append(body_plain('（1）求质子与电子之间的静电力大小。'))
parts.append(body_plain('（2）求质子与电子之间的万有引力大小。'))
parts.append(body_plain('（3）求静电力与万有引力的比值，并说明研究氢原子内部电子运动时是否需要考虑万有引力。'))
parts.append(label('【答案】'))
parts.append(body(
    cn('（1）'), sci_txt('8.2', '-8', 'N'), cn('；（2）'), sci_txt('3.6', '-47', 'N'),
    cn('；（3）静电力约为万有引力的 '), sci_txt('2.3', '39'), cn(' 倍，万有引力可以忽略。')
))
parts.append(label('【解析】'))
parts.append(body_plain('（1）求质子与电子之间的静电力。'))
parts.append(body_plain('根据库仑定律可得'))
parts.append(eq(F('e') + mop('=') + mv('k') + ratio(sq(mv('e')), sq(mv('r')))))
parts.append(body_plain('代入数据可得'))
parts.append(eq(
    F('e') + mop('=') + mparen(
        sci('9.0', '9') + mop('×') +
        ratio(sq(mparen(sci('1.60', '-19'))), sq(mparen(sci('5.3', '-11'))))
    ) + munit('N') + mop('=') + sci('8.2', '-8') + munit('N')
))
parts.append(body_plain('质子与电子带异种电荷，因此静电力表现为吸引力。'))
parts.append(body_plain('（2）求质子与电子之间的万有引力。'))
parts.append(body_plain('根据万有引力定律可得'))
parts.append(eq(F('G') + mop('=') + mv('G') + ratio(mass('p') + mass('e'), sq(mv('r')))))
parts.append(body_plain('代入数据可得'))
parts.append(eq(
    F('G') + mop('=') + mparen(
        sci('6.67', '-11') + mop('×') +
        ratio(sci('1.67', '-27') + mop('×') + sci('9.11', '-31'),
              sq(mparen(sci('5.3', '-11'))))
    ) + munit('N') + mop('=') + sci('3.6', '-47') + munit('N')
))
parts.append(body_plain('（3）比较两种力。'))
parts.append(body_plain('根据两种力的比值可得'))
parts.append(eq(
    ratio(F('e'), F('G')) + mop('=') +
    ratio(mv('k') + ratio(sq(mv('e')), sq(mv('r'))),
          mv('G') + ratio(mass('p') + mass('e'), sq(mv('r')))) +
    mop('=') + ratio(mv('k') + sq(mv('e')), mv('G') + mass('p') + mass('e'))
))
parts.append(eq(
    ratio(F('e'), F('G')) + mop('=') +
    ratio(sci('8.2', '-8'), sci('3.6', '-47')) + mop('≈') + sci('2.3', '39')
))
parts.append(body_plain('两种力的比值没有单位。由计算结果可知，静电力远大于万有引力，因此研究氢原子内部电子运动时可以忽略万有引力。'))

# 二、情境变式
parts.append(heading('二、情境变式——新情境迁移', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('氘是氢的同位素。氘核带电荷 '), roman('+'), sym('e'), cn('，质量近似为 '),
    roman('2'), mtxt('p'), cn('。设氢原子中质子与电子的距离为 '), sym('r'), roman('0', sub=True),
    cn('，氘原子中氘核与电子的距离为 '), roman('2'), sym('r'), roman('0', sub=True), cn('。')
))
parts.append(body(
    cn('分别用 '), Ftxt('eH'), cn('、'), Ftxt('GH'), cn(' 表示氢原子中的静电力和万有引力，用 '),
    Ftxt('eD'), cn('、'), Ftxt('GD'), cn(' 表示氘原子中的静电力和万有引力。')
))
parts.append(body_plain('（1）求氘原子中的静电力与氢原子中的静电力之比。'))
parts.append(body_plain('（2）求氘原子中的万有引力与氢原子中的万有引力之比。'))
parts.append(body(
    cn('（3）设氘原子和氢原子中静电力与万有引力的比值分别为 '), Rtxt('D'), cn('、'), Rtxt('H'),
    cn('，求二者的关系，并估算 '), Rtxt('D'), cn('。')
))
parts.append(label('【答案】'))
parts.append(body_plain('（1）1/4；（2）1/2；（3）氘原子中的比值为氢原子中比值的 1/2，约为 1.1×10 的 39 次方。'))
parts.append(label('【解析】'))
parts.append(body_plain('（1）比较两种原子中的静电力。'))
parts.append(body_plain('根据库仑定律可得'))
parts.append(eq(
    ratio(F('eD'), F('eH')) + mop('=') +
    ratio(mv('k') + ratio(sq(mv('e')), sq(mn('2') + r0())),
          mv('k') + ratio(sq(mv('e')), sq(r0()))) + mop('=') + ratio(mn('1'), mn('4'))
))
parts.append(body_plain('（2）比较两种原子中的万有引力。'))
parts.append(body_plain('根据万有引力定律可得'))
parts.append(eq(
    ratio(F('GD'), F('GH')) + mop('=') +
    ratio(mv('G') + ratio(mparen(mn('2') + mass('p')) + mass('e'), sq(mn('2') + r0())),
          mv('G') + ratio(mass('p') + mass('e'), sq(r0()))) + mop('=') + ratio(mn('1'), mn('2'))
))
parts.append(body_plain('（3）比较两种力的比值。'))
parts.append(body_plain('根据比值的定义可得'))
parts.append(eq(
    ratio(Rsym('D'), Rsym('H')) + mop('=') +
    ratio(ratio(F('eD'), F('eH')), ratio(F('GD'), F('GH'))) +
    mop('=') + ratio(ratio(mn('1'), mn('4')), ratio(mn('1'), mn('2'))) +
    mop('=') + ratio(mn('1'), mn('2'))
))
parts.append(eq(
    Rsym('D') + mop('=') + ratio(mn('1'), mn('2')) + Rsym('H') +
    mop('≈') + sci('1.1', '39')
))
parts.append(body_plain('虽然氘核质量增大使万有引力的相对作用有所增强，但静电力仍远大于万有引力。'))

# 三、素养提升
parts.append(heading('三、素养提升——多表征综合', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('某类“类氢离子”由一个带电荷 '), roman('+'), sym('Z'), sym('e'), cn('、质量近似为 '),
    sym('A'), mtxt('p'), cn(' 的原子核和一个电子组成。设原子核与电子之间的距离为 '),
    sym('λ'), sym('r'), roman('0', sub=True), cn('。定义静电力与万有引力的比值为 '),
    sym('R'), roman('='), Ftxt('e'), roman('/'), Ftxt('G'), cn('，氢原子中的对应比值为 '), Rtxt('H'), cn('。')
))
parts.append(body_plain('下表给出了四种粒子系统的参数。'))
parts.append(table([
    [hei('粒子系统', size=22), hei('Z', size=22), hei('A', size=22), hei('λ', size=22), hei('R/R_H', size=22)],
    [cn('氢原子 H', size=22), roman('1', size=22), roman('1', size=22), roman('1', size=22), cn('', size=22)],
    [cn('氘原子 D', size=22), roman('1', size=22), roman('2', size=22), roman('2', size=22), cn('', size=22)],
    [cn('氦离子 He+', size=22), roman('2', size=22), roman('4', size=22), roman('1/2', size=22), cn('', size=22)],
    [cn('锂离子 Li2+', size=22), roman('3', size=22), roman('7', size=22), roman('3', size=22), cn('', size=22)],
], [2300, 900, 900, 1100, 1800]))
parts.append(body_plain('（1）推导一般情况下两力比值与氢原子中对应比值、核电荷数、质量数和距离因子的关系。'))
parts.append(body_plain('（2）在表格最后一列填出各系统的相对比值。'))
parts.append(body_plain('（3）上述系统中，哪一种系统里万有引力的相对作用最明显？说明判断依据。'))
parts.append(body_plain('（4）已知氢原子中两力比值约为 2.3×10 的 39 次方，估算锂离子中的两力比值，并判断研究电子运动时能否忽略万有引力。'))
parts.append(label('【答案】'))
parts.append(body_plain('（1）一般关系为 R=(Z/A)R_H，与距离因子无关；（2）依次为 1、1/2、1/2、3/7；（3）锂离子，因为其比值最小；（4）约为 9.9×10 的 38 次方，仍可忽略万有引力。'))
parts.append(label('【解析】'))
parts.append(body_plain('（1）推导一般关系。'))
parts.append(body_plain('根据库仑定律可得'))
parts.append(eq(F('e') + mop('=') + mv('k') + ratio(mv('Z') + sq(mv('e')), sq(mv('λ') + r0()))))
parts.append(body_plain('根据万有引力定律可得'))
parts.append(eq(F('G') + mop('=') + mv('G') + ratio(mv('A') + mass('p') + mass('e'), sq(mv('λ') + r0()))))
parts.append(body_plain('两式相除可得'))
parts.append(eq(
    mv('R') + mop('=') + ratio(F('e'), F('G')) + mop('=') +
    ratio(mv('k') + mv('Z') + sq(mv('e')), mv('G') + mv('A') + mass('p') + mass('e')) +
    mop('=') + ratio(mv('Z'), mv('A')) + Rsym('H')
))
parts.append(body_plain('因此，两力比值与距离因子无关，只由核电荷数与质量数的比值决定。'))
parts.append(body_plain('（2）根据 R/R_H=Z/A，可得四种系统的相对比值依次为 1、1/2、1/2、3/7。'))
parts.append(body_plain('（3）比值越小，说明万有引力相对于静电力越明显。四种系统中锂离子的 3/7 最小，因此其万有引力的相对作用最明显。'))
parts.append(body_plain('（4）估算锂离子中的两力比值。'))
parts.append(body_plain('根据一般关系可得'))
parts.append(eq(
    mv('R') + mop('=') + mparen(ratio(mn('3'), mn('7')) + mop('×') + sci('2.3', '39')) +
    mop('=') + sci('9.9', '38')
))
parts.append(body_plain('该比值仍接近 10 的 39 次方，说明静电力仍远大于万有引力，因此研究电子运动时仍可以忽略万有引力。'))

# 四、高考真题
parts.append(heading('四、高考真题——同原理与同方法拓展', page_break=True))
parts.append(label('【真题来源】'))
parts.append(body_plain('2023年高考新课标卷物理第12题。题干根据原题文字整理，原示意图所包含的装置信息改用文字完整表述。'))
parts.append(label('【题目】'))
parts.append(body_plain('密立根油滴实验装置由两块水平放置、间距固定的金属板组成，可从上板中央的小孔向两板间喷入大小不同、带电荷量不同但密度相同的小油滴。'))
parts.append(body_plain('两板间未加电压时，油滴 a、b 在重力和空气阻力作用下竖直向下做匀速直线运动，其速率满足'))
parts.append(eq(va() + mop('=') + v0() + mop('，') + vb() + mop('=') + ratio(v0(), mn('4'))))
parts.append(body_plain('两板间加上电压后，上板接正极，两个油滴很快达到相同的速率，仍均竖直向下做匀速直线运动，其共同速率为'))
parts.append(eq(v0() + ratio(mn('1'), mn('2'))))
parts.append(body_plain('油滴可视为球形，所受空气阻力大小与油滴半径和运动速率的乘积成正比，比例系数为常量。不计空气浮力和油滴间的相互作用。'))
parts.append(body_plain('（1）求油滴 a 和油滴 b 的质量之比。'))
parts.append(body_plain('（2）判断油滴 a、b 所带电荷的正负，并求二者所带电荷量绝对值之比。'))
parts.append(label('【答案】'))
parts.append(body_plain('（1）8∶1；（2）油滴 a 带负电，油滴 b 带正电，二者所带电荷量绝对值之比为 4∶1。'))
parts.append(label('【解析】'))
parts.append(body_plain('设油滴半径为 r，密度为 ρ，空气阻力大小为 f=λrv，其中 λ 为比例系数。'))
parts.append(body_plain('（1）求质量之比。'))
parts.append(body_plain('两板间未加电压时，油滴均做匀速直线运动。根据平衡条件可得'))
parts.append(eq(mv('m') + mv('g') + mop('=') + mv('λ') + mv('r') + mv('v', font='Book Antiqua')))
parts.append(body_plain('根据球体质量公式可得'))
parts.append(eq(mv('m') + mop('=') + ratio(mn('4'), mn('3')) + mv('π') + mv('ρ') + cube(mv('r'))))
parts.append(body_plain('联立可得'))
parts.append(eq(mv('v', font='Book Antiqua') + mop('=') + ratio(mn('4') + mv('π') + mv('ρ') + mv('g'), mn('3') + mv('λ')) + sq(mv('r'))))
parts.append(body_plain('因此，油滴的匀速运动速率与半径的平方成正比。根据速度关系可得'))
parts.append(eq(ratio(va(), vb()) + mop('=') + ratio(v0(), ratio(v0(), mn('4'))) + mop('=') + mn('4')))
parts.append(body_plain('根据速度与半径的关系可得'))
parts.append(eq(ratio(sq(msub(mv('r'), mu('a'))), sq(msub(mv('r'), mu('b')))) + mop('=') + mn('4')))
parts.append(body_plain('解得'))
parts.append(eq(ratio(msub(mv('r'), mu('a')), msub(mv('r'), mu('b'))) + mop('=') + mn('2')))
parts.append(body_plain('由于油滴质量与半径的三次方成正比，可得'))
parts.append(eq(
    ratio(msub(mv('m'), mu('a')), msub(mv('m'), mu('b'))) + mop('=') +
    cube(ratio(msub(mv('r'), mu('a')), msub(mv('r'), mu('b')))) +
    mop('=') + cube(mn('2')) + mop('=') + mn('8')
))
parts.append(body_plain('因此，油滴 a 和油滴 b 的质量之比为 8∶1。'))
parts.append(body_plain('（2）判断电性并求电荷量之比。'))
parts.append(body_plain('上板接正极，因此板间电场方向竖直向下。油滴 a 加电压后速率减小，说明电场力方向竖直向上，所以油滴 a 带负电；油滴 b 加电压后速率增大，说明电场力方向竖直向下，所以油滴 b 带正电。'))
parts.append(body_plain('对于油滴 a，加电压后再次匀速下降。根据平衡条件可得'))
parts.append(eq(
    msub(mv('m'), mu('a')) + mv('g') + mop('=') +
    mv('λ') + msub(mv('r'), mu('a')) + ratio(v0(), mn('2')) +
    mop('+') + mop('|') + msub(mv('q'), mu('a')) + mop('|') + mv('E')
))
parts.append(body_plain('未加电压时有'))
parts.append(eq(msub(mv('m'), mu('a')) + mv('g') + mop('=') + mv('λ') + msub(mv('r'), mu('a')) + v0()))
parts.append(body_plain('两式相减可得'))
parts.append(eq(
    mop('|') + msub(mv('q'), mu('a')) + mop('|') + mv('E') + mop('=') +
    mv('λ') + msub(mv('r'), mu('a')) + ratio(v0(), mn('2'))
))
parts.append(body_plain('对于油滴 b，加电压后再次匀速下降。根据平衡条件可得'))
parts.append(eq(
    mv('λ') + msub(mv('r'), mu('b')) + ratio(v0(), mn('2')) + mop('=') +
    msub(mv('m'), mu('b')) + mv('g') + mop('+') +
    mop('|') + msub(mv('q'), mu('b')) + mop('|') + mv('E')
))
parts.append(body_plain('未加电压时有'))
parts.append(eq(msub(mv('m'), mu('b')) + mv('g') + mop('=') + mv('λ') + msub(mv('r'), mu('b')) + ratio(v0(), mn('4'))))
parts.append(body_plain('两式相减可得'))
parts.append(eq(
    mop('|') + msub(mv('q'), mu('b')) + mop('|') + mv('E') + mop('=') +
    mv('λ') + msub(mv('r'), mu('b')) + ratio(v0(), mn('4'))
))
parts.append(body_plain('两式相比可得'))
parts.append(eq(
    ratio(mop('|') + msub(mv('q'), mu('a')) + mop('|'),
          mop('|') + msub(mv('q'), mu('b')) + mop('|')) + mop('=') +
    ratio(mv('λ') + msub(mv('r'), mu('a')) + ratio(v0(), mn('2')),
          mv('λ') + msub(mv('r'), mu('b')) + ratio(v0(), mn('4'))) +
    mop('=') + mn('2') + ratio(msub(mv('r'), mu('a')), msub(mv('r'), mu('b'))) +
    mop('=') + mn('4')
))
parts.append(body_plain('因此，油滴 a、b 所带电荷量绝对值之比为 4∶1。'))
parts.append(label('【与课本例题的关联】'))
parts.append(body_plain('课本例题比较静电力和万有引力，本题比较电场力、重力与空气阻力。两题都需要先明确不同性质的力，再利用平衡关系或比值关系消去共同因素，并通过数量关系判断各种力的相对作用。课本例题体现微观粒子中重力可忽略，本题则说明对质量较大的带电油滴，重力不能忽略。'))

# 末尾质量检查提示
parts.append(heading('出版规范检查', page_break=True))
parts.append(body_plain('1. 物理量符号采用规范斜体，说明性下标采用正体，数字、单位和运算符采用正体。'))
parts.append(body_plain('2. 正文中的科学计数法采用 Word 真上标，不使用 Unicode 伪上标。'))
parts.append(body_plain('3. 独立公式采用可编辑 OMML 结构，分式使用横线分式。'))
parts.append(body_plain('4. 所有比值均为无量纲量，不附加单位。'))
parts.append(body_plain('5. 高考真题注明年份、卷别和题号，答案与解析均重新独立核算。'))

sect = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1247" w:right="1191" w:bottom="1247" w:left="1191" '
    'w:header="720" w:footer="720" w:gutter="0"/>'
    '</w:sectPr>'
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

print(OUT)
