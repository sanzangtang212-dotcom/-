from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape
from math import sqrt
import xml.etree.ElementTree as ET

OUT = Path('downloads/一题三变加高考真题_库仑定律例题2_出版规范版.docx')
OUT.parent.mkdir(parents=True, exist_ok=True)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


# ==============================
# 普通 Word 文本
# ==============================
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


def cn(text, bold=False, size=24):
    return wrun(text, font='宋体', size=size, bold=bold)


def hei(text, size=28):
    return wrun(text, font='黑体', size=size, bold=True)


def roman(text, size=24, sub=False, sup=False):
    return wrun(text, font='Times New Roman', size=size, sub=sub, sup=sup)


def phys(text, size=24):
    return wrun(text, font='Times New Roman', size=size, italic=True)


def Ftxt(subscript=None, size=24):
    out = phys('F', size=size)
    if subscript is not None:
        out += roman(subscript, size=size, sub=True)
    return out


def qtxt(subscript=None, size=24):
    out = phys('q', size=size)
    if subscript is not None:
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


# ==============================
# 可编辑 OMML 公式
# ==============================
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


def mrad(expr):
    return (
        '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
        f'<m:deg/><m:e>{expr}</m:e></m:rad>'
    )


def mparen(expr):
    return f'<m:d><m:dPr/><m:e>{expr}</m:e></m:d>'


def eq(expr, after=80):
    return paragraph(
        f'<m:oMathPara><m:oMath>{expr}</m:oMath></m:oMathPara>',
        align='center', before=35, after=after, line=300
    )


def F(sub=None):
    return msub(mv('F'), mu(sub)) if sub is not None else mv('F')


def q(sub=None):
    return msub(mv('q'), mu(sub)) if sub is not None else mv('q')


def rsub(sub):
    return msub(mv('r'), mu(sub))


def sq(expr):
    return msup(expr, mn('2'))


def power10(exponent):
    return msup(mn('10'), mu(str(exponent).replace('-', '−')))


def sci(coef, exponent):
    return mn(coef) + mop('×') + power10(exponent)


def ratio(numerator, denominator):
    return mfrac(numerator, denominator)


def unit(text):
    return mu(' ' + text, preserve=True)


def cos30():
    return mu('cos') + mn('30') + mu('°')


def cos60():
    return mu('cos') + mn('60') + mu('°')


def sin60():
    return mu('sin') + mn('60') + mu('°')


# ==============================
# 正文
# ==============================
parts = []
parts.append(paragraph(hei('第九章  静电场及其应用', size=34), align='center', after=55))
parts.append(paragraph(hei('2. 库仑定律', size=30), align='center', after=75))
parts.append(paragraph(hei('例题2·一题三变＋高考真题', size=28), align='center', after=200))

parts.append(heading('原例题模型提炼'))
parts.append(body(
    cn('原例题中，三个电荷量均为 '), sci_text('2.0', '-6', 'C'),
    cn(' 的正点电荷位于边长为 '), roman('0.50 m'), cn(' 的等边三角形三个顶点。')
))
parts.append(plain('对任一顶点电荷，其余两个点电荷分别产生大小相等、夹角为60°的斥力。'))
parts.append(plain('根据库仑定律，任意两个点电荷之间的静电力大小为'))
parts.append(eq(
    F('0') + mop('=') + mv('k') + ratio(sq(q()), sq(mv('L'))) + mop('=') +
    mparen(sci('9.0', '9') + mop('×') + ratio(sq(mparen(sci('2.0', '-6'))), sq(mn('0.50')))) +
    unit('N') + mop('=') + mn('0.144') + unit('N')
))
parts.append(plain('根据两个等大力的合成规律可得'))
parts.append(eq(
    F() + mop('=') + mn('2') + F('0') + cos30() + mop('=') +
    mrad(mn('3')) + F('0') + mop('≈') + mn('0.25') + unit('N')
))
parts.append(plain('方向沿该顶点角的外角平分线，垂直于对边并背离三角形中心。'))

# 基础变式
parts.append(heading('一、基础变式——同模型巩固'))
parts.append(label('【题目】'))
parts.append(body(
    cn('真空中，三个电荷量均为 '), sci_text('1.0', '-6', 'C'),
    cn(' 的正点电荷固定在边长为 '), roman('0.30 m'),
    cn(' 的等边三角形 '), roman('ABC'), cn(' 的三个顶点。求位于顶点 '), roman('A'),
    cn(' 的点电荷所受静电力的大小和方向。')
))
parts.append(label('【答案】'))
parts.append(body(cn('静电力大小为 '), roman('0.17 N'),
                  cn('，方向沿顶角 '), roman('A'), cn(' 的外角平分线，垂直于 '), roman('BC'),
                  cn(' 并背离三角形中心。')))
parts.append(label('【解析】'))
parts.append(plain('顶点B、C处的两个点电荷对A处点电荷产生的静电力大小相等，夹角为60°。'))
parts.append(plain('根据库仑定律可得每一个分力的大小'))
parts.append(eq(
    F('0') + mop('=') + mv('k') + ratio(sq(q()), sq(mv('L'))) + mop('=') +
    mparen(sci('9.0', '9') + mop('×') + ratio(sq(mparen(sci('1.0', '-6'))), sq(mn('0.30')))) +
    unit('N') + mop('=') + mn('0.10') + unit('N')
))
parts.append(plain('根据两个等大力的合成规律可得合力大小'))
parts.append(eq(
    F() + mop('=') + mn('2') + F('0') + cos30() + mop('=') +
    mrad(mn('3')) + F('0') + mop('=') + mn('0.17') + unit('N')
))
parts.append(plain('两个分力关于顶角A的角平分线对称，因此合力沿该角平分线向三角形外侧。'))

# 情境变式
parts.append(heading('二、情境变式——电性改变后的方向迁移', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('在上题相同的等边三角形 '), roman('ABC'), cn(' 中，顶点 '), roman('B'), cn('、'), roman('C'),
    cn(' 处固定电荷量为 '), roman('+'), qtxt(), cn(' 的点电荷，顶点 '), roman('A'),
    cn(' 处固定电荷量为 '), roman('−'), qtxt(), cn(' 的点电荷。已知 '), qtxt(), roman('='),
    sci_text('1.0', '-6', 'C'), cn('，三角形边长为 '), roman('0.30 m'), cn('。')
))
parts.append(plain('（1）求A处点电荷所受静电力的大小和方向。'))
parts.append(plain('（2）将结果与基础变式比较，说明电性改变对合力大小和方向的影响。'))
parts.append(label('【答案】'))
parts.append(plain('（1）静电力大小为0.17 N，方向沿A点角平分线指向BC边中点。'))
parts.append(plain('（2）合力大小不变，方向与基础变式中的合力方向相反。'))
parts.append(label('【解析】'))
parts.append(plain('A处为负电荷，B、C处为正电荷，因此B、C处电荷对A处电荷的作用力均为吸引力。'))
parts.append(plain('两个分力的大小仍为'))
parts.append(eq(F('0') + mop('=') + mn('0.10') + unit('N')))
parts.append(plain('两个分力夹角仍为60°，根据力的合成可得'))
parts.append(eq(
    F() + mop('=') + mn('2') + F('0') + cos30() + mop('=') + mn('0.17') + unit('N')
))
parts.append(plain('由于两个分力均指向B、C两点，合力沿A点角平分线指向BC边中点。电荷量大小和距离均未改变，所以合力大小不变；静电力由斥力变为引力，所以方向反向。'))

# 素养提升
parts.append(heading('三、素养提升——中心电荷与系统平衡', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('三个电荷量均为 '), roman('+'), qtxt(), cn(' 的点电荷分别位于边长为 '), phys('L'),
    cn(' 的等边三角形三个顶点。现把一个电荷量为 '), phys('Q'), cn(' 的点电荷放在三角形中心 '),
    roman('O'), cn(' 点。')
))
parts.append(plain('（1）若要求每个顶点电荷所受静电力的合力为零，求Q与q的关系，并判断Q的正负。'))
parts.append(body(cn('（2）当 '), qtxt(), roman('='), sci_text('2.0', '-6', 'C'), cn(' 时，求 '), phys('Q'), cn(' 的数值。')))
parts.append(plain('（3）说明所得Q与三角形边长L是否有关，并判断O点电荷所受合力是否为零。'))
parts.append(label('【答案】'))
parts.append(body(cn('（1）'), phys('Q'), roman('=−'), qtxt(), roman('/√3'), cn('；'), phys('Q'), cn(' 为负电荷。')))
parts.append(body(cn('（2）'), phys('Q'), roman('≈−'), sci_text('1.2', '-6', 'C'), cn('。')))
parts.append(plain('（3）Q与L无关；O点电荷受到三个顶点电荷的作用力大小相等、方向互成120°，合力为零。'))
parts.append(label('【解析】'))
parts.append(plain('以顶点A处电荷为研究对象。其余两个顶点电荷产生的两个斥力大小相等，夹角为60°。'))
parts.append(plain('根据库仑定律和力的合成规律可得两个顶点电荷产生的合力大小'))
parts.append(eq(
    F('v') + mop('=') + mrad(mn('3')) + mv('k') + ratio(sq(q()), sq(mv('L')))
))
parts.append(plain('等边三角形中心到顶点的距离为'))
parts.append(eq(mv('AO') + mop('=') + ratio(mv('L'), mrad(mn('3')))))
parts.append(plain('中心电荷对A处电荷的作用力必须指向O点，因此Q应为负电荷。根据库仑定律可得该力大小'))
parts.append(eq(
    F('O') + mop('=') + mv('k') + ratio(q() + mop('|') + mv('Q') + mop('|'),
                                             sq(ratio(mv('L'), mrad(mn('3'))))) +
    mop('=') + mn('3') + mv('k') + ratio(q() + mop('|') + mv('Q') + mop('|'), sq(mv('L')))
))
parts.append(plain('根据平衡条件可得'))
parts.append(eq(F('O') + mop('=') + F('v')))
parts.append(eq(
    mn('3') + mv('k') + ratio(q() + mop('|') + mv('Q') + mop('|'), sq(mv('L'))) +
    mop('=') + mrad(mn('3')) + mv('k') + ratio(sq(q()), sq(mv('L')))
))
parts.append(plain('约去共同因素后可得'))
parts.append(eq(mop('|') + mv('Q') + mop('|') + mop('=') + ratio(q(), mrad(mn('3')))))
parts.append(eq(mv('Q') + mop('=') + mop('−') + ratio(q(), mrad(mn('3')))))
parts.append(plain('代入q=2.0×10的−6次方C可得'))
parts.append(eq(
    mv('Q') + mop('=') + mop('−') + ratio(sci('2.0', '-6'), mrad(mn('3'))) + unit('C') +
    mop('=') + mop('−') + sci('1.15', '-6') + unit('C') + mop('≈') + mop('−') + sci('1.2', '-6') + unit('C')
))
parts.append(plain('推导过程中L被完全约去，所以Q与三角形边长无关。O点位于对称中心，三个顶点电荷对Q的作用力大小相等、方向互成120°，其矢量和为零。'))

# 高考真题
parts.append(heading('四、高考真题——库仑力方向与矢量合成', page_break=True))
parts.append(label('【真题来源】'))
parts.append(plain('2003年普通高等学校招生全国统一考试（全国卷·理科综合）第15题。原题图示选项在本稿中按方向关系文字化，物理条件和正确选项不变。'))
parts.append(label('【题目】'))
parts.append(plain('三个完全相同的金属小球a、b、c位于等边三角形的三个顶点。规定a、b位于下方水平边的左、右两端，c位于上方顶点。a和c带正电，b带负电，且a所带电荷量的绝对值小于b所带电荷量的绝对值。'))
parts.append(plain('c受到a和b的静电力的合力方向应为（　　）'))
parts.append(plain('A．沿ca的延长线斜向右上方'))
parts.append(plain('B．斜向右下方，位于水平向右方向与由c指向b的方向之间'))
parts.append(plain('C．沿水平方向由c指向右方'))
parts.append(plain('D．完全沿由c指向b的方向'))
parts.append(label('【答案】'))
parts.append(plain('B。'))
parts.append(label('【解析】'))
parts.append(body(
    cn('设a、b、c三球所带电荷量分别为 '), qtxt('a'), cn('、'), qtxt('b'), cn('、'), qtxt('c'),
    cn('，等边三角形边长为 '), phys('L'), cn('。')
))
parts.append(plain('a、c带同种电荷，因此a对c的静电力为斥力，方向沿ac连线并背离a；b、c带异种电荷，因此b对c的静电力为引力，方向由c指向b。'))
parts.append(plain('根据库仑定律可得两个分力的大小'))
parts.append(eq(
    F('a') + mop('=') + mv('k') + ratio(mop('|') + q('a') + q('c') + mop('|'), sq(mv('L')))
))
parts.append(eq(
    F('b') + mop('=') + mv('k') + ratio(mop('|') + q('b') + q('c') + mop('|'), sq(mv('L')))
))
parts.append(body(
    cn('由于 '), roman('|'), qtxt('a'), roman('|<|'), qtxt('b'), roman('|'), cn('，所以 '), Ftxt('a'), roman('<'), Ftxt('b'), cn('。')
))
parts.append(plain('取水平方向向右为x轴正方向，竖直向上为y轴正方向。根据力的分解可得'))
parts.append(eq(
    F('x') + mop('=') + mparen(F('a') + mop('+') + F('b')) + cos60() + mop('>') + mn('0')
))
parts.append(eq(
    F('y') + mop('=') + mparen(F('a') + mop('−') + F('b')) + sin60() + mop('<') + mn('0')
))
parts.append(plain('因此合力的水平分量向右、竖直分量向下，合力斜向右下方。由于b对c的作用力较大，合力方向更靠近由c指向b的方向，但不与cb完全重合，故选B。'))
parts.append(label('【与原例题的关联】'))
parts.append(plain('原例题和本题都需要先分别判断各个库仑力的方向，再比较分力大小并进行矢量合成。原例题中的两个分力大小相等，可直接利用对称性求合力；本题中的两个分力大小不等，需要进一步通过分量符号或平行四边形定则判断合力所在区域。'))

# ==============================
# OOXML 打包
# ==============================
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


# ==============================
# 自检：内容、计算、结构、出版排版
# ==============================
k = 9.0e9
assert abs(k * (2.0e-6) ** 2 / 0.50 ** 2 - 0.144) < 1e-12
assert abs(sqrt(3) * 0.144 - 0.249415316) < 1e-6
assert abs(k * (1.0e-6) ** 2 / 0.30 ** 2 - 0.10) < 1e-12
assert abs(sqrt(3) * 0.10 - 0.173205081) < 1e-6
assert abs(2.0e-6 / sqrt(3) - 1.154700538e-6) < 1e-14

assert OUT.exists() and OUT.stat().st_size > 5000
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

    # 禁止 Unicode 伪上下标，避免角标空格异常
    forbidden = '⁰¹²³⁴⁵⁶⁷⁸⁹⁻₀₁₂₃₄₅₆₇₈₉ₑₚ'
    assert not any(ch in xml for ch in forbidden)

    # 真上下标与可编辑公式
    assert '<w:vertAlign w:val="superscript"/>' in xml
    assert '<w:vertAlign w:val="subscript"/>' in xml
    assert '<m:sSub>' in xml
    assert '<m:sSup>' in xml
    assert '<m:f>' in xml
    assert '<m:oMathPara>' in xml
    assert '<m:rad>' in xml

    # 固定结构与高考真题核验字段
    assert xml.count('【题目】') == 4
    assert xml.count('【答案】') == 4
    assert xml.count('【解析】') == 4
    assert '2003年普通高等学校招生全国统一考试' in xml
    assert '故选B' in xml
    assert '0.17 N' in xml
    assert 'Q与L无关' in xml

print(OUT)
