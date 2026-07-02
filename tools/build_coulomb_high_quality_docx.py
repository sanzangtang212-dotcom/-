from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape

OUT = Path('downloads/一题三变_第九章_库仑定律_高质量修正版.docx')
OUT.parent.mkdir(parents=True, exist_ok=True)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def wr(text, bold=False, italic=False, font='宋体', size=24):
    text = escape(str(text))
    props = [
        f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>',
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    ]
    if bold:
        props.append('<w:b/>')
    if italic:
        props.append('<w:i/>')
    return f'<w:r><w:rPr>{"".join(props)}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def wp(inner='', style=None, align=None, before=0, after=100, line=360,
       page_break=False, keep_next=False):
    ppr = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    if keep_next:
        ppr.append('<w:keepNext/>')
    if page_break:
        ppr.append('<w:pageBreakBefore/>')
    ppr.append(
        f'<w:spacing w:before="{before}" w:after="{after}" '
        f'w:line="{line}" w:lineRule="auto"/>'
    )
    return f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>{inner}</w:p>'


def mrun(text, upright=False):
    """Create an OMML run.

    Variables are Times New Roman italic. Numbers, operators, units and
    descriptive subscripts are Times New Roman upright.
    """
    text = escape(str(text))
    style = 'p' if upright else 'i'
    return (
        '<m:r>'
        f'<m:rPr><m:sty m:val="{style}"/></m:rPr>'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" '
        'w:eastAsia="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>'
        f'<m:t xml:space="preserve">{text}</m:t>'
        '</m:r>'
    )


def var(text):
    return mrun(text, upright=False)


def up(text):
    return mrun(text, upright=True)


def num(text):
    return mrun(text, upright=True)


def op(text):
    return mrun(text, upright=True)


def unit(text):
    return mrun(text, upright=True)


def msup(base, exponent):
    return f'<m:sSup><m:sSupPr/><m:e>{base}</m:e><m:sup>{exponent}</m:sup></m:sSup>'


def msub(base, subscript):
    return f'<m:sSub><m:sSubPr/><m:e>{base}</m:e><m:sub>{subscript}</m:sub></m:sSub>'


def mfrac(numerator, denominator):
    return f'<m:f><m:fPr/><m:num>{numerator}</m:num><m:den>{denominator}</m:den></m:f>'


def math_p(expr):
    return wp(
        f'<m:oMathPara><m:oMath>{expr}</m:oMath></m:oMathPara>',
        align='center', before=40, after=80, line=300
    )


def label(text):
    return wp(wr(text, bold=True, font='黑体', size=24), before=120, after=50, line=320)


def body(text):
    return wp(wr(text, font='宋体', size=24), after=70, line=360)


def heading(text, page_break=False):
    return wp(
        wr(text, bold=True, font='黑体', size=28),
        style='Heading2', before=180, after=100, line=340,
        page_break=page_break, keep_next=True
    )


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
        cells = []
        for j, cell in enumerate(row):
            shade = '<w:shd w:val="clear" w:fill="EAF2F8"/>' if i == 0 else ''
            tcpr = (
                f'<w:tcW w:w="{widths[j]}" w:type="dxa"/>'
                f'{shade}<w:vAlign w:val="center"/>'
            )
            run = wr(cell, bold=(i == 0), font='宋体', size=22)
            cells.append(
                f'<w:tc><w:tcPr>{tcpr}</w:tcPr>'
                f'{wp(run, align="center", after=0, line=300)}</w:tc>'
            )
        trs.append(f'<w:tr>{"".join(cells)}</w:tr>')
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        f'<w:jc w:val="center"/>{borders}</w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>{"".join(trs)}</w:tbl>'
    )


def F(subscript):
    # e、G、eD、GD 等都是说明性角标，统一使用正体真下标。
    return msub(var('F'), up(subscript))


def mass(subscript):
    return msub(var('m'), up(subscript))


def Rsym(subscript):
    return msub(var('R'), up(subscript))


def r0():
    return msub(var('r'), num('0'))


def sq(expr):
    return msup(expr, num('2'))


def sci(coef, exponent):
    return num(coef) + op('×') + msup(num('10'), num(exponent))


def ratio(numerator, denominator):
    return mfrac(numerator, denominator)


parts = []
parts.append(wp(wr('第九章  静电场及其应用', bold=True, font='黑体', size=34), align='center', after=60))
parts.append(wp(wr('2. 库仑定律', bold=True, font='黑体', size=30), align='center', after=80))
parts.append(wp(wr('一题三变｜微观粒子间静电力与万有引力的比较', bold=True, font='黑体', size=28), align='center', after=220))
parts.append(body('核心模型：微观带电粒子之间同时存在静电力和万有引力。'))
parts.append(body('核心方法：分别应用库仑定律和万有引力定律，再通过比值法消去共同因素。'))
parts.append(body('训练路径：直接计算 → 比例迁移 → 一般化推导与表格综合。'))
parts.append(body('难度控制：基础变式与课本例题相当；情境变式突出比例推理；素养提升突出模型概括与多表征分析。'))

# 基础变式
parts.append(heading('▌基础变式——同模型巩固'))
parts.append(label('【题目】'))
parts.append(body('在氢原子中，氢原子核可视为质子。质子与电子之间的距离为 5.3×10⁻¹¹ m。已知元电荷 e=1.60×10⁻¹⁹ C，质子质量 mₚ=1.67×10⁻²⁷ kg，电子质量 mₑ=9.11×10⁻³¹ kg，静电力常量 k=9.0×10⁹ N·m²/C²，万有引力常量 G=6.67×10⁻¹¹ N·m²/kg²。'))
parts.append(body('（1）求质子与电子之间的静电力大小。'))
parts.append(body('（2）求质子与电子之间的万有引力大小。'))
parts.append(body('（3）求静电力与万有引力的比值，并说明研究氢原子内部电子运动时是否需要考虑万有引力。'))
parts.append(label('【答案】'))
parts.append(body('（1）8.2×10⁻⁸ N；（2）3.6×10⁻⁴⁷ N；（3）静电力约为万有引力的 2.3×10³⁹ 倍，万有引力可以忽略。'))
parts.append(label('【解析】'))
parts.append(body('（1）求质子与电子之间的静电力。'))
parts.append(body('根据库仑定律可得'))
parts.append(math_p(
    F('e') + op('=') + var('k') + ratio(sq(var('e')), sq(var('r')))
))
parts.append(body('代入数据可得'))
parts.append(math_p(
    F('e') + op('=') + op('(') + sci('9.0', '9') + op('×')
    + ratio(sq(op('(') + sci('1.60', '−19') + op(')')),
            sq(op('(') + sci('5.3', '−11') + op(')')))
    + op(')') + unit(' N') + op('=') + sci('8.2', '−8') + unit(' N')
))
parts.append(body('质子与电子带异种电荷，因此静电力表现为吸引力。'))
parts.append(body('（2）求质子与电子之间的万有引力。'))
parts.append(body('根据万有引力定律可得'))
parts.append(math_p(
    F('G') + op('=') + var('G')
    + ratio(mass('p') + mass('e'), sq(var('r')))
))
parts.append(body('代入数据可得'))
parts.append(math_p(
    F('G') + op('=') + op('(') + sci('6.67', '−11') + op('×')
    + ratio(sci('1.67', '−27') + op('×') + sci('9.11', '−31'),
            sq(op('(') + sci('5.3', '−11') + op(')')))
    + op(')') + unit(' N') + op('=') + sci('3.6', '−47') + unit(' N')
))
parts.append(body('（3）比较两种力。'))
parts.append(body('根据两种力的比值可得'))
parts.append(math_p(
    ratio(F('e'), F('G')) + op('=')
    + ratio(var('k') + ratio(sq(var('e')), sq(var('r'))),
            var('G') + ratio(mass('p') + mass('e'), sq(var('r'))))
    + op('=') + ratio(var('k') + sq(var('e')),
                       var('G') + mass('p') + mass('e'))
))
parts.append(math_p(
    ratio(F('e'), F('G')) + op('=')
    + ratio(sci('8.2', '−8'), sci('3.6', '−47'))
    + op('≈') + sci('2.3', '39')
))
parts.append(body('两种力的比值没有单位。由计算结果可知，静电力远大于万有引力，因此研究氢原子内部电子运动时可以忽略万有引力。'))
parts.append(body('命题意图：巩固库仑定律和万有引力定律的直接应用，突出“比值法”在数量级比较中的作用。'))

# 情境变式
parts.append(heading('▌情境变式——新情境迁移', page_break=True))
parts.append(label('【题目】'))
parts.append(body('氘是氢的同位素。氘核带电荷 +e，质量近似为 2mₚ。设氢原子中质子与电子的距离为 r₀，氘原子中氘核与电子的距离为 2r₀。分别用 FₑH、F_GH 表示氢原子中的静电力和万有引力，用 FₑD、F_GD 表示氘原子中的静电力和万有引力。'))
parts.append(body('（1）求 FₑD/FₑH。'))
parts.append(body('（2）求 F_GD/F_GH。'))
parts.append(body('（3）求氘原子中静电力与万有引力的比值 R_D 与氢原子中对应比值 R_H 的关系，并估算 R_D。'))
parts.append(label('【答案】'))
parts.append(body('（1）1/4；（2）1/2；（3）R_D=R_H/2≈1.1×10³⁹。'))
parts.append(label('【解析】'))
parts.append(body('（1）比较两种原子中的静电力。'))
parts.append(body('根据库仑定律可得'))
parts.append(math_p(
    ratio(F('eD'), F('eH')) + op('=')
    + ratio(var('k') + ratio(sq(var('e')), sq(num('2') + r0())),
            var('k') + ratio(sq(var('e')), sq(r0())))
    + op('=') + ratio(num('1'), num('4'))
))
parts.append(body('（2）比较两种原子中的万有引力。'))
parts.append(body('根据万有引力定律可得'))
parts.append(math_p(
    ratio(F('GD'), F('GH')) + op('=')
    + ratio(var('G') + ratio(op('(') + num('2') + mass('p') + op(')') + mass('e'), sq(num('2') + r0())),
            var('G') + ratio(mass('p') + mass('e'), sq(r0())))
    + op('=') + ratio(num('1'), num('2'))
))
parts.append(body('（3）比较两种力的比值。'))
parts.append(body('根据比值的定义可得'))
parts.append(math_p(
    ratio(Rsym('D'), Rsym('H')) + op('=')
    + ratio(ratio(F('eD'), F('GD')), ratio(F('eH'), F('GH')))
    + op('=')
    + ratio(ratio(F('eD'), F('eH')), ratio(F('GD'), F('GH')))
    + op('=') + ratio(ratio(num('1'), num('4')), ratio(num('1'), num('2')))
    + op('=') + ratio(num('1'), num('2'))
))
parts.append(math_p(
    Rsym('D') + op('=') + ratio(num('1'), num('2')) + Rsym('H')
    + op('=') + op('(') + ratio(num('1'), num('2')) + op('×') + sci('2.3', '39') + op(')')
    + op('=') + sci('1.15', '39') + op('≈') + sci('1.1', '39')
))
parts.append(body('虽然氘核质量增大使万有引力相对增强，但静电力仍远大于万有引力。'))
parts.append(body('命题意图：避免重复代入大、小数量级数据，改用比例推理迁移同一物理模型。'))

# 素养提升
parts.append(heading('▌素养提升——多表征综合', page_break=True))
parts.append(label('【题目】'))
parts.append(body('某类“类氢离子”由一个带电荷 +Ze、质量近似为 Amₚ 的原子核和一个电子组成。设核与电子之间的距离为 λr₀，其中 r₀ 为氢原子中质子与电子的距离。定义静电力与万有引力的比值为 R=Fₑ/F_G，氢原子中的对应比值为 R_H。'))
parts.append(body('下表给出了四种粒子系统的参数。'))
parts.append(table([
    ['粒子系统', 'Z', 'A', 'λ', 'R/R_H'],
    ['氢原子 H', '1', '1', '1', ''],
    ['氘原子 D', '1', '2', '2', ''],
    ['氦离子 He⁺', '2', '4', '1/2', ''],
    ['锂离子 Li²⁺', '3', '7', '3', ''],
], [2200, 900, 900, 1100, 1800]))
parts.append(body('（1）推导一般情况下 R 与 R_H、Z、A、λ 的关系。'))
parts.append(body('（2）在表格最后一列填出各系统的 R/R_H。'))
parts.append(body('（3）上述系统中，哪一种系统里万有引力的相对作用最明显？说明判断依据。'))
parts.append(body('（4）已知 R_H≈2.3×10³⁹，估算锂离子 Li²⁺ 中的 R，并判断研究电子运动时能否忽略万有引力。'))
parts.append(label('【答案】'))
parts.append(body('（1）R=(Z/A)R_H，与 λ 无关；（2）依次为 1、1/2、1/2、3/7；（3）Li²⁺，因为其 R 最小；（4）R≈9.9×10³⁸，仍可忽略万有引力。'))
parts.append(label('【解析】'))
parts.append(body('（1）推导一般关系。'))
parts.append(body('根据库仑定律可得'))
parts.append(math_p(
    F('e') + op('=') + var('k')
    + ratio(var('Z') + sq(var('e')), sq(var('λ') + r0()))
))
parts.append(body('根据万有引力定律可得'))
parts.append(math_p(
    F('G') + op('=') + var('G')
    + ratio(var('A') + mass('p') + mass('e'), sq(var('λ') + r0()))
))
parts.append(body('两式相除可得'))
parts.append(math_p(
    var('R') + op('=') + ratio(F('e'), F('G'))
    + op('=') + ratio(var('k') + var('Z') + sq(var('e')),
                       var('G') + var('A') + mass('p') + mass('e'))
    + op('=') + ratio(var('Z'), var('A')) + Rsym('H')
))
parts.append(body('因此，R 与距离因子 λ 无关，只由核电荷数 Z 与质量数 A 的比值决定。'))
parts.append(body('（2）根据 R/R_H=Z/A，可得：H 为 1，D 为 1/2，He⁺ 为 1/2，Li²⁺ 为 3/7。'))
parts.append(body('（3）R 越小，说明万有引力相对于静电力越明显。四种系统中 Li²⁺ 的 3/7 最小，因此其万有引力的相对作用最明显。'))
parts.append(body('（4）估算 Li²⁺ 中的两力之比。'))
parts.append(body('根据 R=(Z/A)R_H 可得'))
parts.append(math_p(
    var('R') + op('=') + op('(') + ratio(num('3'), num('7'))
    + op('×') + sci('2.3', '39') + op(')')
    + op('=') + sci('9.9', '38')
))
parts.append(body('虽然 Li²⁺ 中万有引力的相对作用比氢原子更明显，但 R 仍接近 10³⁹，因此研究电子运动时仍可忽略万有引力。'))
parts.append(body('命题意图：通过一般化公式、参数表格和比较判断，考查模型概括、比例推理与多表征信息处理能力。'))
parts.append(body('易错提醒：①比值 R=Fₑ/F_G 没有单位；②两种力都与距离平方成反比，距离因素在比值中完全消去；③比较不同系统时，应比较 Z/A，而不是只比较电荷量或质量。'))

sect = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1247" w:right="1191" w:bottom="1247" '
    'w:left="1191" w:header="720" w:footer="720" w:gutter="0"/>'
    '</w:sectPr>'
)
document_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:r="{R}">'
    f'<w:body>{"".join(parts)}{sect}</w:body></w:document>'
)

styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:ascii="黑体" w:eastAsia="黑体" w:hAnsi="黑体"/><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
</w:styles>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

with ZipFile(OUT, 'w', ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', root_rels)
    z.writestr('word/document.xml', document_xml)
    z.writestr('word/styles.xml', styles_xml)
    z.writestr('word/_rels/document.xml.rels', doc_rels)

print(OUT)
