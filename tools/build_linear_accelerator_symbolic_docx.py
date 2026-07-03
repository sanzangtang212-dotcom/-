from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

OUT = Path('downloads/多级直线加速器_一题三变_字母推导版.docx')
OUT.parent.mkdir(parents=True, exist_ok=True)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ---------- Word runs ----------
def wrun(text, font='宋体', size=24, bold=False, italic=False, sub=False, sup=False):
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
    return wrun(text, '宋体', size=size, bold=bold)


def hei(text, size=28):
    return wrun(text, '黑体', size=size, bold=True)


def roman(text, size=24, sub=False, sup=False):
    return wrun(text, 'Times New Roman', size=size, sub=sub, sup=sup)


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
    ppr.append(f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>')
    return f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>{inner}</w:p>'


def body(text, indent=True, after=70):
    return paragraph(cn(text), first_line=480 if indent else 0, after=after)


def heading(text, page_break=False):
    return paragraph(hei(text, 28), before=180, after=100, line=340,
                     page_break=page_break, keep_next=True)


def label(text):
    return paragraph(hei(text, 24), before=110, after=45, line=320, keep_next=True)

# ---------- OMML ----------
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


def mv(text, font='Times New Roman'):
    return mr(text, upright=False, font=font)


def mu(text, preserve=False):
    return mr(text, upright=True, preserve=preserve)


def mn(text):
    return mr(text, upright=True)


def mop(text):
    return mr(text, upright=True)


def msub(base, sub):
    return f'<m:sSub><m:sSubPr/><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>'


def msup(base, sup):
    return f'<m:sSup><m:sSupPr/><m:e>{base}</m:e><m:sup>{sup}</m:sup></m:sSup>'


def mfrac(num, den):
    return f'<m:f><m:fPr/><m:num>{num}</m:num><m:den>{den}</m:den></m:f>'


def mrad(expr):
    return '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>' + expr + '</m:e></m:rad>'


def mparen(expr):
    return f'<m:d><m:dPr/><m:e>{expr}</m:e></m:d>'


def eq(expr, after=80):
    return paragraph(f'<m:oMathPara><m:oMath>{expr}</m:oMath></m:oMathPara>',
                     align='center', before=35, after=after, line=300)


def V(sub=None):
    base = mv('v', font='Book Antiqua')
    return msub(base, mu(sub)) if sub is not None else base


def S(symbol, sub=None):
    base = mv(symbol)
    return msub(base, mu(sub)) if sub is not None else base


def S_expr(symbol, subexpr):
    return msub(mv(symbol), subexpr)


def SQ(expr):
    return msup(expr, mn('2'))


def FRAC(a, b):
    return mfrac(a, b)


parts = []
parts.append(paragraph(hei('多级直线加速器', 34), align='center', after=55))
parts.append(paragraph(hei('一题三变·字母推导版', 30), align='center', after=190))

# Core model
parts.append(heading('原例题核心模型'))
parts.append(body('电子在金属圆筒内部近似不受静电力，做匀速直线运动；只在相邻圆筒之间的间隙中受到静电力而加速。采用最短同步设计时，电子通过每个圆筒的时间均为交变电压周期的一半。'))
parts.append(body('设电子质量为m，电荷量的绝对值为e，交变电压的绝对值为u，周期为T。电子进入第n个圆筒前共经过n次加速。根据动能定理可得'))
parts.append(eq(S('K', 'n') + mop('=') + mv('n') + mv('e') + mv('u') + mop('=') + FRAC(mn('1'), mn('2')) + mv('m') + SQ(V('n'))))
parts.append(eq(V('n') + mop('=') + mrad(FRAC(mn('2') + mv('n') + mv('e') + mv('u'), mv('m')))))
parts.append(body('采用最短同步条件时，电子通过第n个圆筒的时间为'))
parts.append(eq(S('t', 'n') + mop('=') + FRAC(mv('T'), mn('2'))))
parts.append(body('因此，第n个圆筒的长度为'))
parts.append(eq(S('L', 'n') + mop('=') + FRAC(mv('T'), mn('2')) + mrad(FRAC(mn('2') + mv('n') + mv('e') + mv('u'), mv('m'))) + mop('=') + mrad(mv('n')) + S('L', '1')))
parts.append(body('所以各圆筒长度不构成等差数列，而满足平方根规律。'))
parts.append(eq(S('L', '1') + mop(':') + S('L', '2') + mop(':') + S('L', '3') + mop(':') + mop('⋯') + mop(':') + S('L', 'n') + mop('=') + mn('1') + mop(':') + mrad(mn('2')) + mop(':') + mrad(mn('3')) + mop(':') + mop('⋯') + mop(':') + mrad(mv('n'))))

# Variation 1
parts.append(heading('一、基础变式——同模型规律巩固', page_break=True))
parts.append(label('【题目】'))
parts.append(body('某电子直线加速器由多个同轴金属圆筒组成，电子从静止开始运动。交变电压的绝对值为u，周期为T。电子通过圆筒间隙的时间忽略不计，且每次通过间隙时均受到加速。已知第1个圆筒的长度为L。'))
parts.append(body('（1）求电子进入第n个圆筒时的速度。'))
parts.append(body('（2）用L和n表示第n个圆筒的长度。'))
parts.append(body('（3）求第p个圆筒与第q个圆筒的长度之比。'))
parts.append(body('（4）求电子进入第p个圆筒与进入第q个圆筒时的动能之比。'))
parts.append(body('（5）电子依次通过前N个圆筒所用的时间是多少？'))
parts.append(label('【答案】'))
parts.append(eq(V('n') + mop('=') + mrad(FRAC(mn('2') + mv('n') + mv('e') + mv('u'), mv('m')))))
parts.append(eq(S('L', 'n') + mop('=') + mrad(mv('n')) + mv('L')))
parts.append(eq(S('L', 'p') + mop(':') + S('L', 'q') + mop('=') + mrad(mv('p')) + mop(':') + mrad(mv('q'))))
parts.append(eq(S('K', 'p') + mop(':') + S('K', 'q') + mop('=') + mv('p') + mop(':') + mv('q')))
parts.append(eq(mv('t') + mop('=') + FRAC(mv('N') + mv('T'), mn('2'))))
parts.append(label('【解析】'))
parts.append(body('电子进入第n个圆筒前共经过n个加速间隙。根据动能定理可得'))
parts.append(eq(mv('n') + mv('e') + mv('u') + mop('=') + FRAC(mn('1'), mn('2')) + mv('m') + SQ(V('n'))))
parts.append(body('解得'))
parts.append(eq(V('n') + mop('=') + mrad(FRAC(mn('2') + mv('n') + mv('e') + mv('u'), mv('m')))))
parts.append(body('电子通过各圆筒的时间均为T/2，因此圆筒长度与电子进入该圆筒时的速度成正比。又因为第1个圆筒的长度为L，所以'))
parts.append(eq(S('L', 'n') + mop('=') + mrad(mv('n')) + mv('L')))
parts.append(body('由此可得第p、q个圆筒的长度之比为'))
parts.append(eq(FRAC(S('L', 'p'), S('L', 'q')) + mop('=') + mrad(FRAC(mv('p'), mv('q')))))
parts.append(body('电子进入第n个圆筒时的动能为neu，因此动能之比等于加速次数之比。电子通过每个圆筒所需时间均为T/2，故通过前N个圆筒的总时间为NT/2。'))

# Variation 2
parts.append(heading('二、情境变式——粒子具有初速度', page_break=True))
parts.append(label('【题目】'))
parts.append(body('某带电粒子的质量为m，电荷量的绝对值为q，以给定初速度从序号为0的金属圆板中央进入直线加速器。交变电压的绝对值为U，周期为T。粒子每通过一个圆筒间隙，静电力均对其做正功，且通过间隙的时间忽略不计。'))
parts.append(body('（1）求粒子进入第n个圆筒时的速度。'))
parts.append(body('（2）采用最短同步条件，求第n个圆筒的长度。'))
parts.append(body('（3）证明圆筒长度的平方随序号n均匀增加。'))
parts.append(body('（4）求相邻两个圆筒长度平方之差。'))
parts.append(label('【答案】'))
parts.append(eq(V('n') + mop('=') + mrad(SQ(V('0')) + mop('+') + FRAC(mn('2') + mv('n') + mv('q') + mv('U'), mv('m')))))
parts.append(eq(S('L', 'n') + mop('=') + FRAC(mv('T'), mn('2')) + mrad(SQ(V('0')) + mop('+') + FRAC(mn('2') + mv('n') + mv('q') + mv('U'), mv('m')))))
parts.append(eq(SQ(S('L', 'n')) + mop('=') + FRAC(SQ(mv('T')) + SQ(V('0')), mn('4')) + mop('+') + FRAC(SQ(mv('T')) + mv('q') + mv('U'), mn('2') + mv('m')) + mv('n')))
parts.append(eq(SQ(S_expr('L', mv('n') + mop('+') + mn('1'))) + mop('−') + SQ(S('L', 'n')) + mop('=') + FRAC(SQ(mv('T')) + mv('q') + mv('U'), mn('2') + mv('m'))))
parts.append(label('【解析】'))
parts.append(body('粒子进入第n个圆筒前共经过n次加速。根据动能定理可得'))
parts.append(eq(mv('n') + mv('q') + mv('U') + mop('=') + FRAC(mn('1'), mn('2')) + mv('m') + SQ(V('n')) + mop('−') + FRAC(mn('1'), mn('2')) + mv('m') + SQ(V('0'))))
parts.append(body('整理可得粒子进入第n个圆筒时的速度。采用最短同步条件时，粒子通过每个圆筒的时间为T/2，因此'))
parts.append(eq(S('L', 'n') + mop('=') + FRAC(mv('T'), mn('2')) + V('n')))
parts.append(body('将速度表达式代入并平方，可得'))
parts.append(eq(SQ(S('L', 'n')) + mop('=') + FRAC(SQ(mv('T')) + SQ(V('0')), mn('4')) + mop('+') + FRAC(SQ(mv('T')) + mv('q') + mv('U'), mn('2') + mv('m')) + mv('n')))
parts.append(body('该式是关于序号n的一次函数，所以圆筒长度的平方构成等差数列；相邻两项之差为常量。'))

# Variation 3
parts.append(heading('三、素养提升——同一加速器加速不同粒子', page_break=True))
parts.append(label('【题目】'))
parts.append(body('某直线加速器原来用于加速粒子甲。粒子甲的质量、电荷量绝对值、交变电压绝对值和周期均在公式中用下标1表示。现用同一组圆筒加速粒子乙，粒子乙的相应物理量均在公式中用下标2表示。两种粒子的初速度均忽略不计。'))
parts.append(body('（1）推导粒子乙能够使用原有圆筒同步加速的条件。'))
parts.append(body('（2）若两种情况下交变电压保持不变，求两个周期的关系。'))
parts.append(body('（3）若两种情况下交变电压周期保持不变，求两个电压的关系。'))
parts.append(body('（4）两种粒子分别经过n次加速后，求其动能之比。'))
parts.append(label('【答案】'))
parts.append(eq(S('T', '2') + mrad(FRAC(S('q', '2') + S('U', '2'), S('m', '2'))) + mop('=') + S('T', '1') + mrad(FRAC(S('q', '1') + S('U', '1'), S('m', '1')))))
parts.append(eq(FRAC(S('T', '2'), S('T', '1')) + mop('=') + mrad(FRAC(S('m', '2') + S('q', '1'), S('m', '1') + S('q', '2')))))
parts.append(eq(FRAC(S('U', '2'), S('U', '1')) + mop('=') + FRAC(S('m', '2') + S('q', '1'), S('m', '1') + S('q', '2'))))
parts.append(eq(FRAC(S('K', '2'), S('K', '1')) + mop('=') + FRAC(S('q', '2') + S('U', '2'), S('q', '1') + S('U', '1'))))
parts.append(label('【解析】'))
parts.append(body('粒子甲进入第n个圆筒时的速度为'))
parts.append(eq(S_expr('v', mu('1n')) + mop('=') + mrad(FRAC(mn('2') + mv('n') + S('q', '1') + S('U', '1'), S('m', '1')))))
parts.append(body('因此第n个圆筒的长度为'))
parts.append(eq(S('L', 'n') + mop('=') + FRAC(S('T', '1'), mn('2')) + mrad(FRAC(mn('2') + mv('n') + S('q', '1') + S('U', '1'), S('m', '1')))))
parts.append(body('粒子乙使用同一圆筒时，也必须满足相同的圆筒长度。于是'))
parts.append(eq(FRAC(S('T', '1'), mn('2')) + mrad(FRAC(mn('2') + mv('n') + S('q', '1') + S('U', '1'), S('m', '1'))) + mop('=') + FRAC(S('T', '2'), mn('2')) + mrad(FRAC(mn('2') + mv('n') + S('q', '2') + S('U', '2'), S('m', '2')))))
parts.append(body('约去相同因子后，即得到同步条件。若电压保持不变，可由同步条件解出周期关系；若周期保持不变，可解出电压关系。'))
parts.append(body('粒子经过n次加速后获得的动能为nqU，因此两种粒子的动能之比等于各自电荷量与加速电压乘积之比。'))

# Reminders
parts.append(heading('易错提醒', page_break=True))
parts.append(body('1. 粒子只在圆筒间隙中加速，在金属圆筒内部做匀速直线运动。', indent=False))
parts.append(body('2. 每次通过间隙后，粒子的动能增加量相同，而速度增加量并不相同。', indent=False))
parts.append(body('3. 由动能与加速次数成正比可知，速度与加速次数的平方根成正比，因此圆筒长度也按平方根规律增加。', indent=False))
parts.append(body('4. 取T/2是最短同步设计；理论上取半周期的奇数倍也能实现同步，但会使圆筒显著变长。', indent=False))

# ---------- Package ----------
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

# ---------- Validation ----------
assert OUT.exists() and OUT.stat().st_size > 4000
with ZipFile(OUT) as z:
    xml = z.read('word/document.xml').decode('utf-8')
    ET.fromstring(xml)
    ET.fromstring(z.read('word/styles.xml'))
    ET.fromstring(z.read('word/settings.xml'))
    forbidden = '⁰¹²³⁴⁵⁶⁷⁸⁹⁻₀₁₂₃₄₅₆₇₈₉ₑₚ'
    assert not any(ch in xml for ch in forbidden)
    assert xml.count('【题目】') == 3
    assert xml.count('【答案】') == 3
    assert xml.count('【解析】') == 3
    assert '<m:oMathPara>' in xml
    assert '<m:f>' in xml and '<m:rad>' in xml and '<m:sSub>' in xml and '<m:sSup>' in xml
    assert 'Book Antiqua' in xml and 'Times New Roman' in xml
    for text in ['同模型规律巩固', '粒子具有初速度', '同一加速器加速不同粒子', '圆筒长度的平方构成等差数列']:
        assert text in xml
    for bad in ['TODO', '待补充', 'XXXX', '�']:
        assert bad not in xml

print(OUT)
