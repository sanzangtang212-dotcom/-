from pathlib import Path

BASE = Path('tools/build_potential_difference_regenerated_v2.py')
source = BASE.read_text(encoding='utf-8')
source = source.replace(
    '一题三变加高考真题_第十章_电势差_重新生成自检版.docx',
    '一题三变加高考真题_第十章_电势差_全题精修版.docx',
)

start = source.index('# ---------- 基础变式 ----------')
end = source.index('# ---------- 打包 ----------')

replacement = r'''# ---------- 基础变式 ----------
parts.append(heading('一、基础变式——同模型递进计算', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('在匀强电场中，把电荷量为 '), sci_text('3.0', '-9', 'C'), cn(' 的正点电荷从 '), roman('A'),
    cn(' 点移动到 '), roman('B'), cn(' 点，静电力做功为 '), sci_text('1.8', '-7', 'J'),
    cn('；再把该电荷从 '), roman('B'), cn(' 点移动到 '), roman('C'),
    cn(' 点，静电力做功为 '), sci_text('−4.5', '-7', 'J'), cn('。')
))
parts.append(plain('（1）比较A、B、C三点电势的高低。'))
parts.append(plain('（2）求A、B间，B、C间和A、C间的电势差。'))
parts.append(body(cn('（3）把电荷量为 '), sci_text('2.0', '-9', 'C'), cn(' 的正点电荷从A点移动到C点，求静电力做的功。'))
parts.append(plain('（4）若A、B、C三点位于同一条电场线上，定性画出一种可能的位置关系，并求CA与AB的距离之比。'))
parts.append(label('【答案】'))
parts.append(body(cn('（1）'), inline_sub('φ', 'C'), roman('>'), inline_sub('φ', 'A'), roman('>'), inline_sub('φ', 'B'), cn('。'))
parts.append(body(cn('（2）'), inline_sub('U', 'AB'), roman('=60 V，'), inline_sub('U', 'BC'), roman('=−150 V，'), inline_sub('U', 'AC'), roman('=−90 V。'))
parts.append(body(cn('（3）'), sci_text('−1.8', '-7', 'J'), cn('。 （4）电场方向由C指向B，CA∶AB=3∶2。'))
parts.append(label('【解析】'))
parts.append(plain('根据静电力做功与电势差的关系可得'))
parts.append(eq(U('AB') + mop('=') + FRAC(SCI('1.8', '-7'), SCI('3.0', '-9')) + UNIT('V') + mop('=') + mn('60') + UNIT('V')))
parts.append(eq(U('BC') + mop('=') + FRAC(mop('−') + SCI('4.5', '-7'), SCI('3.0', '-9')) + UNIT('V') + mop('=') + mop('−') + mn('150') + UNIT('V')))
parts.append(plain('由第一段电势差为正可知A点电势高于B点；由第二段电势差为负可知C点电势高于B点。取B点电势为0，可得A点电势为60 V，C点电势为150 V。'))
parts.append(eq(U('AC') + mop('=') + U('AB') + mop('+') + U('BC') + mop('=') + mop('−') + mn('90') + UNIT('V')))
parts.append(plain('根据静电力做功公式可得'))
parts.append(eq(WW('AC') + mop('=') + mparen(SCI('2.0', '-9') + mop('×') + mparen(mop('−') + mn('90'))) + UNIT('J') + mop('=') + mop('−') + SCI('1.8', '-7') + UNIT('J')))
parts.append(para(mono('C（150 V）──────A（60 V）────B（0 V）      E →'), align='center', after=45, line=300))
parts.append(plain('在匀强电场中，沿电场方向电势降低，且电势差大小与沿电场方向的距离成正比，因此CA∶AB=90∶60=3∶2。'))

# ---------- 情境变式 ----------
parts.append(heading('二、情境变式——电子束能量标定', page_break=True))
parts.append(label('【题目】'))
parts.append(body(
    cn('在电子束能量标定装置中，一个电子的电荷量为 '), sci_text('−1.60', '-19', 'C'),
    cn('。电子从A点移动到B点时，静电力做功为 '), sci_text('3.2', '-18', 'J'),
    cn('；从B点移动到C点时，静电力做功为 '), sci_text('−4.8', '-18', 'J'), cn('。')
))
parts.append(plain('（1）求A、B间，B、C间和A、C间的电势差。'))
parts.append(plain('（2）比较A、B、C三点电势的高低。'))
parts.append(plain('（3）求电子从A点移动到C点时静电力做的功，并判断电子的电势能如何变化。'))
parts.append(body(cn('（4）若换成电荷量为 '), sci_text('1.60', '-19', 'C'), cn(' 的质子从A点移动到C点，静电力做功为多少？'))
parts.append(label('【答案】'))
parts.append(body(cn('（1）'), inline_sub('U', 'AB'), roman('=−20 V，'), inline_sub('U', 'BC'), roman('=30 V，'), inline_sub('U', 'AC'), roman('=10 V。'))
parts.append(body(cn('（2）'), inline_sub('φ', 'B'), roman('>'), inline_sub('φ', 'A'), roman('>'), inline_sub('φ', 'C'), cn('。'))
parts.append(body(cn('（3）'), sci_text('−1.6', '-18', 'J'), cn('，电子的电势能增加 '), sci_text('1.6', '-18', 'J'), cn('。'))
parts.append(body(cn('（4）'), sci_text('1.6', '-18', 'J'), cn('。'))
parts.append(label('【解析】'))
parts.append(plain('电子带负电，使用电势差定义式时必须把电荷量的负号一并代入。'))
parts.append(eq(U('AB') + mop('=') + FRAC(SCI('3.2', '-18'), mop('−') + SCI('1.60', '-19')) + UNIT('V') + mop('=') + mop('−') + mn('20') + UNIT('V')))
parts.append(eq(U('BC') + mop('=') + FRAC(mop('−') + SCI('4.8', '-18'), mop('−') + SCI('1.60', '-19')) + UNIT('V') + mop('=') + mn('30') + UNIT('V')))
parts.append(eq(U('AC') + mop('=') + U('AB') + mop('+') + U('BC') + mop('=') + mn('10') + UNIT('V')))
parts.append(plain('取C点电势为0，可得A点电势为10 V，B点电势为30 V。'))
parts.append(eq(WW('AC') + mop('=') + WW('AB') + mop('+') + WW('BC') + mop('=') + mop('−') + SCI('1.6', '-18') + UNIT('J')))
parts.append(plain('静电力做负功，电子的电势能增加，增加量等于静电力做功的绝对值。换成等量正电荷后，电势差不变而电荷量符号改变，静电力做功变为正值。'))

# ---------- 素养提升 ----------
parts.append(heading('三、素养提升——多组数据验证电势差与场强', page_break=True))
parts.append(label('【题目】'))
parts.append(plain('某实验小组在同一匀强电场中选取不同试探电荷，分别测量A、B、C三点间静电力做功。A、B、C位于同一条电场线上，AB=0.20 m，BC=0.30 m，实验数据如下表。'))
parts.append(table([
    [hei('移动过程', 22), hei('试探电荷量', 22), hei('静电力做功', 22), hei('电势差', 22)],
    [roman('A→B', 22), sci_text('+2.0', '-9', 'C', 22), sci_text('+1.6', '-7', 'J', 22), cn('待求', 22)],
    [roman('B→C', 22), sci_text('−4.0', '-9', 'C', 22), sci_text('−4.8', '-7', 'J', 22), cn('待求', 22)],
    [roman('A→C', 22), sci_text('+1.0', '-9', 'C', 22), sci_text('+2.0', '-7', 'J', 22), cn('待求', 22)],
], [1900, 2600, 2600, 1600]))
parts.append(plain('（1）完成表格中的电势差，并验证电势差的可加性。'))
parts.append(plain('（2）取C点电势为0，求A、B两点的电势。'))
parts.append(plain('（3）分别利用AB段和BC段数据求电场强度，判断实验数据是否符合匀强电场模型。'))
parts.append(plain('（4）判断电场方向，并以A点为位置原点、沿A到C方向建立坐标轴，画出电势随位置变化的图像。'))
parts.append(body(cn('（5）预测电荷量为 '), sci_text('3.0', '-9', 'C'), cn(' 的正点电荷从A点移动到C点时静电力做的功。'))
parts.append(label('【答案】'))
parts.append(body(cn('（1）'), inline_sub('U', 'AB'), roman('=80 V，'), inline_sub('U', 'BC'), roman('=120 V，'), inline_sub('U', 'AC'), roman('=200 V，且200 V=80 V+120 V。'))
parts.append(body(cn('（2）'), inline_sub('φ', 'A'), roman('=200 V，'), inline_sub('φ', 'B'), roman('=120 V。'))
parts.append(body(cn('（3）两段均得 '), phys('E'), roman('=400 V/m'), cn('，数据符合匀强电场模型。'))
parts.append(body(cn('（4）电场方向由A指向C，图像为从（0，200 V）到（0.50 m，0）的下降直线。 （5）'), sci_text('6.0', '-7', 'J'), cn('。'))
parts.append(label('【解析】'))
parts.append(plain('根据静电力做功与电势差的关系，逐行计算可得'))
parts.append(eq(U('AB') + mop('=') + FRAC(SCI('1.6', '-7'), SCI('2.0', '-9')) + UNIT('V') + mop('=') + mn('80') + UNIT('V')))
parts.append(eq(U('BC') + mop('=') + FRAC(mop('−') + SCI('4.8', '-7'), mop('−') + SCI('4.0', '-9')) + UNIT('V') + mop('=') + mn('120') + UNIT('V')))
parts.append(eq(U('AC') + mop('=') + FRAC(SCI('2.0', '-7'), SCI('1.0', '-9')) + UNIT('V') + mop('=') + mn('200') + UNIT('V')))
parts.append(plain('三段数据满足电势差的可加性。取C点电势为0，可得B点电势为120 V，A点电势为200 V。'))
parts.append(eq(mv('E') + mop('=') + FRAC(U('AB'), mn('0.20') + UNIT('m')) + mop('=') + mn('400') + UNIT('V/m')))
parts.append(eq(mv('E') + mop('=') + FRAC(U('BC'), mn('0.30') + UNIT('m')) + mop('=') + mn('400') + UNIT('V/m')))
parts.append(plain('两段计算得到的场强相同，说明数据与匀强电场模型一致。沿电场方向电势降低，因此电场方向由A指向C。'))
parts.append(para(mono('φ/V'), align='center', after=0, line=260))
parts.append(para(mono('200 ● A'), align='center', after=0, line=260))
parts.append(para(mono('       ╲'), align='center', after=0, line=260))
parts.append(para(mono('120      ● B'), align='center', after=0, line=260))
parts.append(para(mono('             ╲'), align='center', after=0, line=260))
parts.append(para(mono('  0               ● C──── x/m'), align='center', after=45, line=260))
parts.append(eq(WW('AC') + mop('=') + mparen(SCI('3.0', '-9') + mop('×') + mn('200')) + UNIT('J') + mop('=') + SCI('6.0', '-7') + UNIT('J')))

# ---------- 高考真题 ----------
parts.append(heading('四、高考真题——等势线与静电力做功', page_break=True))
parts.append(label('【真题来源】'))
parts.append(plain('2015年普通高等学校招生全国统一考试新课标全国卷Ⅰ理科综合物理第15题。'))
parts.append(label('【题目】'))
parts.append(plain('如图，直线a、b和c、d是处于匀强电场中的两组平行线，M、N、P、Q是它们的交点。'))
parts.append(para(mono('                 a'), align='center', after=0, line=260))
parts.append(para(mono('        M────────────N'), align='center', after=0, line=260))
parts.append(para(mono('       ╱              ╱ c'), align='center', after=0, line=260))
parts.append(para(mono('    d ╱              ╱'), align='center', after=0, line=260))
parts.append(para(mono('     Q──────────────P'), align='center', after=0, line=260))
parts.append(para(mono('                 b'), align='center', after=40, line=260))
parts.append(plain('一电子由M点分别运动到N点和P点的过程中，静电力所做的负功相等，则（　　）'))
parts.append(body(cn('A．直线a位于某一等势面内，'), inline_sub('φ', 'M'), roman('>'), inline_sub('φ', 'Q')))
parts.append(body(cn('B．直线c位于某一等势面内，'), inline_sub('φ', 'M'), roman('>'), inline_sub('φ', 'N')))
parts.append(plain('C．若电子由M点运动到Q点，静电力做正功'))
parts.append(plain('D．若电子由P点运动到Q点，静电力做负功'))
parts.append(label('【答案】'))
parts.append(plain('B。'))
parts.append(label('【解析】'))
parts.append(plain('电子带负电。由题意可知，电子从M点到N点和从M点到P点时，静电力做功相等且均为负。'))
parts.append(eq(WW('MN') + mop('=') + WW('MP') + mop('<') + mn('0')))
parts.append(plain('根据静电力做功与电势差的关系可得'))
parts.append(eq(U('MN') + mop('=') + FRAC(WW('MN'), Q()) + mop('，') + U('MP') + mop('=') + FRAC(WW('MP'), Q())))
parts.append(plain('两段功相等且电子电荷量相同，所以两段电势差相等；又因为功和电荷量均为负，所以两段电势差均为正。'))
parts.append(eq(U('MN') + mop('=') + U('MP') + mop('>') + mn('0')))
parts.append(plain('由电势差定义可得N、P两点电势相等，且M点电势高于N点电势。'))
parts.append(eq(PHI('N') + mop('=') + PHI('P') + mop('，') + PHI('M') + mop('>') + PHI('N')))
parts.append(plain('N、P同在直线c上，因此直线c为等势线，B项正确。M、N两点电势不同，所以直线a不是等势线，A项错误。'))
parts.append(plain('MNPQ为平行四边形，MQ与NP平行。NP为等势方向，因此M、Q两点电势相等，电子由M点运动到Q点时静电力不做功，C项错误。'))
parts.append(plain('P点电势低于Q点电势，电子由P点运动到Q点时，电势差与电荷量均为负，静电力做正功，D项错误。'))
parts.append(label('【与课本例题的关联】'))
parts.append(plain('课本例题由静电力做功定量求电势差和电势高低；本题仍以静电力做功为入口，进一步判断电势差符号、等势线和不同路径上的做功。知识、方法和原理均与课本例题高度贴合。'))

'''

source = source[:start] + replacement + source[end:]

check = source.index('# ---------- 自检 ----------')
source = source[:check] + r'''# ---------- 自检 ----------
assert abs(1.6e-7 / 2.0e-9 - 80.0) < 1e-12
assert abs(-4.0e-7 / 2.0e-9 + 200.0) < 1e-12
assert abs(1.5e-9 * (-120.0) + 1.8e-7) < 1e-20
assert abs(1.8e-7 / 3.0e-9 - 60.0) < 1e-12
assert abs(-4.5e-7 / 3.0e-9 + 150.0) < 1e-12
assert abs(2.0e-9 * (-90.0) + 1.8e-7) < 1e-20
assert abs(3.2e-18 / (-1.60e-19) + 20.0) < 1e-12
assert abs((-4.8e-18) / (-1.60e-19) - 30.0) < 1e-12
assert abs((-1.6e-18) - (3.2e-18 - 4.8e-18)) < 1e-30
assert abs(1.6e-7 / 2.0e-9 - 80.0) < 1e-12
assert abs((-4.8e-7) / (-4.0e-9) - 120.0) < 1e-12
assert abs(2.0e-7 / 1.0e-9 - 200.0) < 1e-12
assert abs(80.0 / 0.20 - 400.0) < 1e-12
assert abs(120.0 / 0.30 - 400.0) < 1e-12
assert abs(3.0e-9 * 200.0 - 6.0e-7) < 1e-20

assert OUT.exists() and OUT.stat().st_size > 7000
with ZipFile(OUT) as z:
    xml = z.read('word/document.xml').decode('utf-8')
    ET.fromstring(xml)
    ET.fromstring(z.read('word/styles.xml'))
    ET.fromstring(z.read('word/settings.xml'))
    forbidden = '⁰¹²³⁴⁵⁶⁷⁸⁹⁻₀₁₂₃₄₅₆₇₈₉ₑₚ'
    assert not any(ch in xml for ch in forbidden)
    for token in ['UAB', 'UBC', 'UAC', 'WAB', 'WBC', 'WAC', 'φA', 'φB', 'φC', 'φM', 'φN', 'φP', 'φQ']:
        assert token not in xml
    assert '<w:vertAlign w:val="superscript"/>' in xml
    assert '<w:vertAlign w:val="subscript"/>' in xml
    assert '<m:sSub>' in xml and '<m:sSup>' in xml and '<m:f>' in xml
    assert '<m:oMathPara>' in xml and '<w:tbl>' in xml
    assert xml.count('【题目】') == 4
    assert xml.count('【答案】') == 4
    assert xml.count('【解析】') == 4
    assert '2015年普通高等学校招生全国统一考试新课标全国卷Ⅰ理科综合物理第15题' in xml
    assert 'B。' in xml
    for text in ['CA∶AB=3∶2', '电子束能量标定', 'E=400 V/m', '直线c为等势线']:
        assert text in xml
    for bad in ['TODO', '待补充', 'XXXX', '�']:
        assert bad not in xml

print(OUT)
'''

exec(compile(source, 'build_potential_difference_all_refined_generated.py', 'exec'), {'__name__': '__main__'})
