from pathlib import Path

BASE = Path('tools/build_potential_difference_regenerated_v2.py')
source = BASE.read_text(encoding='utf-8')

source = source.replace(
    "一题三变加高考真题_第十章_电势差_重新生成自检版.docx",
    "一题三变加高考真题_第十章_电势差_高考题替换版.docx",
)

start_marker = '# ---------- 高考真题 ----------'
end_marker = '# ---------- 打包 ----------'
start = source.index(start_marker)
end = source.index(end_marker)

replacement = r'''# ---------- 高考真题 ----------
parts.append(heading('四、高考真题——等势线与静电力做功', page_break=True))
parts.append(label('【真题来源】'))
parts.append(plain('2015年普通高等学校招生全国统一考试新课标全国卷Ⅰ理科综合物理第15题。'))
parts.append(label('【题目】'))
parts.append(plain('如图，直线a、b和c、d是处于匀强电场中的两组平行线，M、N、P、Q是它们的交点，四点处的电势分别用带相应下标的φ表示。'))
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
parts.append(plain('设电子的电荷量为q，q<0。由题意可知，电子从M点到N点、从M点到P点时，静电力所做的功相等且均为负值。'))
parts.append(eq(WW('MN') + mop('=') + WW('MP') + mop('<') + mn('0')))
parts.append(plain('根据静电力做功与电势差的关系可得'))
parts.append(eq(U('MN') + mop('=') + FRAC(WW('MN'), Q()) + mop('，') + U('MP') + mop('=') + FRAC(WW('MP'), Q())))
parts.append(plain('由于两段功相等，电子电荷量相同，所以两段电势差相等；又因为功为负、电荷量为负，所以两段电势差均为正。'))
parts.append(eq(U('MN') + mop('=') + U('MP') + mop('>') + mn('0')))
parts.append(plain('根据电势差的定义可得'))
parts.append(eq(PHI('M') + mop('−') + PHI('N') + mop('=') + PHI('M') + mop('−') + PHI('P')))
parts.append(eq(PHI('N') + mop('=') + PHI('P') + mop('，') + PHI('M') + mop('>') + PHI('N')))
parts.append(plain('N、P两点电势相等，且N、P同在直线c上。在匀强电场中，过两个等势点的直线若与电场方向垂直，则整条直线位于同一等势面内，因此直线c位于某一等势面内，B项正确。'))
parts.append(plain('M、N两点电势不相等，所以直线a不可能位于同一等势面内，A项错误。'))
parts.append(plain('MNPQ为平行四边形，MQ与NP平行。既然NP方向为等势方向，则MQ方向也为等势方向，所以M、Q两点电势相等，电子由M点运动到Q点时静电力不做功，C项错误。'))
parts.append(plain('P点电势低于Q点电势，电子由P点运动到Q点时，电势差为负，电子电荷量也为负，因此静电力做正功，D项错误。'))
parts.append(label('【与课本例题的关联】'))
parts.append(plain('课本例题利用已知静电力做功定量求电势差，并据此比较A、B、C三点电势；本题同样以静电力做功为入口，但改为定性判断电势差符号、等势线和做功正负。两题的核心方法均为先用W=qU建立做功与电势差的联系，再由电势差判断电势高低。'))

'''

source = source[:start] + replacement + source[end:]
source = source.replace(
    "2020年普通高等学校招生全国统一考试江苏卷物理第9题",
    "2015年普通高等学校招生全国统一考试新课标全国卷Ⅰ理科综合物理第15题",
)
source = source.replace("assert 'A、B。' in xml", "assert 'B。' in xml")
source = source.replace("assert xml.count('【题目】') == 4", "assert xml.count('【题目】') == 4")
source = source.replace(
    "for value in ['−6.0', '200 V/m', '3∶2']:",
    "for value in ['−6.0', '200 V/m', '3∶2', '直线c位于某一等势面内']:",
)

exec(compile(source, 'build_potential_difference_gaokao_2015_generated.py', 'exec'), {'__name__': '__main__'})
