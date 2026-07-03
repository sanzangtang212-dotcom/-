from pathlib import Path

BASE = Path('tools/build_linear_accelerator_symbolic_docx.py')
source = BASE.read_text(encoding='utf-8')

source = source.replace(
    "downloads/多级直线加速器_一题三变_字母推导版.docx",
    "downloads/多级直线加速器_一题三变加高考真题_字母推导版.docx",
)
source = source.replace(
    "parts.append(paragraph(hei('一题三变·字母推导版', 30), align='center', after=190))",
    "parts.append(paragraph(hei('一题三变＋高考真题·字母推导版', 30), align='center', after=190))",
)

marker = '# Reminders\n'
insert = r'''# Gaokao question
parts.append(heading('四、高考真题——回旋加速器的最大速率', page_break=True))
parts.append(label('【真题来源】'))
parts.append(body('2023年广东省普通高中学业水平选择性考试物理第5题。', indent=False))
parts.append(label('【题目】'))
parts.append(paragraph(
    cn('某小型医用回旋加速器，最大回旋半径为0.50 m，磁感应强度大小为1.12 T，质子加速后获得的最大动能为')
    + roman('1.5×10') + roman('7', sup=True) + roman(' eV')
    + cn('。根据给出的数据，可计算质子经该回旋加速器加速后的最大速率约为（忽略相对论效应，')
    + roman('1 eV=1.6×10') + roman('−19', sup=True) + roman(' J') + cn('）（　　）'),
    first_line=480, after=90
))
parts.append(paragraph(cn('A．') + roman('3.6×10') + roman('6', sup=True) + roman(' m/s'), after=45))
parts.append(paragraph(cn('B．') + roman('1.2×10') + roman('7', sup=True) + roman(' m/s'), after=45))
parts.append(paragraph(cn('C．') + roman('5.4×10') + roman('7', sup=True) + roman(' m/s'), after=45))
parts.append(paragraph(cn('D．') + roman('2.4×10') + roman('8', sup=True) + roman(' m/s'), after=70))
parts.append(label('【答案】'))
parts.append(body('C。'))
parts.append(label('【解析】'))
parts.append(body('质子运动到最大回旋半径时，洛伦兹力提供向心力。根据牛顿第二定律可得'))
parts.append(eq(mv('q') + V() + mv('B') + mop('=') + FRAC(mv('m') + SQ(V()), mv('R'))))
parts.append(body('整理可得质子的动量关系'))
parts.append(eq(mv('m') + V() + mop('=') + mv('q') + mv('B') + mv('R')))
parts.append(body('根据动能表达式，将上式代入可得'))
parts.append(eq(msub(mv('E'), mu('k,max')) + mop('=') + FRAC(mn('1'), mn('2')) + mv('m') + SQ(V()) + mop('=') + FRAC(mn('1'), mn('2')) + mv('q') + mv('B') + mv('R') + V()))
parts.append(body('因此，质子的最大速率为'))
parts.append(eq(V() + mop('=') + FRAC(mn('2') + msub(mv('E'), mu('k,max')), mv('q') + mv('B') + mv('R'))))
parts.append(body('质子的电荷量为元电荷。把以电子伏特为单位的动能换算为焦耳时，分子中出现的元电荷与分母中的质子电荷量相约，故'))
parts.append(eq(
    V() + mop('=')
    + FRAC(mn('2') + mop('×') + mn('1.5') + mop('×') + msup(mn('10'), mn('7')),
           mn('1.12') + mop('×') + mn('0.50'))
    + mu(' m/s')
    + mop('≈') + mn('5.4') + mop('×') + msup(mn('10'), mn('7')) + mu(' m/s')
))
parts.append(body('故选C。'))
parts.append(label('【与原例题的关联】'))
parts.append(body('原例题研究直线加速器中粒子每次通过间隙获得相同能量，并利用同步条件设计圆筒长度；本题研究回旋加速器中粒子反复通过交变电场获得能量。两题的共同核心是“交变电场重复加速＋能量关系”，本题进一步把粒子的最大动能与磁场、最大轨道半径联系起来。'))

'''
source = source.replace(marker, insert + marker)

source = source.replace("assert xml.count('【题目】') == 3", "assert xml.count('【题目】') == 4")
source = source.replace("assert xml.count('【答案】') == 3", "assert xml.count('【答案】') == 4")
source = source.replace("assert xml.count('【解析】') == 3", "assert xml.count('【解析】') == 4")
source = source.replace(
    "['同模型规律巩固', '粒子具有初速度', '同一加速器加速不同粒子', '圆筒长度的平方构成等差数列']",
    "['同模型规律巩固', '粒子具有初速度', '同一加速器加速不同粒子', '圆筒长度的平方构成等差数列', '2023年广东省普通高中学业水平选择性考试物理第5题', '回旋加速器的最大速率']",
)

exec(compile(source, 'build_linear_accelerator_symbolic_with_gaokao_generated.py', 'exec'), {'__name__': '__main__'})
