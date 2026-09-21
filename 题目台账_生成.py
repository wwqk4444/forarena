# -*- coding: utf-8 -*-
"""生成"已用题目台账"：题目名称/来源/代号/考点/得分率/难度系数/使用时间。
读取 parsed.json（由 成绩统计_解析.py 生成），输出 题目台账.xlsx 与 题目台账.md。

来源与考点维护说明：
  META[(场次, 题名)] = (来源, 考点, 状态, 备注)
  状态: '确认' = 已核实出处; '待核' = 高度疑似、待老师确认; '原创' = 校内原创/待确认
  考点为 '待补充' 时表中以黄色高亮，提示老师补充。
"""
import json, os, statistics
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))

LABEL = {'20260620 tongbu.cdf': '0620同步', '0701.cdf': '0701', '260702S.cdf': '0702S',
         '0703S.cdf': '0703S', '260706S.cdf': '0706S', '260707S.cdf': '0707S',
         '260708S.cdf': '0708S', '0710S.cdf': '0710S', '0713S.cdf': '0713S',
         'contest.cdf': '0715赛', '0714S.cdf': '0714S', '0716S.cdf': '0716S',
         '0717S.cdf': '0717S', '20260717入门级.cdf': '0717入门', '0822test.cdf': '0822',
         '0823test.cdf': '0823', '0824Stest.cdf': '0824S', '0825test.cdf': '0825',
         '0826S.cdf': '0826S', '0826testrumen.cdf': '0826入门', '260905test.cdf': '0905',
         '0912S.cdf': '0912'}

# ---------------- 来源 / 考点 信息表 ----------------
# 状态: 确认=已核实  待核=高度疑似待确认  原创=校内原创待确认
C, W, O = '确认', '待核', '原创'
META = {
    # ---- 0620同步（入门组）----
    ('0620同步', 'calculator'): ('校内原创', '待补充', O, ''),
    ('0620同步', 'city'):       ('校内原创', '待补充', O, ''),
    ('0620同步', 'magic'):      ('校内原创', '待补充', O, ''),
    ('0620同步', 'number'):     ('校内原创', '待补充', O, ''),
    ('0620同步', 'power'):      ('校内原创', '待补充', O, ''),
    # ---- 0701 ----
    ('0701', 'cryptarithm'): ('待确认', '待补充', W, ''),
    ('0701', 'debug'):       ('待确认', '待补充', W, ''),
    ('0701', 'bond'):        ('COCI 2006/2007 Contest 1', '状压DP（任务分配）、子集枚举+记忆化', C, 'StackOverflow 题解可查证'),
    ('0701', 'winter'):      ('待确认', '待补充', W, ''),
    # ---- 0702S（拼盘）----
    ('0702S', 'abc'):       ('疑似COCI（轮次待核）', '模拟（推测）', W, ''),
    ('0702S', 'kolone'):    ('疑似COCI（轮次待核）', '模拟（推测）', W, ''),
    ('0702S', 'r2'):        ('疑似COCI 2016/17 R2', '数学/解一元方程（推测）', W, ''),
    ('0702S', 'sjecista'):  ('疑似COCI（轮次待核）', '组合计数 C(n,4)（推测）', W, ''),
    ('0702S', 'stol'):      ('疑似COCI（轮次待核）', '枚举/几何（推测）', W, ''),
    ('0702S', 'bookshelf'): ('疑似USACO（待核）', 'DP/枚举（推测）', W, ''),
    # ---- 0703S ----
    ('0703S', 'trade'):   ('疑似COCI 2024/2025（待核）', '待补充', W, ''),
    ('0703S', 'history'): ('疑似COCI 2024/2025（待核）', '待补充', W, ''),
    ('0703S', 'oasis'):   ('疑似COCI 2024/2025（待核）', '待补充', W, ''),
    ('0703S', 'trains'):  ('疑似COCI 2024/2025（待核）', '待补充', W, '全场0分，难度过高/可能超纲'),
    # ---- 0706S ----
    ('0706S', 'ari'): ('疑似COCI 2022/2023（待核）', '待补充', W, ''),
    ('0706S', 'bod'): ('疑似COCI 2022/2023（待核）', '待补充', W, ''),
    ('0706S', 'mp3'): ('疑似COCI 2022/2023（待核）', '待补充', W, ''),
    ('0706S', 'tow'): ('疑似COCI 2022/2023（待核）', '待补充', W, ''),
    # ---- 0707S ----
    ('0707S', 'class'): ('待确认', '待补充', W, ''),
    ('0707S', 'port'):  ('待确认', '待补充', W, ''),
    ('0707S', 'road'):  ('待确认', '待补充', W, ''),
    ('0707S', 'tower'): ('待确认', '待补充', W, ''),
    # ---- 0708S ----
    ('0708S', 'bic'): ('疑似COCI 2023/2024 R1（待核）', '待补充', W, ''),
    ('0708S', 'dva'): ('疑似COCI 2023/2024 R1（待核）', '待补充', W, ''),
    ('0708S', 'pat'): ('疑似COCI 2023/2024 R1（待核）', '待补充', W, ''),
    ('0708S', 'zbr'): ('疑似COCI 2023/2024 R1（待核）', '待补充', W, ''),
    # ---- 0710S ----
    ('0710S', 'holding'): ('COCI 2019/2020 Round 4', '区间DP+前缀和（k次交换最小化差）', C, 'oj.uz COCI20_holding 可查证'),
    ('0710S', 'planine'): ('COCI 2020/2021 Round 5', '排序+贪心/几何（山顶可见性）', C, 'VNOJ 可查证'),
    ('0710S', 'sjekira'): ('COCI 2020/2021 Round 2', '树形贪心/DFS（删边分割树）', C, 'VNOJ coci2021_r2_sjekira'),
    ('0710S', 'patkice'): ('COCI 2020/2021 Round 1', '网格模拟/BFS', C, 'VNOJ coci2021_r1_patkice'),
    # ---- 0713S（拼盘）----
    ('0713S', 'sjeckanje'): ('疑似COCI 2021/2022 R4', '线段树+DP、区间差分（推测）', W, ''),
    ('0713S', 'stogovi'):   ('疑似COCI 2016/17 R5', '树上倍增/LCA、离线处理（推测）', W, ''),
    ('0713S', 'tajna'):     ('COCI 2016/2017 Round 2', '模拟/矩阵重排', C, '经典COCI签到题'),
    ('0713S', 'utrka'):     ('疑似COCI 2016/17 R4', '哈希/map统计（推测）', W, ''),
    # ---- 0714S ----
    ('0714S', 'alarms'):         ('疑似USACO 2025-26赛季（待核）', '待补充', W, ''),
    ('0714S', 'backrooms'):      ('疑似USACO 2025-26赛季（待核）', '待补充', W, ''),
    ('0714S', 'catchingapples'): ('疑似USACO 2025-26赛季（待核）', '待补充', W, '与USACO 2022 Open Gold "Apple Catching"名近，是否同题待核'),
    ('0714S', 'namechange'):     ('疑似USACO 2025-26赛季（待核）', '待补充', W, ''),
    # ---- 0715赛（contest.cdf）----
    ('0715赛', 'queue'):   ('待确认', '待补充', W, ''),
    ('0715赛', 'compare'): ('待确认', '待补充', W, ''),
    ('0715赛', 'xor'):     ('待确认', '位运算/前缀异或（按题名推测）', W, ''),
    ('0715赛', 'yugo'):    ('疑似COCI（轮次待核）', '模拟/数学（推测）', W, ''),
    # ---- 0716S ----
    ('0716S', 'exhibition'): ('待确认', '待补充', W, ''),
    ('0716S', 'growing'):    ('待确认', '待补充', W, '得分率仅7%，难度超预期'),
    ('0716S', 'lamps'):      ('待确认', '待补充', W, ''),
    ('0716S', 'stove'):      ('疑似AtCoder ABC', '排序+贪心/区间合并（推测）', W, ''),
    # ---- 0717S ----
    ('0717S', 'fruits'):     ('待确认', '待补充', W, '全场仅1.5%，几乎无人得分'),
    ('0717S', 'gymbadges'):  ('待确认', '待补充', W, ''),
    ('0717S', 'towers'):     ('待确认', '待补充', W, ''),
    ('0717S', 'votingcity'): ('待确认', '待补充', W, ''),
    # ---- 0717入门 ----
    ('0717入门', 'add'):      ('校内原创/经典混编', '待补充', O, ''),
    ('0717入门', 'draw'):     ('校内原创/经典混编', '待补充', O, ''),
    ('0717入门', 'exchange'): ('校内原创/经典混编', '待补充', O, ''),
    ('0717入门', 'phone'):    ('校内原创/经典混编', '待补充', O, ''),
    ('0717入门', 'ride'):     ('疑似USACO Training "Your Ride Is Here"', '字符串取模/模拟（推测）', W, 'USACO训练官网经典首题'),
    ('0717入门', 'print'):    ('校内原创/经典混编', '待补充', O, ''),
    ('0717入门', 'sum'):      ('校内原创/经典混编', '待补充', O, ''),
    ('0717入门', 'time'):     ('校内原创/经典混编', '待补充', O, ''),
    ('0717入门', 'zero'):     ('校内原创/经典混编', '待补充', O, ''),
    ('0717入门', 'judge'):    ('校内原创/经典混编', '待补充', O, ''),
    # ---- 0822 ----
    ('0822', 'bouquet'): ('待确认', '待补充', W, ''),
    ('0822', 'garaza'):  ('疑似COCI 2023/2024（待核）', '待补充', W, ''),
    ('0822', 'notdiv'):  ('疑似AtCoder ABC170 D "Not Divisible"', '筛法/调和级数枚举（推测）', W, ''),
    ('0822', 'skip'):    ('待确认', '待补充', W, ''),
    ('0822', 'gcdsum'):  ('疑似Codeforces（待核）', 'gcd/数位和/贪心（推测）', W, ''),
    # ---- 0823 ----
    ('0823', 'colors'):  ('待确认', '待补充', W, ''),
    ('0823', 'connect'): ('待确认', '图连通性 BFS/并查集（按题名推测）', W, ''),
    ('0823', 'oneway'):  ('待确认', '有向图/拓扑或SCC（按题名推测）', W, ''),
    ('0823', 'walk500'): ('待确认', '待补充', W, ''),
    # ---- 0824S ----
    ('0824S', 'ant'):     ('疑似USACO "Ant Counting"（场次待核）', 'DP/多重集组合计数（推测）', W, ''),
    ('0824S', 'nochange'): ('USACO 2015-12 Gold P3 "No Change"', '状压DP+前缀和（硬币支付）', C, ''),
    ('0824S', 'piano'):   ('疑似AtCoder ABC（待核）', '构造/分类讨论（推测）', W, ''),
    ('0824S', 'subperm'): ('待确认', '排列/计数（按题名推测）', W, ''),
    # ---- 0825 ----
    ('0825', 'cake'):    ('疑似USACO 2024-25 Dec Silver "Cake Game"', '博弈/枚举（推测）', W, ''),
    ('0825', 'closing'): ('USACO 2016 US Open Silver P3 "Closing the Farm"', '并查集（时间倒流/离线加点）', C, ''),
    ('0825', 'cowland'): ('USACO 2019-02 Platinum P2 "Cow Land"', '树链剖分/线段树维护区间异或', C, ''),
    ('0825', 'reveg'):   ('USACO 2019-02 Silver P3 "Revegetate"', '贪心染色/并查集', C, ''),
    # ---- 0826S ----
    ('0826S', 'palindrome'): ('待确认', '回文判定/模拟（按题名推测）', W, ''),
    ('0826S', 'lostarray'):  ('Codeforces 1043B "Lost Array"', '思维构造/前缀和验证', C, 'CF Round #519 Div2 B'),
    ('0826S', 'charges'):    ('待确认', '待补充', W, ''),
    ('0826S', 'feast'):      ('USACO 2016-12 Gold P1 "Feast"', 'DP（吃/喝水两阶段转移）', C, ''),
    # ---- 0826入门 ----
    ('0826入门', 'T1'): ('校内原创', '待补充', O, ''),
    ('0826入门', 'T2'): ('校内原创', '待补充', O, ''),
    ('0826入门', 'T3'): ('校内原创', '待补充', O, ''),
    ('0826入门', 'T4'): ('校内原创', '待补充', O, ''),
    ('0826入门', 'T5'): ('校内原创', '待补充', O, ''),
    ('0826入门', 'T6'): ('校内原创', '待补充', O, ''),
    ('0826入门', 'T7'): ('校内原创', '待补充', O, ''),
    # ---- 0905 ----
    ('0905', 'haircut'):  ('USACO 2019-12 Silver P3 "Haircut"', '树状数组/逆序对', C, ''),
    ('0905', 'hilo'):     ('USACO 2021-12 Gold "HILO"', '递推/前后缀分析', C, ''),
    ('0905', 'cave'):     ('疑似USACO（待核）', '待补充', W, ''),
    ('0905', 'sleeping'): ('待确认', '待补充', W, '得分率仅3%'),
    # ---- 0912 ----
    ('0912', 'vacation'):  ('待确认（疑似JOI/经典DP）', '待补充', W, ''),
    ('0912', 'jjooii'):    ('JOI（AOJ 0571 或 JOI 2020 Final "JJOOII 2"，版本待核）', '前缀和/双指针、字符串扫描', C, ''),
    ('0912', 'numbers'):   ('待确认', '待补充', W, ''),
    ('0912', 'exhausted'): ('待确认', '待补充', W, ''),
}

# 未启用归档的题目（无学生成绩，仅供参考，不计入台账主表）
UNUSED = [
    ('0826S加赛', '—', 'pilot', '待确认', '待补充', W, ''),
    ('0826S加赛', '—', 'lasers', 'USACO 2016-01 Gold "Lasers and Mirrors"', 'BFS/状态图建模', C, ''),
    ('0826S加赛', '—', 'riggedroads', '疑似USACO 2023-24 Gold（待核）', '待补充', W, ''),
    ('20260715 senior', '—', 'apple/bus/game/lock/road/struct/tree/uqe（8题）', '疑似校内原创组别赛', '待补充', O, '归档无选手成绩'),
    ('20260717 senior', '—', 'airport/bracket/candy/fruit/network/palin/sort/traffic（8题）', '疑似校内原创组别赛', '待补充', O, '归档无选手成绩'),
    ('2025-08 三场', '—', '题名不可辨（1~7）', '—', '—', O, '仅含标程std，未投入使用'),
]

# ---------------- 难度档位 ----------------
def band(rate):
    if rate >= 0.85: return '★1 热身'
    if rate >= 0.65: return '★2 基础'
    if rate >= 0.45: return '★3 中档'
    if rate >= 0.25: return '★4 较难'
    if rate >= 0.10: return '★5 难'
    return '★6 压轴'

HDR_FILL = PatternFill('solid', fgColor='305496')
SUB_FILL = PatternFill('solid', fgColor='8EA9DB')
TITLE_FILL = PatternFill('solid', fgColor='203864')
G9 = PatternFill('solid', fgColor='C6EFCE')
G7 = PatternFill('solid', fgColor='E2EFDA')
G4 = PatternFill('solid', fgColor='FFEB9C')
G1 = PatternFill('solid', fgColor='FCE4D6')
G0 = PatternFill('solid', fgColor='D9D9D9')
TODO = PatternFill('solid', fgColor='FFF2CC')
WHITE_BOLD = Font(bold=True, color='FFFFFF')
BOLD = Font(bold=True)
SMALL = Font(size=9, color='595959')
CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def rate_fill(rate):
    if rate >= 0.9: return G9
    if rate >= 0.7: return G7
    if rate >= 0.4: return G4
    if rate > 0: return G1
    return G0


def fmt(x):
    if isinstance(x, float) and abs(x - round(x)) < 1e-6:
        return int(round(x))
    return x


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def main():
    with open(os.path.join(HERE, 'parsed.json'), encoding='utf-8') as f:
        data = json.load(f)
    rows = []
    for cinfo in data['contests']:
        if cinfo['file'] not in LABEL or not cinfo['students']:
            continue
        n = len(cinfo['students'])
        for qi, (tn, fs) in enumerate(cinfo['tasks']):
            vals = [s['ps'][qi] for s in cinfo['students']]
            avg = sum(vals) / n
            rate = avg / fs if fs else 0
            meta = META.get((LABEL[cinfo['file']], str(tn)), ('待确认', '待补充', W, ''))
            rows.append({
                'label': LABEL[cinfo['file']], 'date': cinfo.get('judge_date'),
                'group': cinfo['group'], 'tidx': qi + 1, 'name': str(tn), 'full': fs,
                'n': n, 'avg': round(avg, 1), 'rate': rate,
                'fulln': sum(1 for v in vals if v == fs and fs > 0),
                'zeron': sum(1 for v in vals if v == 0),
                'src': meta[0], 'tags': meta[1], 'status': meta[2], 'note': meta[3],
            })
    rows.sort(key=lambda r: (r['date'] or '', r['label'], r['tidx']))

    wb = Workbook()
    # ============ 说明 ============
    ws = wb.active
    ws.title = '说明'
    lines = [
        ('已用题目台账 —— cdf归档 22 场比赛 · 101 道题', TITLE_FILL, WHITE_BOLD, 14),
        ('', None, None, None),
        ('字段说明：', HDR_FILL, WHITE_BOLD, 11),
        ('· 代号：题目在评测归档中的内部名（即评测文件夹名）；场次题号 = "场次-T数字"。', None, None, None),
        ('· 得分率：本场参赛学生的平均得分 ÷ 满分。难度系数 = 1 − 得分率（越大越难）。', None, None, None),
        ('· 难度档：★1热身 ≥85% ＞ ★2基础 ≥65% ＞ ★3中档 ≥45% ＞ ★4较难 ≥25% ＞ ★5难 ≥10% ＞ ★6压轴', None, None, None),
        ('  （该难度反映的是本届学生对本题的实际掌握程度，非客观难度）', None, SMALL, None),
        ('· 确认状态：', None, None, None),
        ('  - 确认：已通过 oj.uz / VNOJ / hsin.hr / usaco.org 等官方源核实出处；', None, None, None),
        ('  - 待核：按题名高度疑似某来源，但未逐题核实，黄色底纹提示；', None, None, None),
        ('  - 原创：判断为校内原创/组别内部题，考点需老师补充。', None, None, None),
        ('· 考点列"待补充"（黄色）处请老师根据题面补填；多技巧组合用"+"或"、"分隔。', None, None, None),
        ('', None, None, None),
        ('工作表：', HDR_FILL, WHITE_BOLD, 11),
        ('台账总表：22 场比赛用过的全部 101 道题（按使用日期排序）。', None, None, None),
        ('未启用归档：已建比赛但无学生成绩的 6 个归档（0826S加赛、两场senior、2025年8月三场）。', None, None, None),
        ('来源统计：按题目来源汇总数量与平均得分率。', None, None, None),
        ('难度分布：按难度档汇总。', None, None, None),
        ('', None, None, None),
        ('维护方式：来源/考点集中在 题目台账_生成.py 的 META 表中，改后重跑即可更新 Excel/MD。', None, SMALL, None),
    ]
    for i, (txt, fill, font, size) in enumerate(lines, 1):
        cell = ws.cell(row=i, column=1, value=txt)
        if fill: cell.fill = fill
        if font: cell.font = font
        if size and font is None: cell.font = Font(size=size, bold=True)
    ws.column_dimensions['A'].width = 105

    # ============ 台账总表 ============
    ws = wb.create_sheet('台账总表')
    headers = ['序号', '使用日期', '场次', '组别', '题号', '代号', '题目名称', '来源（出处）',
               '考点/技巧', '确认状态', '满分', '参赛人数', '平均分', '得分率', '难度系数',
               '难度档', '满分人数', '零分人数', '备注']
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill = HDR_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[1].height = 24
    r = 2
    for i, rw in enumerate(rows, 1):
        code = f"{rw['label']}-T{rw['tidx']}"
        vals = [i, rw['date'], rw['label'], rw['group'], f"T{rw['tidx']}", rw['name'], rw['name'],
                rw['src'], rw['tags'], rw['status'], rw['full'], rw['n'], fmt(rw['avg']),
                round(rw['rate'], 3), round(1 - rw['rate'], 2), band(rw['rate']),
                rw['fulln'], rw['zeron'], rw['note']]
        for j, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = BORDER
            cell.alignment = LEFT if j in (8, 9, 19) else CENTER
            if j == 14:
                cell.number_format = '0.0%'
                cell.fill = rate_fill(rw['rate'])
            if j == 15: cell.number_format = '0.00'
            if j in (6, 7): cell.font = BOLD if j == 6 else Font()
            if j == 10 and v == '确认':
                cell.font = Font(color='006100')
        # 待补充高亮
        if rw['tags'] == '待补充':
            ws.cell(row=r, column=9).fill = TODO
        if '待核' in rw['status'] or rw['status'] == '原创':
            ws.cell(row=r, column=10).fill = TODO
        r += 1
    set_widths(ws, [5, 11, 9, 7, 5, 13, 15, 33, 30, 9, 6, 8, 7, 8, 8, 9, 8, 8, 26])
    ws.freeze_panes = 'H2'
    ws.auto_filter.ref = f'A1:S{r-1}'

    # ============ 未启用归档 ============
    ws = wb.create_sheet('未启用归档')
    headers = ['场次', '日期', '题目', '来源（出处）', '考点/技巧', '确认状态', '备注']
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill = HDR_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BORDER
    for i, row in enumerate(UNUSED, 2):
        for j, v in enumerate(row, 1):
            cell = ws.cell(row=i, column=j, value=v)
            cell.border = BORDER
            cell.alignment = LEFT if j in (3, 4, 5, 7) else CENTER
            if j == 6 and v != '确认': cell.fill = TODO
    set_widths(ws, [16, 8, 48, 34, 24, 9, 30])

    # ============ 来源统计 ============
    ws = wb.create_sheet('来源统计')
    def src_key(rw):
        s = rw['src']
        if s.startswith('COCI'): return 'COCI（已确认）'
        if s.startswith('USACO'): return 'USACO（已确认）'
        if s.startswith('Codeforces'): return 'Codeforces（已确认）'
        if s.startswith('AtCoder'): return 'AtCoder ABC（已确认）'
        if s.startswith('JOI'): return 'JOI（已确认）'
        if s.startswith('校内原创'): return '校内原创/混编'
        if '疑似COCI' in s: return '疑似COCI（待核）'
        if '疑似USACO' in s: return '疑似USACO（待核）'
        if '疑似AtCoder' in s or '疑似Codeforces' in s: return '疑似AtCoder/CF（待核）'
        if s == '待确认' or s.startswith('待确认'): return '待确认'
        return '其他待核'
    agg = {}
    for rw in rows:
        k = src_key(rw)
        a = agg.setdefault(k, {'n': 0, 'rate': 0.0, 'zero': 0})
        a['n'] += 1; a['rate'] += rw['rate']; a['zero'] += (1 if rw['rate'] < 0.1 else 0)
    headers = ['来源', '题目数', '平均得分率', '得分率<10%题数']
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill = HDR_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BORDER
    r = 2
    show = []
    for k in ['COCI（已确认）', 'USACO（已确认）', 'Codeforces（已确认）', 'AtCoder ABC（已确认）',
              'JOI（已确认）', '校内原创/混编', '疑似COCI（待核）', '疑似USACO（待核）',
              '疑似AtCoder/CF（待核）', '待确认', '其他待核']:
        if k in agg: show.append((k, agg[k]))
    for k, a in show:
        vals = [k, a['n'], round(a['rate'] / a['n'], 3), a['zero']]
        for j, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = BORDER; cell.alignment = CENTER if j > 1 else LEFT
            if j == 3:
                cell.number_format = '0.0%'; cell.fill = rate_fill(v)
        r += 1
    set_widths(ws, [24, 9, 11, 15])

    # ============ 难度分布 ============
    ws = wb.create_sheet('难度分布')
    headers = ['难度档', '题数', '占比', '平均得分率', '代表题目（得分率居中的示例）']
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill = HDR_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BORDER
    bands = ['★1 热身', '★2 基础', '★3 中档', '★4 较难', '★5 难', '★6 压轴']
    r = 2
    for b in bands:
        sel = [rw for rw in rows if band(rw['rate']) == b]
        if not sel: continue
        mid = sorted(sel, key=lambda x: x['rate'])[len(sel) // 2]
        vals = [b, len(sel), round(len(sel) / len(rows), 3),
                round(sum(x['rate'] for x in sel) / len(sel), 3),
                f"{mid['label']}-{mid['name']}（{mid['rate']:.0%}）"]
        for j, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = BORDER; cell.alignment = CENTER if j < 5 else LEFT
            if j in (3, 4): cell.number_format = '0.0%'
            if j == 4: cell.fill = rate_fill(v)
        r += 1
    set_widths(ws, [12, 8, 8, 11, 34])

    out = os.path.join(HERE, '题目台账.xlsx')
    wb.save(out)
    print('saved', out)

    # ============ Markdown ============
    md = []
    md.append('# 已用题目台账（cdf归档 · 22 场 · 101 题）\n')
    md.append('> 难度系数 = 1 − 得分率（按本届学生实际表现，非客观难度）；确认状态：**确认**=已核实出处，**待核**=疑似待确认，**原创**=校内题。')
    md.append('> 考点列"待补充"处请老师补填。未启用归档（0826S加赛、两场 senior、2025-08 三场）见文末。\n')
    cur_label = None
    md.append('| 使用日期 | 场次 | 题号 | 代号 | 来源（出处） | 考点/技巧 | 状态 | 得分率 | 难度系数 | 难度档 | 备注 |')
    md.append('|---|---|---|---|---|---|---|---|---|---|---|')
    for rw in rows:
        st = {'确认': '确认', '待核': '待核', '原创': '原创'}[rw['status']]
        note = rw['note']
        md.append(f"| {rw['date']} | {rw['label']} | T{rw['tidx']} | {rw['name']} | {rw['src']} "
                  f"| {rw['tags']} | {st} | {rw['rate']:.1%} | {1-rw['rate']:.2f} | {band(rw['rate'])} | {note} |")
    md.append('\n## 未启用归档（无学生成绩）\n')
    md.append('| 场次 | 题目 | 来源 | 备注 |')
    md.append('|---|---|---|---|')
    for row in UNUSED:
        md.append(f"| {row[0]} | {row[2]} | {row[3]} | {row[6]} |")
    md.append('\n---\n')
    md.append(f'### 已核实出处一览（确认状态 = 确认 的 {sum(1 for r2 in rows if r2["status"]=="确认")} 题）\n')
    for rw in rows:
        if rw['status'] == '确认':
            md.append(f"- **{rw['name']}**（{rw['label']} T{rw['tidx']}，{rw['date']}）→ {rw['src']}｜考点：{rw['tags']}")
    out_md = os.path.join(HERE, '题目台账.md')
    with open(out_md, 'w', encoding='utf-8') as f:
        f.write('\n'.join(md))
    print('saved', out_md)


if __name__ == '__main__':
    main()
