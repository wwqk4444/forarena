# -*- coding: utf-8 -*-
"""读取 parsed.json，生成 成绩统计大表.xlsx 与 成绩统计大表.md。"""
import json, os, statistics
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------- 基本配置 ----------------
LABEL = {  # 文件名 -> 短标签（矩阵列名 / 明细 sheet 名）
    '20260620 tongbu.cdf': '0620同步',
    '0701.cdf': '0701',
    '260702S.cdf': '0702S',
    '0703S.cdf': '0703S',
    '260706S.cdf': '0706S',
    '260707S.cdf': '0707S',
    '260708S.cdf': '0708S',
    '0710S.cdf': '0710S',
    '0713S.cdf': '0713S',
    '0714S.cdf': '0714S',
    'contest.cdf': '0715赛',
    '0716S.cdf': '0716S',
    '0717S.cdf': '0717S',
    '20260717入门级.cdf': '0717入门',
    '0822test.cdf': '0822',
    '0823test.cdf': '0823',
    '0824Stest.cdf': '0824S',
    '0825test.cdf': '0825',
    '0826S.cdf': '0826S',
    '0826testrumen.cdf': '0826入门',
    '260905test.cdf': '0905',
    '0912S.cdf': '0912',
    '20250823 .cdf': '2025-0823',
    '20250830.cdf': '2025-0830',
    '20258022.cdf': '2025-0822',
    '0826S加赛.cdf': '0826S加赛',
    '20260715 senior.cdf': '0715senior',
    '20260717 senior.cdf': '0717senior',
}
EXCLUDE_NOTE = {'std': '标程', 'solutions': '题解', 'wwqk4444': '管理员帐号',
                '姓名年级补题': '占位模板', '姓名补题年级': '占位模板'}

# ---------------- 样式 ----------------
HDR_FILL = PatternFill('solid', fgColor='305496')
SUB_FILL = PatternFill('solid', fgColor='8EA9DB')
TITLE_FILL = PatternFill('solid', fgColor='203864')
G9 = PatternFill('solid', fgColor='C6EFCE')   # >=90%
G7 = PatternFill('solid', fgColor='E2EFDA')   # >=70%
G4 = PatternFill('solid', fgColor='FFEB9C')   # >=40%
G1 = PatternFill('solid', fgColor='FCE4D6')   # >0
G0 = PatternFill('solid', fgColor='D9D9D9')   # =0
WHITE_BOLD = Font(bold=True, color='FFFFFF')
BOLD = Font(bold=True)
SMALL = Font(size=9, color='595959')
CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def rate_fill(rate):
    if rate >= 0.9: return G9
    if rate >= 0.7: return G7
    if rate >= 0.4: return G4
    if rate > 0: return G1
    return G0


def fmt_num(x):
    if x is None: return ''
    if isinstance(x, float) and abs(x - round(x)) < 1e-6: x = int(round(x))
    return x


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def main():
    with open(os.path.join(HERE, 'parsed.json'), encoding='utf-8') as f:
        data = json.load(f)
    junior_names = set(data['junior_names'])
    contests = data['contests']

    # ---- 排序：按日期 ----
    def sort_key(c):
        d = c.get('judge_date') or '9999'
        return (d, LABEL.get(c['file'], c['file']))
    contests.sort(key=sort_key)
    with_data = [c for c in contests if c['students']]
    senior_ct = [c for c in with_data if c['group'] == '中学组']
    junior_ct = [c for c in with_data if c['group'] == '入门组']

    wb = Workbook()

    # ============ Sheet: 说明 ============
    ws = wb.active
    ws.title = '说明'
    lines = [
        ('成绩统计大表 —— cdf归档 全量统计', TITLE_FILL, WHITE_BOLD, 14),
        ('', None, None, None),
        ('数据来源：本仓库 cdf归档/ 目录下全部 28 个 .cdf 评测归档文件。', None, None, None),
        ('统计范围：其中 22 场含学生成绩（中学组 19 场、入门组 3 场）；', None, None, None),
        ('　2025年8月三场(20250822/0823/0830)仅含标程std、0826S加赛与20260715/20260717 senior 无选手成绩，只列入"比赛一览"。', None, None, None),
        ('', None, None, None),
        ('统计口径：', HDR_FILL, WHITE_BOLD, 11),
        ('1. 题目得分：按测试点求和。测试点内含多个输入文件时，按"通过率×该测试点满分"折算（与评测机一致）。', None, None, None),
        ('2. 同一学生同场比赛有多条提交记录（如正式+补题各交一次）时，逐题取已评测的最高分合并；合并的原始帐号在明细表"原始帐号"列列出。', None, None, None),
        ('3. 排名：按合并后总分从高到低，同分并列（1,2,2,4式）。', None, None, None),
        ('4. 不计入成绩的帐号：std(标程)、solutions(题解)、wwqk4444(管理员)、"姓名年级补题/姓名补题年级"(占位模板)。', None, None, None),
        ('5. 帐号归并：按姓名归并（如"丁冠华_六升七_补题"="丁冠华"）；拼音帐号已映射（liminghe=李明赫、nichunyu=倪淳彧、LuBoren=陆博仁、', None, None, None),
        ('　 shijingxin=史靖心、zbl_buti=张柏林、caiboan=蔡伯安）；无法确认归属的帐号（tjx、tjxnb、乌鲁鲁、我想吃、最强奶龙、兰田之胱、', None, None, None),
        ('　 FarmerJohn2、Gay帐号、将于 等）按原帐号单列，带※标记。', None, None, None),
        ('6. 颜色：绿≥90%　浅绿≥70%　黄≥40%　橙>0　灰=0　空白=未参加。', None, None, None),
        ('', None, None, None),
        ('工作表导航：', HDR_FILL, WHITE_BOLD, 11),
        ('比赛一览：27 个归档文件的基本信息与整场统计。', None, None, None),
        ('中学组成绩矩阵 / 中学组排名矩阵：中学组学生 × 19 场比赛的得分与名次大表。', None, None, None),
        ('入门组成绩矩阵 / 入门组排名矩阵：入门组学生 × 3 场比赛（0620同步、0717入门、0826入门）。', None, None, None),
        ('明细·XX：每场比赛的成绩单（含每题得分、合并帐号说明）。', None, None, None),
        ('题目难度分析：每场比赛每题的平均分、得分率、满分/零分人数。', None, None, None),
        ('', None, None, None),
        ('※ 帐号名与真实姓名的对应关系由脚本自动归并，个别错漏可在 cdf归档 原始记录中核对。', None, SMALL, None),
    ]
    for i, (txt, fill, font, size) in enumerate(lines, 1):
        cell = ws.cell(row=i, column=1, value=txt)
        if fill: cell.fill = fill
        if font: cell.font = font
        if size and font is None: cell.font = Font(size=size, bold=True)
    ws.column_dimensions['A'].width = 110

    # ============ Sheet: 比赛一览 ============
    ws = wb.create_sheet('比赛一览')
    headers = ['文件名', '比赛标题', '日期', '组别', '题数', '题目(满分)', '总满分',
               '有效人数', '平均分', '最高分', '满分人数', '备注']
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill = HDR_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BORDER
    r = 2
    for cinfo in contests:
        n = len(cinfo['students'])
        totals = [s['total'] for s in cinfo['students']]
        full = cinfo['full_total']
        task_str = ' '.join(f"{tn}({fs})" for tn, fs in cinfo['tasks'])
        if len(task_str) > 60: task_str = task_str[:58] + '…'
        note = []
        if not n:
            note.append('归档无学生成绩' + ('（仅标程）' if cinfo['entries'] else '（空）'))
        if cinfo['file'] == 'contest.cdf':
            note.append('标题为contest，评测日2026-07-15')
        unj = sum(1 for e in cinfo['entries'] if not e['has_data'])
        if unj: note.append(f'{unj}条未评测记录(计0分)')
        excl = cinfo.get('excluded') or []
        if excl:
            note.append('剔除帐号: ' + '、'.join(f"{x}({EXCLUDE_NOTE.get(x,'?')})" for x in excl))
        vals = [cinfo['file'], cinfo['title'].strip(), cinfo.get('judge_date') or '—',
                cinfo['group'], len(cinfo['tasks']), task_str, full,
                n if n else None,
                round(sum(totals) / n, 1) if n else None,
                max(totals) if n else None,
                sum(1 for t in totals if t == full) if n else None,
                '；'.join(note)]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=fmt_num(v))
            c.border = BORDER
            c.alignment = LEFT if j in (1, 2, 6, 12) else CENTER
        r += 1
    set_widths(ws, [20, 16, 11, 8, 5, 52, 7, 8, 7, 7, 8, 46])
    ws.freeze_panes = 'A2'

    # ============ 成绩/排名矩阵 ============
    def student_rows(ct_list, group):
        """返回按总分排序的 (name, per-contest dict) 列表"""
        agg = {}
        for cinfo in ct_list:
            for s in cinfo['students']:
                a = agg.setdefault(s['name'], {'scores': {}, 'ranks': {}, 'mapped': s['mapped']})
                a['scores'][cinfo['file']] = s['total']
                a['ranks'][cinfo['file']] = s['rank']
        rows = []
        for name, a in agg.items():
            tot = round(sum(a['scores'].values()), 1)
            n = len(a['scores'])
            avg = round(tot / n, 1) if n else 0
            best = min(a['ranks'].values())
            avgrank = round(sum(a['ranks'].values()) / len(a['ranks']), 1) if a['ranks'] else None
            rows.append({'name': name, 'a': a, 'total': tot, 'n': n, 'avg': avg,
                         'best': best, 'avgrank': avgrank, 'mapped': a['mapped']})
        rows.sort(key=lambda x: (-x['total'], -x['avg'], x['name']))
        return rows

    def write_matrix(ws, ct_list, rows, kind):
        """kind: 'score' or 'rank'"""
        n_ct = len(ct_list)
        # 标题行
        ws.cell(row=1, column=1, value=f"{'成绩矩阵' if kind=='score' else '排名矩阵'}（{len(rows)}人 × {n_ct}场）"
                + ('　带※为未能对应真实姓名的帐号' if kind == 'score' else ''))
        ws.cell(row=1, column=1).font = Font(bold=True, size=12)
        # 表头
        cols = ['姓名'] + [LABEL[c['file']] for c in ct_list] + \
               ['参赛\n场数', '总分', '场均', '满分\n场数', '最佳\n名次', '平均\n名次']
        for j, h in enumerate(cols, 1):
            c = ws.cell(row=2, column=j, value=h)
            c.fill = HDR_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BORDER
            ws.row_dimensions[2].height = 26
        # 满分行
        ws.cell(row=3, column=1, value='满分').fill = SUB_FILL
        ws.cell(row=3, column=1).font = BOLD
        ws.cell(row=3, column=1).alignment = CENTER
        for j, cinfo in enumerate(ct_list, 2):
            c = ws.cell(row=3, column=j, value=cinfo['full_total'])
            c.fill = SUB_FILL; c.alignment = CENTER; c.border = BORDER; c.font = BOLD
        for j in range(n_ct + 2, n_ct + 8):
            ws.cell(row=3, column=j).fill = SUB_FILL
            ws.cell(row=3, column=j).border = BORDER
        # 数据行
        r = 4
        for row in rows:
            name = row['name'] + ('' if row['mapped'] else '※')
            c = ws.cell(row=r, column=1, value=name)
            c.border = BORDER; c.alignment = LEFT; c.font = BOLD
            for j, cinfo in enumerate(ct_list, 2):
                f = cinfo['file']
                cell = ws.cell(row=r, column=j)
                cell.border = BORDER; cell.alignment = CENTER
                if f in row['a']['scores']:
                    if kind == 'score':
                        v = fmt_num(row['a']['scores'][f])
                        cell.value = v
                        cell.fill = rate_fill(row['a']['scores'][f] / cinfo['full_total'])
                    else:
                        cell.value = row['a']['ranks'][f]
                if kind == 'rank' and f in row['a']['ranks']:
                    rk = row['a']['ranks'][f]
                    if rk == 1: cell.font = Font(bold=True, color='C00000')
                    elif rk <= 3: cell.font = Font(bold=True, color='7030A0')
            # 汇总列
            summary = [row['n'], fmt_num(row['total']), fmt_num(row['avg']),
                       sum(1 for f, v in row['a']['scores'].items()
                           if v == next(ci['full_total'] for ci in ct_list if ci['file'] == f) and v > 0),
                       row['best'],
                       round(row['avgrank'], 1) if row['avgrank'] else None]
            for k, v in enumerate(summary):
                cell = ws.cell(row=r, column=n_ct + 2 + k, value=v)
                cell.border = BORDER; cell.alignment = CENTER
                if k in (1, 4): cell.font = BOLD
            r += 1
        # 底部：参赛人数
        ws.cell(row=r, column=1, value='参赛人数').font = BOLD
        ws.cell(row=r, column=1).fill = SUB_FILL
        ws.cell(row=r, column=1).border = BORDER
        for j, cinfo in enumerate(ct_list, 2):
            cell = ws.cell(row=r, column=j, value=len(cinfo['students']))
            cell.fill = SUB_FILL; cell.border = BORDER; cell.alignment = CENTER; cell.font = BOLD
        for j in range(n_ct + 2, n_ct + 8):
            ws.cell(row=r, column=j).fill = SUB_FILL
            ws.cell(row=r, column=j).border = BORDER
        set_widths(ws, [14] + [8.5] * n_ct + [7, 8, 7, 7, 7, 7])
        ws.freeze_panes = 'B4'

    # 中学组
    s_rows = student_rows(senior_ct, '中学组')
    ws = wb.create_sheet('中学组成绩矩阵')
    write_matrix(ws, senior_ct, s_rows, 'score')
    ws = wb.create_sheet('中学组排名矩阵')
    write_matrix(ws, senior_ct, s_rows, 'rank')
    # 入门组
    j_rows = student_rows(junior_ct, '入门组')
    ws = wb.create_sheet('入门组成绩矩阵')
    write_matrix(ws, junior_ct, j_rows, 'score')
    ws = wb.create_sheet('入门组排名矩阵')
    write_matrix(ws, junior_ct, j_rows, 'rank')

    # ============ 明细 sheets ============
    for cinfo in with_data:
        lab = LABEL[cinfo['file']]
        ws = wb.create_sheet('明细·' + lab)
        ntask = len(cinfo['tasks'])
        # 标题
        ws.cell(row=1, column=1,
                value=f"{lab}　{cinfo['title'].strip()}　{cinfo.get('judge_date','')}　"
                      f"共{len(cinfo['tasks'])}题 满分{cinfo['full_total']}　{len(cinfo['students'])}人")
        ws.cell(row=1, column=1).font = Font(bold=True, size=12)
        # 表头
        ws.cell(row=2, column=1, value='排名')
        ws.cell(row=2, column=2, value='姓名')
        ws.cell(row=2, column=3, value='提交\n次数')
        for j, (tn, fs) in enumerate(cinfo['tasks'], 4):
            ws.cell(row=2, column=j, value=str(tn))
            ws.cell(row=3, column=j, value=fs)
        last = 4 + ntask
        tails = ['总分', '满分', '得分率', '满分\n题数', '原始帐号(归并前)']
        for k, h in enumerate(tails):
            ws.cell(row=2, column=last + k, value=h)
        for j in range(1, last + len(tails)):
            for rr in (2, 3):
                cell = ws.cell(row=rr, column=j)
                cell.fill = HDR_FILL if rr == 2 else SUB_FILL
                if rr == 2: cell.font = WHITE_BOLD
                cell.alignment = CENTER; cell.border = BORDER
        ws.cell(row=3, column=1, value='—').fill = SUB_FILL
        ws.cell(row=3, column=2, value='满分').fill = SUB_FILL
        ws.cell(row=3, column=2).font = BOLD
        ws.cell(row=3, column=3).fill = SUB_FILL
        ws.row_dimensions[2].height = 24
        # 数据
        r = 4
        for s in cinfo['students']:
            cell = ws.cell(row=r, column=1, value=s['rank']); cell.alignment = CENTER; cell.border = BORDER
            if s['rank'] == 1: cell.font = Font(bold=True, color='C00000')
            elif s['rank'] <= 3: cell.font = Font(bold=True, color='7030A0')
            cell = ws.cell(row=r, column=2, value=s['name'] + ('' if s['mapped'] else '※'))
            cell.border = BORDER; cell.alignment = LEFT; cell.font = BOLD
            cell = ws.cell(row=r, column=3, value=s['n_entries'] if s['n_entries'] > 1 else '')
            cell.border = BORDER; cell.alignment = CENTER
            for j, (v, (tn, fs)) in enumerate(zip(s['ps'], cinfo['tasks']), 4):
                cell = ws.cell(row=r, column=j, value=fmt_num(v))
                cell.border = BORDER; cell.alignment = CENTER
                cell.fill = rate_fill(v / fs if fs else 0)
            cell = ws.cell(row=r, column=last, value=fmt_num(s['total']))
            cell.border = BORDER; cell.alignment = CENTER; cell.font = BOLD
            cell = ws.cell(row=r, column=last + 1, value=cinfo['full_total'])
            cell.border = BORDER; cell.alignment = CENTER
            cell = ws.cell(row=r, column=last + 2, value=round(s['total'] / cinfo['full_total'], 3))
            cell.number_format = '0.0%'
            cell.border = BORDER; cell.alignment = CENTER
            cell = ws.cell(row=r, column=last + 3,
                           value=sum(1 for v, (tn, fs) in zip(s['ps'], cinfo['tasks']) if v == fs and fs > 0))
            cell.border = BORDER; cell.alignment = CENTER
            cell = ws.cell(row=r, column=last + 4,
                           value=' = '.join(s['raws']) if s['n_entries'] > 1 else '')
            cell.font = SMALL; cell.border = BORDER
            r += 1
        # 统计行
        ws.cell(row=r, column=2, value='平均分').font = BOLD
        ws.cell(row=r, column=2).fill = SUB_FILL; ws.cell(row=r, column=2).border = BORDER
        for j in (1, 3): ws.cell(row=r, column=j).fill = SUB_FILL
        for j, (tn, fs) in enumerate(cinfo['tasks'], 4):
            vals = [s['ps'][j - 4] for s in cinfo['students']]
            cell = ws.cell(row=r, column=j, value=round(sum(vals) / len(vals), 1))
            cell.fill = SUB_FILL; cell.border = BORDER; cell.alignment = CENTER
        vals = [s['total'] for s in cinfo['students']]
        cell = ws.cell(row=r, column=last, value=round(sum(vals) / len(vals), 1))
        cell.fill = SUB_FILL; cell.border = BORDER; cell.alignment = CENTER; cell.font = BOLD
        r += 1
        ws.cell(row=r, column=2, value='满分人数').font = BOLD
        ws.cell(row=r, column=2).fill = SUB_FILL; ws.cell(row=r, column=2).border = BORDER
        for j in (1, 3): ws.cell(row=r, column=j).fill = SUB_FILL
        for j, (tn, fs) in enumerate(cinfo['tasks'], 4):
            cnt = sum(1 for s in cinfo['students'] if s['ps'][j - 4] == fs and fs > 0)
            cell = ws.cell(row=r, column=j, value=cnt)
            cell.fill = SUB_FILL; cell.border = BORDER; cell.alignment = CENTER
        cnt = sum(1 for s in cinfo['students'] if s['total'] == cinfo['full_total'])
        cell = ws.cell(row=r, column=last, value=cnt)
        cell.fill = SUB_FILL; cell.border = BORDER; cell.alignment = CENTER
        set_widths(ws, [6, 16, 6] + [8] * ntask + [8, 7, 8, 7, 40])
        ws.freeze_panes = 'D4'

    # ============ 题目难度分析 ============
    ws = wb.create_sheet('题目难度分析')
    headers = ['比赛', '日期', '组别', '题号', '题名', '满分', '有效人数', '平均分',
               '得分率', '中位数', '满分人数', '满分率', '零分人数', '零分率']
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill = HDR_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BORDER
    r = 2
    for cinfo in with_data:
        for qi, (tn, fs) in enumerate(cinfo['tasks']):
            vals = [s['ps'][qi] for s in cinfo['students']]
            n = len(vals)
            avg = round(sum(vals) / n, 1)
            med = round(statistics.median(vals), 1)
            fulln = sum(1 for v in vals if v == fs and fs > 0)
            zeron = sum(1 for v in vals if v == 0)
            vals2 = [cinfo['file'], cinfo.get('judge_date'), cinfo['group'], f"T{qi+1}", str(tn),
                     fs, n, fmt_num(avg), avg / fs if fs else 0, fmt_num(med),
                     fulln, round(fulln / n, 3), zeron, round(zeron / n, 3)]
            for j, v in enumerate(vals2, 1):
                cell = ws.cell(row=r, column=j, value=v)
                cell.border = BORDER
                cell.alignment = CENTER if j != 5 else LEFT
                if j == 9:
                    cell.number_format = '0.0%'
                    cell.fill = rate_fill(v)
                if j in (12, 14): cell.number_format = '0.0%'
            r += 1
    set_widths(ws, [16, 11, 8, 6, 18, 7, 9, 8, 8, 8, 9, 8, 9, 8])
    ws.freeze_panes = 'A2'

    out_xlsx = os.path.join(HERE, '成绩统计大表.xlsx')
    wb.save(out_xlsx)
    print('saved', out_xlsx)

    # ============ Markdown 摘要 ============
    def num_s(x):
        x = fmt_num(x)
        return str(x)

    md = []
    md.append('# 成绩统计大表（cdf归档 全量统计）\n')
    md.append('> 生成自 `cdf归档/` 下全部 28 个 .cdf 评测归档；详细数据见 **成绩统计大表.xlsx**。')
    md.append('> 口径：同场多次提交（正式/补题）按题取最高分合并；排名同分并列；std/题解/管理员/模板帐号不计入。\n')

    md.append('## 一、比赛一览\n')
    md.append('| # | 比赛 | 日期 | 组别 | 题数 | 总满分 | 人数 | 平均分 | 最高分 | 满分 | 备注 |')
    md.append('|---|------|------|------|------|--------|------|--------|--------|------|------|')
    for i, cinfo in enumerate(contests, 1):
        n = len(cinfo['students'])
        totals = [s['total'] for s in cinfo['students']]
        note = ''
        if not n:
            note = '归档无学生成绩' + ('（仅标程）' if cinfo['entries'] else '（空）')
        elif cinfo['file'] == 'contest.cdf':
            note = '标题contest，评测日07-15'
        md.append(f"| {i} | {LABEL.get(cinfo['file'], cinfo['file'])} | {cinfo.get('judge_date') or '—'} "
                  f"| {cinfo['group']} | {len(cinfo['tasks'])} | {cinfo['full_total']} "
                  f"| {n} | {round(sum(totals)/n,1) if n else '—'} | {num_s(max(totals)) if n else '—'} "
                  f"| {sum(1 for t in totals if t==cinfo['full_total']) if n else '—'} | {note} |")
    md.append('')

    for gname, ct_list, rows in [('中学组', senior_ct, s_rows), ('入门组', junior_ct, j_rows)]:
        md.append(f'## {"二" if gname=="中学组" else "三"}、{gname}个人总表（{len(rows)}人 × {len(ct_list)}场）\n')
        if gname == '中学组':
            md.append('总分列后为每场"分数(名次)"；※为未能对应真实姓名的帐号。\n')
            md.append('| 姓名 | ' + ' | '.join(LABEL[c['file']] for c in ct_list) +
                      ' | 场数 | 总分 | 场均 | 最佳名次 |')
            md.append('|------|' + '---|' * (len(ct_list) + 4))
            for row in rows:
                cells = [row['name'] + ('' if row['mapped'] else '※')]
                for cinfo in ct_list:
                    f = cinfo['file']
                    if f in row['a']['scores']:
                        cells.append(f"{num_s(row['a']['scores'][f])}({row['a']['ranks'][f]})")
                    else:
                        cells.append('')
                cells += [str(row['n']), num_s(row['total']), num_s(row['avg']), str(row['best'])]
                md.append('| ' + ' | '.join(cells) + ' |')
        else:
            md.append('| 姓名 | ' + ' | '.join(LABEL[c['file']] for c in ct_list) +
                      ' | 场数 | 总分 | 场均 | 最佳名次 |')
            md.append('|------|' + '---|' * (len(ct_list) + 4))
            for row in rows:
                cells = [row['name'] + ('' if row['mapped'] else '※')]
                for cinfo in ct_list:
                    f = cinfo['file']
                    cells.append(f"{num_s(row['a']['scores'][f])}({row['a']['ranks'][f]})" if f in row['a']['scores'] else '')
                cells += [str(row['n']), num_s(row['total']), num_s(row['avg']), str(row['best'])]
                md.append('| ' + ' | '.join(cells) + ' |')
        md.append('')

    md.append('## 四、每场比赛前三名\n')
    md.append('| 比赛 | 日期 | 🥇 第一 | 🥈 第二 | 🥉 第三 |')
    md.append('|------|------|--------|--------|--------|')
    for cinfo in with_data:
        tops = cinfo['students'][:3]
        cells = [f"{s['name']}（{num_s(s['total'])}分）" for s in tops]
        while len(cells) < 3: cells.append('—')
        md.append(f"| {LABEL[cinfo['file']]} | {cinfo.get('judge_date')} | " + ' | '.join(cells) + ' |')
    md.append('')

    md.append('## 五、题目难度速览（得分率最低 = 最难）\n')
    all_probs = []
    for cinfo in with_data:
        for qi, (tn, fs) in enumerate(cinfo['tasks']):
            vals = [s['ps'][qi] for s in cinfo['students']]
            avg = sum(vals) / len(vals)
            all_probs.append((avg / fs if fs else 0, LABEL[cinfo['file']], str(tn), fs,
                              round(avg, 1), sum(1 for v in vals if v == fs and fs > 0), len(vals)))
    all_probs.sort()
    md.append('最难 10 题：\n')
    md.append('| 比赛 | 题名 | 满分 | 平均分 | 得分率 | 满分人数/参赛 |')
    md.append('|------|------|------|--------|--------|----------------|')
    for rate, lab, tn, fs, avg, fn, n in all_probs[:10]:
        md.append(f'| {lab} | {tn} | {fs} | {avg} | {rate:.1%} | {fn}/{n} |')
    md.append('')
    md.append('最易 10 题（得分率最高）：\n')
    md.append('| 比赛 | 题名 | 满分 | 平均分 | 得分率 | 满分人数/参赛 |')
    md.append('|------|------|------|--------|--------|----------------|')
    for rate, lab, tn, fs, avg, fn, n in sorted(all_probs, reverse=True)[:10]:
        md.append(f'| {lab} | {tn} | {fs} | {avg} | {rate:.1%} | {fn}/{n} |')
    md.append('')
    md.append('---\n')
    md.append('### 统计口径备注\n')
    md.append('- **得分折算**：测试点内多个输入文件时按通过率×满分折算（与评测机一致）。')
    md.append('- **多次提交**：同一场比赛同一学生"正式 + 补题"多条记录，逐题取已评测最高分合并（明细表"原始帐号"列可查）。')
    md.append('- **已剔除帐号**：std（标程）、solutions（题解）、wwqk4444（管理员）、姓名年级补题/姓名补题年级（占位模板）。')
    md.append('- **拼音帐号映射**：liminghe→李明赫，nichunyu→倪淳彧，LuBoren→陆博仁，shijingxin→史靖心，zbl_buti→张柏林，caiboan→蔡伯安。')
    md.append('- **未归并帐号（※）**：tjx、tjxnb、乌鲁鲁、我想吃、最强奶龙、兰田之胱、FarmerJohn2、Gay帐号、将于——无法确认真实姓名，按原帐号单列。')
    md.append('- 2025-08 的三场（20250822/0823/0830）归档仅含标程，0826S加赛、20260715/20260717 senior 无选手成绩，均未计入个人总表。')
    out_md = os.path.join(HERE, '成绩统计大表.md')
    with open(out_md, 'w', encoding='utf-8') as f:
        f.write('\n'.join(md))
    print('saved', out_md)


if __name__ == '__main__':
    main()
