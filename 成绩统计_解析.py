# -*- coding: utf-8 -*-
"""解析 cdf归档/*.cdf，归一化学生姓名，输出中间结果 JSON 供生成报表使用。"""
import json, re, glob, os, sys, datetime
from collections import OrderedDict

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'cdf归档')

# ---------------- 已知学生姓名（中学组） ----------------
SENIOR_NAMES = """丁欣森 丁冠华 倪淳彧 刘善岩 刘辰翔 史如茗 姚畅 姚艺宬 季子清 尤晗昱
张亦夫 张尧轩 张思宇 张智淏 张弛 张柏林 徐康杰 徐德屹 忻士云 施皓腾 易皓轩 李一晨 李明赫
杜承泽 杨孙尧 杨耘齐 檀炎希 江裕 汤景行 沈天星 沈峻霆 沈徽 沈永承 涂泽宇 滕梓睿 潘家文 潘史亮
王业阳 王东诚 王乐维 王奕涵 王子涵 王宣贻 王熙辰 王衿 王金翔 王哲熙 田函硕 祁宥嘉 苏畅 苏睿洋
范子源 蔡伯安 蔡敦行 蔡涵秋 薛梓轩 邵元瞰 邵辰璋 郝熠轩 金启轩 钟箫谦 钟远 陆博仁 陈子昂 陈子琚
陈斯涵 严一篪 乔云起 顾语承 张曦元 史靖心 朱浩坤 朱鸿 罗晨轩""".split()

# 拼音/英文帐号别名
ALIAS = [
    ('luboren', '陆博仁'),
    ('liminghe', '李明赫'),
    ('nichunyu', '倪淳彧'),
    ('shijingxin', '史靖心'),
    ('zbl_buti', '张柏林'),
    ('caiboan', '蔡伯安'),
    ('gay', 'Gay帐号'),
]

# 不计入学生成绩的帐号（管理员/标程/占位模板）
EXCLUDE = {'std', 'solutions', 'wwqk4444', '姓名年级补题', '姓名补题年级'}

JUNIOR_FILES = ['20260620 tongbu.cdf', '0826testrumen.cdf', '20260717入门级.cdf']

# 入门组姓名在解析入门组文件时自动提取；此处只做异体字归并
VARIANT = {'蕫': '董'}


def load(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def extract_junior_names():
    names = set()
    for fn in JUNIOR_FILES:
        d = load(os.path.join(SRC, fn))
        for c in d['contestants']:
            r = c['contestantName'].strip()
            if r in EXCLUDE:
                continue
            nm = re.sub(r'^\d{2,3}-\d{1,2}_', '', r)      # 去掉班级座号前缀
            nm = re.sub(r'^\d{2,3}-\d{1,2}\s*', '', nm)   # 去掉重复出现的座号
            cn = re.findall(r'[\u4e00-\u9fff]+', nm)
            if cn:
                s = ''.join(cn)
                s = ''.join(VARIANT.get(ch, ch) for ch in s)
                names.add(s)
    return sorted(names, key=lambda x: -len(x))


def make_resolver(senior, junior):
    known = sorted(set(senior) | set(junior), key=lambda x: -len(x))
    junior_set = set(junior)

    def resolve(raw):
        """返回 (身份名, 是否映射到真实姓名)"""
        r = raw.strip()
        r = ''.join(VARIANT.get(ch, ch) for ch in r)
        if r in EXCLUDE:
            return None, False
        if r.lower() == 'tjx':
            return 'tjx', False
        if r.lower() == 'tjxnb':
            return 'tjxnb', False
        # 帐号名里找已知中文姓名（长名优先、占位不重叠）
        taken = [False] * len(r)
        found = []
        for nm in known:
            start = 0
            while True:
                i = r.find(nm, start)
                if i < 0:
                    break
                if not any(taken[i:i + len(nm)]):
                    for k in range(i, i + len(nm)):
                        taken[k] = True
                    found.append(nm)
                start = i + 1
        if found:
            # 多人合测（帐号里含两个姓名）：用 "+" 连接
            name = '+'.join(sorted(found))
            return name, True
        low = r.lower()
        for key, full in ALIAS:
            if key in low:
                return full, True
        # 未匹配：取第一个≥2字的中文串（人名/昵称），否则保留英文帐号
        cn = re.findall(r'[\u4e00-\u9fff]+', r)
        cand = [x for x in cn if len(x) >= 2]
        if cand:
            s = cand[0]
            s = ''.join(VARIANT.get(ch, ch) for ch in s)
            return s, False
        return r, False

    return resolve


def parse_file(path):
    d = load(path)
    tasks = d['tasks']
    task_full = [sum(tc['fullScore'] for tc in t['testCases']) for t in tasks]
    task_names = [t.get('problemTitle', '?') for t in tasks]
    entries = []
    for c in d['contestants']:
        judged = c.get('checkJudged') or []
        comp = c.get('compileState') or []
        scores = c.get('score') or []
        ps, pj = [], []
        for ti in range(len(tasks)):
            # 内层列表 = 该测试点内各输入文件得分(通过=fullScore, 否则0)
            # 测试点得分 = 均值 × 满分 = mean(block)
            s = 0.0
            for block in (scores[ti] if ti < len(scores) else []):
                if len(block):
                    s += sum(block) / len(block)
            s = round(s + 1e-9, 1)
            if abs(s - round(s)) < 1e-6:
                s = int(round(s))
            ps.append(s)
            pj.append(bool(judged[ti]) if ti < len(judged) else True)
        has_data = any(pj) and any(s > 0 for s in ps) or (any(pj) and any((comp[ti] == 0) for ti in range(len(tasks)) if pj[ti]))
        entries.append({
            'raw': c['contestantName'].strip(),
            'ps': ps, 'pj': pj,
            'ce': [comp[ti] != 0 for ti in range(min(len(comp), len(tasks)))],
            'has_data': any(pj),   # 至少有一题被评测过
        })
    jd = None
    if d['contestants'] and d['contestants'][0].get('judgingTime_date'):
        jd = d['contestants'][0]['judgingTime_date']
        jd = datetime.date.fromordinal(datetime.date(1970, 1, 1).toordinal() + jd - 2440588).isoformat()
    return {
        'file': os.path.basename(path),
        'title': d.get('contestTitle', ''),
        'judge_date': jd,
        'tasks': list(zip(task_names, task_full)),
        'entries': entries,
        'full_total': sum(task_full),
    }


def main():
    junior = extract_junior_names()
    senior = SENIOR_NAMES
    resolve = make_resolver(senior, junior)
    files = sorted(glob.glob(os.path.join(SRC, '*.cdf')))
    contests = []
    debug = []
    for path in files:
        info = parse_file(path)
        # 归并：同一学生多条提交，按题取已评测的最高分
        merged = OrderedDict()
        for e in info['entries']:
            ident, mapped = resolve(e['raw'])
            if ident is None:
                info.setdefault('excluded', []).append(e['raw'])
                continue
            m = merged.setdefault(ident, {
                'names': set(), 'ps': [None] * len(info['tasks']),
                'n_entries': 0, 'raws': [], 'mapped': mapped,
            })
            m['names'].add(ident)
            m['n_entries'] += 1
            m['raws'].append(e['raw'])
            for ti, (s, j) in enumerate(zip(e['ps'], e['pj'])):
                if j and s is not None:
                    m['ps'][ti] = max(m['ps'][ti] if m['ps'][ti] is not None else -1, s)
        group = '入门组' if info['file'] in JUNIOR_FILES else '中学组'
        students = []
        for ident, m in merged.items():
            ps = [0 if v is None else v for v in m['ps']]
            total = round(sum(ps), 1)
            if abs(total - round(total)) < 1e-6:
                total = int(round(total))
            students.append({
                'name': ident, 'ps': ps, 'total': total,
                'n_entries': m['n_entries'], 'raws': m['raws'],
                'mapped': m['mapped'],
            })
        students.sort(key=lambda x: -x['total'])
        # competition ranking
        rank, prev = 0, None
        for i, s in enumerate(students):
            s['rank'] = rank if s['total'] == prev else i + 1
            prev, rank = s['total'], s['rank']
        info['group'] = group
        info['students'] = students
        contests.append(info)
        for s in students:
            debug.append((info['file'], s['name'], s['rank'], s['total'],
                          'MAP' if s['mapped'] else 'RAW', '|'.join(s['raws'][:3])))

    with open(os.path.join(HERE, 'parsed.json'), 'w', encoding='utf-8') as f:
        json.dump({'contests': contests, 'junior_names': junior}, f, ensure_ascii=False)
    # 调试输出：未映射帐号 + 每场前3
    print('== 未映射/特殊帐号 ==')
    for row in debug:
        if row[4] == 'RAW':
            print(row)
    print('== 每场前3 ==')
    for c in contests:
        top = ', '.join(f"{s['name']}:{s['total']}" for s in c['students'][:3])
        print(f"{c['file']:24} {c['judge_date']} {c['group']} n={len(c['students']):3d} | {top}")


if __name__ == '__main__':
    main()
