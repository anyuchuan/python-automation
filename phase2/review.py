"""第二阶段：独立输出能力体系建议及关系审核，不改写任何第一阶段文件。"""
import argparse
import hashlib
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
SOURCE = '旅游行业招聘岗位样本库_50条候选_待核验.xlsx'
INPUTS = [SOURCE, 'output/岗位任务编码表.xlsx', 'output/三级职业能力指标体系.xlsx',
          'output/岗位任务能力关系表.xlsx', 'output/岗位数据质量检查.xlsx']
PENDING = '待人工审核'
BOUNDARY = '50条仍为待核验候选样本；摘要不是招聘原文；能力层级、行为指标和课程适配均为研究建议'
COURSE = {'A':'A数据库课程核心相关','B':'B数据库课程延伸相关','C':'C旅游行业背景相关'}
ACTIVE = {'摘要支持（待审）', '待人工判断（仅推断）'}


def read_sheet(path, sheet, required):
    wb = load_workbook(path, data_only=False)
    try:
        if sheet not in wb.sheetnames:
            raise ValueError(f'{path.name}缺少工作表{sheet}')
        values = list(wb[sheet].values)
        headers = list(values[0])
        if len(headers) != len(set(headers)) or not set(required) <= set(headers):
            raise ValueError(f'{path.name}字段缺失或重复：{sheet}')
        return [dict(zip(headers, r)) for r in values[1:] if any(v is not None for v in r)]
    finally:
        wb.close()


def unique(rows, key):
    values = [r[key] for r in rows]
    if any(not v for v in values) or len(set(values)) != len(values):
        raise ValueError(f'{key}为空或重复')
    return {r[key]: r for r in rows}


def load_inputs(root):
    jobs = read_sheet(root/SOURCE, '招聘岗位主库', ['样本ID','岗位职责摘要（转述）','任职要求摘要（不含学历经验）','原始招聘链接','证据摘录','核验状态'])
    tasks = read_sheet(root/INPUTS[1], '任务编码', ['任务编码','岗位ID','证据文本','证据类型'])
    abilities = read_sheet(root/INPUTS[2], '三级能力指标', ['一级编码','二级编码','三级能力编码','三级能力'])
    relations = read_sheet(root/INPUTS[3], '关系明细', ['岗位ID','任务编码','三级能力编码','任务证据','归纳依据'])
    duplicate_rows = read_sheet(root/INPUTS[4], '疑似重复记录', ['重复组','检查类型','岗位ID'])
    js, ts, cs = unique(jobs,'样本ID'), unique(tasks,'任务编码'), unique(abilities,'三级能力编码')
    for t in tasks:
        if t['岗位ID'] not in js or t['证据文本'] not in str(js[t['岗位ID']]['岗位职责摘要（转述）']):
            raise ValueError(f'任务证据无法回溯：{t["任务编码"]}')
    seen = set()
    for r in relations:
        key = (r['岗位ID'],r['任务编码'],r['三级能力编码'])
        if key in seen:
            raise ValueError('第一阶段关系重复键：'+str(key))
        seen.add(key)
        if r['任务编码'] not in ts or r['三级能力编码'] not in cs or r['岗位ID'] not in js:
            raise ValueError('第一阶段关系外键损坏')
        if ts[r['任务编码']]['岗位ID'] != r['岗位ID'] or ts[r['任务编码']]['证据文本'] != r['任务证据']:
            raise ValueError('第一阶段关系证据或岗位ID不一致')
    for level in ['一级','二级']:
        names = defaultdict(set)
        for c in abilities:
            names[c[level+'编码']].add(c[level+'能力'])
        if any(len(v)!=1 for v in names.values()):
            raise ValueError('第一阶段层级名称不唯一')
    return jobs, tasks, abilities, relations, duplicate_rows


def validate_rules(rules, abilities):
    leaves = rules['leaves']
    unique(leaves,'code')
    unique(rules['course_suggestions'],'code')
    old = {c['三级能力编码'] for c in abilities}
    covered = {c for x in leaves for c in x['old']}
    if old != covered:
        raise ValueError('规则与实际原能力编码不一致，请人工更新规则；未覆盖：'+str(old-covered)+'；多出：'+str(covered-old))
    if set(rules['first']) & set(rules['second']):
        raise ValueError('层级编码冲突')
    for leaf in leaves:
        if leaf['code'][:2] not in rules['first'] or leaf['code'][:3] not in rules['second']:
            raise ValueError('新编码无父节点')
        if leaf['course'] not in COURSE:
            raise ValueError('无效课程适配')
        re.compile(leaf['pattern'])


def summary_support(spec, text):
    if not re.search(spec['pattern'], text, re.I):
        return False
    # 备份行为不能在缺少数据对象时直接升级为数据备份。
    if spec['code'] in {'N321','N322'} and not re.search('数据|数据库', text):
        return False
    # 预测建模不能同时无依据地迁移为一般分析模型。
    if spec['code'] == 'N221' and '预测' in text and '分析模型' not in text:
        return False
    # 单独的数据库运维不证明其他业务系统维护，防止对象扩张。
    if spec['code'] == 'N311' and '数据库' in text and not re.search('票务系统|业务系统|景区系统|平台|小程序|CRM|POS|闸机', text):
        return False
    return True


def sql_qualifier(task, job):
    sql = r'(?<![a-zA-Z])SQL(?!\s*Server)'
    if re.search(sql, task['证据文本'], re.I):
        return '职责摘要出现SQL（不是已核验原文）'
    for clause in re.split('[；;，。]', str(job['任职要求摘要（不含学历经验）'] or '')):
        if re.search('未明确|未单独|未.*写明|不要求|无需|待.*核验|需核对', clause):
            continue
        if re.search(sql, clause, re.I):
            return '任职摘要出现SQL，任务使用方式待核验'
    return '未获得该任务使用SQL的明确摘要证据'


def broad_trigger(relation):
    m = re.search('任务触发词：(.*?)；', relation['归纳依据'])
    term = m.group(1) if m else ''
    return term, term in {'分析','统计','维护','运维','系统','指标','模型','需求','优化','管理','流程','配置','实施','故障','监控','报告'}


def relation_review(jobs, tasks, relations, rules):
    js, ts = unique(jobs,'样本ID'), unique(tasks,'任务编码')
    byold = defaultdict(list)
    for x in rules['leaves']:
        for old in x['old']:
            byold[old].append(x)
    originals, proposals = [], []
    for n, old in enumerate(relations, 1):
        rid = f'R1-{n:04d}'
        task, job = ts[old['任务编码']], js[old['岗位ID']]
        targets = byold[old['三级能力编码']]
        support = [x['code'] for x in targets if summary_support(x, task['证据文本'])]
        term, broad = broad_trigger(old)
        reasons = []
        if broad:
            reasons.append('原规则仅由宽泛词触发：'+term+'；需确认动作及对象')
        if not support:
            reasons.append('任务未命中新能力的具体动作或对象，不迁入摘要支持频次')
        if old['三级能力编码'] == 'C131' and 'SQL' not in task['证据文本'].upper():
            reasons.append('SQL来自任职摘要或研究连接，不能认定当前任务已明确使用SQL')
        if old['三级能力编码'] == 'C521' and '文档' not in task['证据文本']:
            reasons.append('流程不等于文档；只支持流程分项时不迁移至文档分项')
        if old['三级能力编码'] == 'C412' and '报表' in task['证据文本']:
            reasons.append('报表开发不能自动证明数据库业务应用开发')
        if old['三级能力编码'] in {'C331','C332'} and '数据库' not in task['证据文本']:
            reasons.append('备份或权限安全对象未明确为数据库，数据库技术细节仍属推断')
        suspicious = bool(reasons)
        originals.append(dict(old, **{'原关系ID':rid,'原触发词':term,'疑似误匹配': '是' if suspicious else '否',
            '复核说明':'；'.join(reasons) if reasons else '摘要含较具体动作或对象，仍需人工确认',
            '严格摘要支持新编码':';'.join(support), '原文证据状态':'未核验；本轮仅使用摘要',
            '人工审核状态':PENDING}))
        for x in targets:
            direct = x['code'] in support
            status = '摘要支持（待审）' if direct else ('不建议迁移（拆分项无支持）' if support else '待人工判断（仅推断）')
            proposals.append({'原关系ID':rid,'原三级能力编码':old['三级能力编码'], '新三级能力编码':x['code'],
                '新能力名称':x['name'], '岗位ID':old['岗位ID'],'任务ID':old['任务编码'],
                '任务证据':task['证据文本'],'任职要求摘要':job['任职要求摘要（不含学历经验）'],
                '原任职要求匹配证据':old.get('任职要求证据'), '原要求性质':old.get('能力要求性质'),
                '证据类型':'岗位摘要支持' if direct else '研究推断',
                '证据状态':'待核验；非招聘原文', '关系建议':status,
                '严格支持规则':x['pattern'],'规则判定说明':'摘要出现新能力动作或对象' if direct else '摘要未明确支持该新能力，保留原匹配轨迹供复核',
                'SQL技术限定':sql_qualifier(task,job),
                '数据库对象限定':'任务摘要出现数据库对象' if '数据库' in task['证据文本'] else '任务未明确数据库对象',
                '课程适配类别':COURSE[x['course']], '优化理由':x['reason'],
                '原始招聘链接':job['原始招聘链接'],'原相关性类别':job['相关性类别'],
                '原证据质量初评':job['证据质量（初评）'],'原核验状态':job['核验状态'],
                '本轮岗位状态':'待核验','人工审核状态':PENDING,
                '摘要支持计数资格':'是（仅候选摘要频次）' if direct else '否',
                '招聘原文核验计数资格':'否'})
    return originals, proposals


def make_node(code, level, name, parent, specs, proposals, reason, course_only=False, behavior=None):
    codes = {s['code'] for s in specs}
    rows = [r for r in proposals if r['新三级能力编码'] in codes and r['关系建议'] in ACTIVE]
    direct = [r for r in rows if r['证据类型']=='岗位摘要支持']
    inferred = [r for r in rows if r['证据类型']=='研究推断']
    categories = {s['course'] for s in specs}
    course = next(iter(categories)) if len(categories)==1 else 'B'
    return {'新编码':code, '父编码':parent,'能力层级':level,'能力名称':name,
        '可观察的工作行为':behavior or '归并下属能力的工作结果，具体评价以三级行为指标为准',
        '对应任务ID':';'.join(sorted({r['任务ID'] for r in rows})),
        '来源岗位ID':';'.join(sorted({r['岗位ID'] for r in rows})),
        '摘要支持任务ID':';'.join(sorted({r['任务ID'] for r in direct})),
        '研究推断任务ID':';'.join(sorted({r['任务ID'] for r in inferred})),
        '证据类型':'研究推断（课程建议，无招聘支持）' if course_only else ('；'.join(sorted({r['证据类型'] for r in rows})) or '研究推断（暂无摘要支持关系）'),
        '证据状态':'课程建议；未获具体招聘证据' if course_only else '全部待核验；无招聘原文明示证据',
        '课程适配类别':COURSE['A' if course_only else course],
        '适配判断边界':'研究判断，不等同于课程标准已覆盖；上位混合类别不能代替逐叶判断',
        '优化理由':reason,'人工审核状态':PENDING,
        '原三级能力编码':';'.join(sorted({o for s in specs for o in s.get('old',[])})),
        '摘要支持岗位数（含疑似重复）':len({r['岗位ID'] for r in direct}),
        '摘要支持任务数':len({r['任务ID'] for r in direct}),
        '研究推断关联岗位数（非频次）':len({r['岗位ID'] for r in inferred}),
        '已核验招聘原文岗位数':0,
        '招聘需求频次':'不计入（课程建议）' if course_only else '不计算正式需求频次；摘要候选数另列',
        '能力来源':'课程建设建议（AI辅助研究提出）' if course_only else '岗位任务归纳建议（AI辅助研究提出）',
        '行为指标来源':'研究拟定的可观察评价表述；不是招聘原文或课程标准'}


def build_nodes(rules, proposals):
    specs = rules['leaves']
    nodes = []
    for code, name in rules['first'].items():
        sub = [s for s in specs if s['code'].startswith(code)]
        nodes.append(make_node(code,'一级',name,'',sub,proposals,'按业务工作领域归并，不按课程知识点倒推'))
    for code, name in rules['second'].items():
        sub = [s for s in specs if s['code'].startswith(code)]
        nodes.append(make_node(code,'二级',name,code[:2],sub,proposals,'按工作对象与结果归组，具体评价下沉至三级'))
    for s in specs:
        nodes.append(make_node(s['code'],'三级',s['name'],s['code'][:3],[s],proposals,s['reason'],behavior=s['behavior']))
    suggestions = []
    cs = [dict(s,old=[],course='A') for s in rules['course_suggestions']]
    suggestions.append(make_node('S1','一级','数据库课程建设建议','',cs,[], '课程建议单独成组，不与岗位能力需求统计混合',True))
    suggestions.append(make_node('S11','二级','数据库结构与记录操作建议','S1',cs,[],'仅是课程建设讨论项，暂无具体岗位证据',True))
    for s in cs:
        suggestions.append(make_node(s['code'],'三级',s['name'],'S11',[s],[],s['reason'],True,s['behavior']))
    return nodes, suggestions


def old_hierarchy(abilities):
    nodes = {}
    for a in abilities:
        for level, key, name in [('一级','一级编码','一级能力'),('二级','二级编码','二级能力'),('三级','三级能力编码','三级能力')]:
            nodes[a[key]] = {'code':a[key], 'name':a[name], 'level':level}
    return sorted(nodes.values(),key=lambda x:(len(x['code']),x['code']))


def compare(abilities, rules, nodes):
    mapping, audits = [], []
    oldnodes = old_hierarchy(abilities)
    for old in oldnodes:
        descendant = {a['三级能力编码'] for a in abilities if a['三级能力编码'].startswith(old['code'])}
        target_specs = [s for s in rules['leaves'] if set(s['old']) & descendant]
        if old['level']=='三级':
            target_codes = {s['code'] for s in target_specs}
            action = '拆分' if len(target_codes)>1 else ('合并' if any(len(s['old'])>1 for s in target_specs) else '保留并规范')
        else:
            length = len(old['code'])
            target_codes = {s['code'][:length] for s in target_specs}
            action = '层级重组/名称规范'
        reason = '；'.join(dict.fromkeys(s['reason'] for s in target_specs))
        audits.append({'原编码':old['code'],'原能力层级':old['level'],'原能力名称':old['name'],
            '名称审核':'建议用动作+对象统一名称；上位保持工作领域',
            '层级审核':'原编码父子关系可解析；工具、行为、业务目的存在混合，按新旧映射重组',
            '重复交叉审核':'查询与SQL工具条目合并' if old['code'] in {'C113','C131'} else '相邻能力按对象和行为界定；可共同支持同一任务但不相加岗位频次',
            '可观察可评价审核':'以三级动作、对象及工作结果评价' if old['level']!='三级' else '新工作行为为研究拟定的评价表述，须专家确认任务粒度',
            '优化动作':action,'新编码集合':';'.join(sorted(target_codes)),'优化理由':reason,'人工审核状态':PENDING})
        for code in sorted(target_codes):
            new = next(n for n in nodes if n['新编码']==code)
            mapping.append({'原编码':old['code'],'原能力层级':old['level'],'原能力名称':old['name'],
                '优化动作':action,'新编码':code,'新能力层级':new['能力层级'],'新能力名称':new['能力名称'],
                '对应任务ID':new['对应任务ID'],'来源岗位ID':new['来源岗位ID'],
                '证据类型':new['证据类型'],'证据状态':new['证据状态'],'课程适配类别':new['课程适配类别'],
                '优化理由':reason,'人工审核状态':PENDING})
    return mapping, audits


def issues_list(jobs, duplicates, originals, audits, suggestions, tasks):
    issues = []
    def add(kind, priority, question, job='', task='', old='', new='', rid='', evidence=''):
        issues.append({'问题ID':f'Q-{len(issues)+1:04d}','问题类型':kind,'优先级':priority,
            '岗位ID':job,'任务ID':task,'原能力编码':old,'新能力编码':new,'原关系ID':rid,
            '证据文本或线索':evidence,'需人工判断的问题':question,'审核人':'','审核结论':'',
            '补证位置或链接':'','人工审核状态':PENDING})
    for j in jobs:
        add('招聘证据核验','高','打开原招聘来源，保存职责及要求逐字片段、访问时间、发布日期与招聘编号；当前摘要不能转为原文明示',
            job=j['样本ID'],evidence=j['原始招聘链接'])
    groups = defaultdict(list)
    for d in duplicates:
        if d['检查类型'] in {'同单位同岗位','单位简称线索+同岗位'}:
            groups[d['重复组']].append(d)
    for group in groups.values():
        add('重复及分类冲突','高','确认是否同一招聘事件、不同批次或转载；确认单位简称与分类，禁止自动合并',
            job=';'.join(d['岗位ID'] for d in group),evidence=group[0]['检查类型'])
    for a in audits:
        add('能力体系审核','中','审核名称、层级、行为指标及课程适配；'+a['优化理由'],
            old=a['原编码'],new=a['新编码集合'])
    for r in originals:
        if r['疑似误匹配']=='是':
            add('关系疑似误匹配','高' if not r['严格摘要支持新编码'] else '中',r['复核说明'],
                r['岗位ID'],r['任务编码'],r['三级能力编码'],r['严格摘要支持新编码'],r['原关系ID'],r['任务证据'])
    for t in tasks:
        if re.search('；|、|并|与|和',t['证据文本']):
            add('任务粒度复核','中','复合任务是否需按独立动作和结果细分；本轮保留原任务ID，未重新切分',
                t['岗位ID'],t['任务编码'],evidence=t['证据文本'])
    for s in suggestions:
        if s['能力层级']=='三级':
            add('课程建议隔离','高','暂无具体招聘证据。仅供课程讨论，不能计入招聘频次；需真实证据与课程标准后另行映射',new=s['新编码'])
    return issues


def write_excel(path, sheets):
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets:
        ws = wb.create_sheet(title)
        headers = list(rows[0]) if rows else ['说明']
        ws.append(headers)
        for r in rows:
            ws.append([r.get(h) for h in headers])
        for row in ws:
            for cell in row:
                if isinstance(cell.value,str):
                    if len(cell.value)>32767:
                        raise ValueError('Excel文本超过32767字符，请改为关系明细引用')
                    cell.data_type='s'
                cell.alignment=Alignment(wrap_text=True,vertical='top')
        for cell in ws[1]:
            cell.font=Font(bold=True,color='FFFFFF')
            cell.fill=PatternFill('solid',fgColor='245A81')
        ws.freeze_panes='A2'
        ws.auto_filter.ref=ws.dimensions
        for column in ws.columns:
            ws.column_dimensions[column[0].column_letter].width=min(55,max(15,len(str(column[0].value))*2))
    wb.save(path)
    wb.close()


def validate_output(nodes, suggestions, proposals, originals, jobs, tasks):
    ids = unique(nodes,'新编码')
    if len(set(ids)&{s['新编码'] for s in suggestions}):
        raise ValueError('课程建议与岗位能力编码冲突')
    js, ts = unique(jobs,'样本ID'),unique(tasks,'任务编码')
    for n in nodes:
        if n['父编码'] and n['父编码'] not in ids:
            raise ValueError('新能力父节点不存在')
    original_keys = {r['原关系ID'] for r in originals}
    if {r['原关系ID'] for r in proposals} != original_keys:
        raise ValueError('原关系有遗漏')
    for r in proposals:
        if r['新三级能力编码'] not in ids or r['任务ID'] not in ts or r['岗位ID'] not in js:
            raise ValueError('新关系外键不存在')
        if ts[r['任务ID']]['岗位ID'] != r['岗位ID'] or ts[r['任务ID']]['证据文本'] != r['任务证据']:
            raise ValueError('新关系证据不一致')
        if r['证据类型']=='招聘原文明示':
            raise ValueError('本批无已核验原文明示证据')
    for s in suggestions:
        if s['对应任务ID'] or s['来源岗位ID'] or s['摘要支持岗位数（含疑似重复）']:
            raise ValueError('课程建议混入招聘计数')


def report_text(stats, audits, originals, nodes):
    a = stats['原三级优化动作数量']
    examples = [r for r in originals if r['疑似误匹配']=='是']
    preferred = []
    for code in ['C131','C521','C412','C332','C211']:
        row = next((r for r in examples if r['三级能力编码']==code),None)
        if row:
            preferred.append(f"- {row['原关系ID']}／{row['岗位ID']}／{row['任务编码']}：原 {code}，摘要“{row['任务证据']}”。{row['复核说明']}。")
    mapping_table = '\n'.join(f"| {x['原编码']} | {x['原能力名称']} | {x['优化动作']} | {x['新编码集合']} |" for x in audits if x['原能力层级']=='三级')
    count_table = '\n'.join(f'| {k} | {v} |' for k,v in stats['实际输入数量'].items())
    return f'''# 第二阶段优化报告

## 一、研究对象与边界

研究对象为高职大数据技术专业《数据库原理及应用》单门课程。本阶段仅审核旅游行业岗位职业能力体系，不构建课程知识图谱，不改写招聘样本和第一阶段成果。本报告由AI辅助研究形成，全部优化建议待人工审核，未代替专家审定。

输入以已合并PR #1的main快照 `{stats['输入main提交']}` 为基准。开始处理前已逐文件比对Git blob；本次输入快照差异数为 {stats['输入快照差异数']}。输入文件哈希及预期快照见“运行统计.json”和 phase2/input_snapshot.json。处理日期为2026-10-22（北京时间），不是招聘网页访问日期。

## 二、实际文件与数量核对

| 实际读取项目 | 数量 |
| --- | ---: |
{count_table}

与预期50岗位、79任务、5/15/26层级能力及188关系的数量差异：{stats['数量核对结论']}。一级、二级由实际能力表与聚合表独立核对；没有复用说明文档中的数字代替读取。原有全部岗位及任务ID保留，第一阶段文件不输出到原目录。

50条岗位仍全部待核验；主库“证据摘录”为空 {stats['原文摘录缺失岗位数']} 条。本阶段未联网逐条核验招聘页。原证据质量初评A/B/C及状态只是原表元数据，不构成原文明示证明。

## 三、能力审核与优化结果

原三级能力按互斥口径分类：保留并规范 {a.get('保留并规范',0)} 项；合并涉及 {a.get('合并',0)} 项原能力、形成 {stats['合并目标三级数']} 项新能力；拆分涉及 {a.get('拆分',0)} 项原能力、形成 {stats['拆分目标三级数']} 项新能力。岗位证据驱动新增三级能力为0；另列课程建议三级能力 {stats['课程建议三级数']} 项，均无具体招聘证据，不计入招聘需求频次。这些数量按原三级项/新三级项分别统计，不能相加。

优化岗位体系为 {stats['优化一级数']} 项一级、{stats['优化二级数']} 项二级、{stats['优化三级数']} 项三级，共 {stats['优化全层级节点数']} 个节点；另有课程建议 {stats['课程建议全层级节点数']} 个节点（含上位分组）。原有 {stats['原体系待审核节点数']} 个节点全部完成逐项研究审核并给出新旧映射；优化岗位节点 {stats['优化全层级节点数']} 个、课程建议节点 {stats['课程建议全层级节点数']} 个均待人工审核，批准数为0。

“保留”包含名称和评价行为规范，不表示已经通过专家审核。新编码为N前缀，课程建议为S前缀。所有原编码包括5项一级和15项二级保存在对照表中；拆分允许一对多、合并允许多对一，不覆盖旧编码。

代表性调整：

- C113“提取查询业务数据”与C131“使用SQL处理和查询数据”合并为N113“查询并提取业务数据”。SQL作为逐关系证据限定单独记录；有查询任务不必然有SQL，有任职摘要SQL不必然证明每项分析任务使用SQL。缺少具体查询动作的旧关系暂缓迁移至摘要支持频次。
- C121拆成N121“维护指标口径与数据字典”和N122“核对与校验业务数据”；同一任务只有明确摘要支持时才进入相应分项，不把数据链路梳理直接当成质量校验。
- C331拆成N321“执行业务数据备份”和N322“恢复业务数据”，不能从备份自动推出恢复需求。C332拆成N331访问权限和N332安全检查；普通系统权限不等同于数据库授权操作。
- C211拆成指标构建与指标统计监测；C231拆成分析模型与预测模型。数据模型不自动等于关系数据库模型，统计也不自动等于SQL GROUP BY或统计检验。
- C521拆成技术文档与数据处理流程，消除“流程”触发“文档能力”的对象混淆。C232建议形成与C233增长执行予以区分保留，未因主题相似强行合并。
- 数据库查询性能由原C4系统建设调整至N3运行保障的N341；上位按工作领域归并，三级按动作、对象和可检查结果描述。少数单叶二级节点保留工作目的边界，不为凑层级虚增能力。

| 原三级编码 | 原名称 | 优化动作 | 新编码 |
| --- | --- | --- | --- |
{mapping_table}

## 四、数据库重点能力与课程适配

A表示数据库课程核心相关，B表示数据库课程延伸相关，C表示旅游行业背景相关。分类均为研究判断，不表示课程标准已经覆盖。三级职业能力分类数量为 {json.dumps(stats['岗位三级课程适配数量'],ensure_ascii=False)}；混合子项的上位节点暂归B，应以三级逐项判断为准。

- SQL查询：归并查询能力并显式标注SQL来自职责摘要、任职摘要还是尚未明示，不增加连接、子查询等未获具体证据的细项。
- 数据库设计和数据表管理：S111关系数据库结构设计、S112建表改表均为课程建议，不能由“建模”“数据资源管理”替代其招聘证据。
- 数据更新：N123只承载业务数据维护更新的摘要支持。S113明确SQL增删改是课程建议；原“数据库更新”不能证明INSERT/UPDATE/DELETE语句要求。
- 统计分析：指标统计、业务解释、报表、看板和建模分开描述，不补写特定统计方法。
- 数据库运行维护、备份恢复和权限安全：分开动作并保留对象限定。A类体现可用于数据库教学的研究判断；任务对象未明确为数据库时，不宣称具体数据库技术需求。

行为表述中的结果记录、检查方式为研究者提出的可观察评价描述，不是招聘原文。课程建议不关联虚构岗位或任务，所有来源ID留空，摘要支持数和已核验原文岗位数均为0。

## 五、任务匹配复核与证据层级

原{stats['实际输入数量']['关系']}条关系逐条保留，输出“原关系ID”便于追踪；按新旧映射展开为 {stats['迁移判定行数']} 条迁移判定，其中 {stats['摘要支持判定行数']} 条为岗位摘要支持，{stats['仅推断待判行数']} 条仅推断待人工判断，{stats['不建议迁移行数']} 条因拆分项无支持而不建议迁移。展开行数不是新增任务数，也不是招聘出现次数。

原关系中 {stats['疑似误匹配原关系数']} 条进入疑似误匹配清单。清单包含宽泛词触发、SQL工具推定、备份权限对象不清、流程/文档混淆等情形；“疑似”并不等于已判定错误。宽泛词即使同时存在具体对象支持，也需要人工核对，不能一概删除。

{chr(10).join(preferred)}

证据分为三层：

1. 招聘原文明示：须有可核对的招聘逐字原文与核验记录，本批为0。原有空摘录和历史状态不能补足这一层。
2. 岗位摘要支持：原职责摘要出现新能力的具体动作或对象，引用原任务文本；仍属候选证据，待回看原文。
3. 研究推断：原宽泛匹配、任职工具到具体任务的连接、能力组织及评价行为的研究判断。与摘要支持分列，不能当成原文明示。

“原要求性质”（如摘要明确/优先）仅透传第一阶段判断，不能作为本阶段招聘原文证据等级。所有任务到能力关系仍是研究者构建的联系，摘要支持只说明摘要与候选行为有对应，不意味着招聘方使用这些能力术语。

只对“摘要支持（待审）”计算候选岗位ID去重数和任务ID去重数；推断及不建议迁移行不计入摘要支持。合并后同一岗位或任务只计一次。计数含尚未解决的疑似重复，不是独立招聘数量；不计算已核验正式需求频次，所有能力“已核验招聘原文岗位数”为0。上位节点也使用ID并集，不能把叶节点频次相加。

## 六、人工审核与交付

人工问题清单共 {stats['人工问题条数']} 条，按问题类型分别为 {json.dumps(stats['人工问题类型数量'],ensure_ascii=False)}。清单包含50岗位逐字原文核验、原有同岗异源及分类冲突、全部原节点审核、疑似关系、复合任务粒度及课程建议隔离。审核人、审核结论、补证链接留空，未模拟人工确认。

交付文件：职业能力体系_优化对照表.xlsx、三级职业能力体系_优化建议版.xlsx、任务能力关系_复核表.xlsx、人工审核问题清单.xlsx。文件分别保留全层级映射、具体可观察行为、原关系及迁移判定、问题和可填写审核字段。Python脚本和规则文件位于phase2/；运行方法及测试命令见phase2/README.md。

优先审核数据库设计、DDL、DML建议是否应进入后续课程建设讨论，并补齐其真实招聘证据；确认同岗异源是否重复；逐条处理高优先级匹配疑问；再审定三级粒度、合并与拆分边界。本阶段未批准最终能力体系，也未建立课程教学项目、知识点或最终课程图谱。

## 七、复现与验证

在仓库根目录执行 `python phase2/review.py`，再执行 `python -m unittest discover -s phase2 -p 'test_*.py' -v`。脚本离线读取，默认独立输出，运行前后验证输入哈希不变。测试覆盖四表可打开、原始文件保护、原编码及关系保留、任务岗位溯源、层级与必需字段、合并去重、上位并集、SQL及备份恢复边界、课程建议零频次和重复运行一致性。本次实际测试结果另存“测试结果.txt”；测试不证明招聘真实性、语义规则完备或专家效度。
'''


def run(input_root=ROOT, output_dir=None, rules_path=None):
    root = Path(input_root).resolve()
    out = Path(output_dir or ROOT/'phase2/output').resolve()
    # 不允许覆盖任何第一阶段目录或原始文件。
    if out == root or out == root/'output' or root/'output' in out.parents:
        raise ValueError('禁止覆盖第一阶段文件：请选择独立第二阶段目录')
    if any(out == p or p in out.parents for base in {root, ROOT} for p in [base/'output',base/'scripts',base/'tests']):
        raise ValueError('禁止写入第一阶段代码目录')
    rules_path = Path(rules_path or ROOT/'phase2/rules.json')
    rules = json.loads(rules_path.read_text(encoding='utf-8'))
    snapshot = json.loads((ROOT/'phase2/input_snapshot.json').read_text(encoding='utf-8'))
    before = {p:hashlib.sha256((root/p).read_bytes()).hexdigest() for p in INPUTS}
    jobs,tasks,abilities,relations,duplicates = load_inputs(root)
    validate_rules(rules,abilities)
    originals,proposals = relation_review(jobs,tasks,relations,rules)
    nodes,suggestions = build_nodes(rules,proposals)
    mapping,audits = compare(abilities,rules,nodes)
    for s in suggestions:
        mapping.append({'原编码':'','原能力层级':'','原能力名称':'','优化动作':'新增课程建议',
                        '新编码':s['新编码'],'新能力层级':s['能力层级'],'新能力名称':s['能力名称'],
                        '对应任务ID':'','来源岗位ID':'','证据类型':s['证据类型'],'证据状态':s['证据状态'],
                        '课程适配类别':s['课程适配类别'],'优化理由':s['优化理由'],'人工审核状态':PENDING})
    issues = issues_list(jobs,duplicates,originals,audits,suggestions,tasks)
    validate_output(nodes,suggestions,proposals,originals,jobs,tasks)
    # 独立读取第一阶段聚合页，避免单纯由三级表推导掩盖实际聚合表差异。
    aggregates = read_sheet(root/INPUTS[2], '一级二级聚合', ['层级','编码','能力名称'])
    actual = {'岗位':len(jobs),'任务':len(tasks),'一级':len({a['一级编码'] for a in abilities}),
              '二级':len({a['二级编码'] for a in abilities}),'三级':len(abilities),'关系':len(relations),
              '聚合页一级':sum(a['层级']=='一级' for a in aggregates),'聚合页二级':sum(a['层级']=='二级' for a in aggregates)}
    for level in ['一级','二级']:
        expected = {(a[level+'编码'],a[level+'能力']) for a in abilities}
        observed = {(a['编码'],a['能力名称']) for a in aggregates if a['层级']==level}
        if expected != observed:
            raise ValueError('实际聚合页与三级表层级不一致：'+level)
    expected = {'岗位':50,'任务':79,'一级':5,'二级':15,'三级':26,'关系':188,'聚合页一级':5,'聚合页二级':15}
    mismatches = {k:{'预期':expected[k],'实际':actual[k]} for k in expected if expected[k]!=actual[k]}
    actions = Counter(a['优化动作'] for a in audits if a['原能力层级']=='三级')
    split_old = {a['原编码'] for a in audits if a['优化动作']=='拆分'}
    evidence_counts = Counter(r['关系建议'] for r in proposals)
    expected_blobs = {x['path']:x['sha'] for x in snapshot['files']}
    drift = []
    for p in INPUTS:
        content = (root/p).read_bytes()
        blob = hashlib.sha1(b'blob '+str(len(content)).encode()+b'\0'+content).hexdigest()
        if expected_blobs.get(p) != blob:
            drift.append(p)
    stats = {'输入main提交':snapshot['main_commit'],'实际输入数量':actual,'数量核对结论':mismatches or '全部一致',
             '输入SHA256':before,'规则SHA256':hashlib.sha256(rules_path.read_bytes()).hexdigest(),
             '审核脚本SHA256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
             '输入快照差异数':len(drift),'输入快照差异文件':drift,
             '原文摘录缺失岗位数':sum(not j['证据摘录'] for j in jobs),
             '原三级优化动作数量':dict(actions),'合并目标三级数':sum(len(s['old'])>1 for s in rules['leaves']),
             '拆分目标三级数':sum(bool(set(s['old'])&split_old) for s in rules['leaves']),
             '优化一级数':sum(n['能力层级']=='一级' for n in nodes), '优化二级数':sum(n['能力层级']=='二级' for n in nodes),
             '优化三级数':sum(n['能力层级']=='三级' for n in nodes),'优化全层级节点数':len(nodes),
             '课程建议三级数':len(rules['course_suggestions']),'课程建议全层级节点数':len(suggestions),
             '原体系待审核节点数':len(audits),'迁移判定行数':len(proposals),
             '摘要支持判定行数':evidence_counts['摘要支持（待审）'], '仅推断待判行数':evidence_counts['待人工判断（仅推断）'],
             '不建议迁移行数':evidence_counts['不建议迁移（拆分项无支持）'],
             '疑似误匹配原关系数':sum(r['疑似误匹配']=='是' for r in originals),
             '岗位三级课程适配数量':dict(Counter(n['课程适配类别'] for n in nodes if n['能力层级']=='三级')),
             '人工问题条数':len(issues),'人工问题类型数量':dict(Counter(i['问题类型'] for i in issues)),
             '招聘原文明示关系数':0,'已人工批准能力数':0,'岗位状态':'全部待核验','边界':BOUNDARY}
    notes = [{'项目':'研究边界','说明':BOUNDARY},
             {'项目':'证据类型','说明':'招聘原文明示本批0；岗位摘要支持与研究推断分开；原明确/优先只表示第一阶段摘要判断'},
             {'项目':'迁移判定','说明':'只在原关系及新旧映射范围复核；未自动穷举所有可能新关系；不建议迁移行不删除原关系'},
             {'项目':'频次口径','说明':'仅摘要支持（待审）按岗位和任务ID去重；含疑似重复；不是已核验需求频次。推断与课程建议不计数'},
             {'项目':'课程类别','说明':'A核心相关、B延伸相关、C行业背景；全部研究判断，非课程标准覆盖证明；与原证据质量初评A/B/C是两套分类，不能互换'},
             {'项目':'人工状态','说明':'所有节点和关系待人工审核；问题表空白审核字段用于人工填写，重运行会覆盖第二阶段同名输出，请先备份人工填写版本'}]
    coverage = [{'任务ID':t['任务编码'],'岗位ID':t['岗位ID'],'任务证据':t['证据文本'],
                 '摘要支持新能力编码':';'.join(sorted({r['新三级能力编码'] for r in proposals if r['任务ID']==t['任务编码'] and r['证据类型']=='岗位摘要支持'})),
                 '原文核验状态':'待核验','人工审核状态':PENDING} for t in tasks]
    # 输入在生成前后检查不变，所有输出仅写入独立目录。
    out.mkdir(parents=True,exist_ok=True)
    write_excel(out/'职业能力体系_优化对照表.xlsx', [('新旧编码映射',mapping),('原体系逐项审核',audits),('方法与边界',notes)])
    write_excel(out/'三级职业能力体系_优化建议版.xlsx', [('优化岗位能力全层级',nodes),('三级岗位能力',[n for n in nodes if n['能力层级']=='三级']),('课程建议能力单列',suggestions),('方法与边界',notes)])
    write_excel(out/'任务能力关系_复核表.xlsx', [('原关系逐条复核',originals),('新旧关系迁移判定',proposals),('原任务覆盖',coverage),('方法与边界',notes)])
    write_excel(out/'人工审核问题清单.xlsx', [('人工审核问题',issues),('疑似关系明细',[r for r in originals if r['疑似误匹配']=='是']),('方法与边界',notes)])
    if before != {p:hashlib.sha256((root/p).read_bytes()).hexdigest() for p in INPUTS}:
        raise RuntimeError('输入文件发生变化，请停止并检查')
    (out/'运行统计.json').write_text(json.dumps(stats,ensure_ascii=False,indent=2),encoding='utf-8')
    (out/'第二阶段优化报告.md').write_text(report_text(stats,audits,originals,nodes),encoding='utf-8')
    return stats


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input-root',type=Path,default=ROOT)
    parser.add_argument('--output',type=Path,default=ROOT/'phase2/output')
    parser.add_argument('--rules',type=Path,default=ROOT/'phase2/rules.json')
    args=parser.parse_args()
    print(json.dumps(run(args.input_root,args.output,args.rules),ensure_ascii=False,indent=2))
