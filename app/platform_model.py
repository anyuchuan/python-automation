"""第五阶段只读视图模型：只索引已有层级、任务、匹配与证据，不生成研究关系。"""
from collections import Counter,defaultdict,deque
from app.data import DataError,Options,visible_data,edge_risk,detail

MODULES=('课程能力图谱','岗位能力需求分析','能力—教学内容映射','课程能力支撑评价','课程教学优化','研究依据与数据管理')
LEVELS=('一级','二级','三级')

def split_ids(value):return [x for x in str(value or '').split(';') if x]

def hierarchy(bundle):
    items={n['节点ID']:{'id':n['节点ID'],'name':n['节点名称'],'parent':bundle.attrs[n['节点ID']].get('父编码'),
                           'level':bundle.attrs[n['节点ID']].get('能力层级')} for n in bundle.nodes if n['节点类型']=='职业能力建议'}
    if Counter(n['level'] for n in items.values())!=Counter({'一级':5,'二级':14,'三级':34}):raise DataError('能力树不是5/14/34层级')
    pairs=set()
    for n in items.values():
        parent=n['parent'];level=LEVELS.index(n['level'])
        if level==0:
            if parent:raise DataError('一级能力不应有父编码')
        elif parent not in items or LEVELS.index(items[parent]['level'])!=level-1:raise DataError('能力父编码或层级无效')
        else:pairs.add((parent,n['id']))
    actual={(e['起点ID'],e['终点ID']) for e in bundle.edges if e['关系类型']=='建议能力层级包含' and e['起点ID'] in items and e['终点ID'] in items}
    if pairs!=actual or len(pairs)!=48:raise DataError('能力属性与源层级关系不一致')
    return dict(sorted(items.items()))

def ancestors(tree,id):
    out=[]
    while id in tree and tree[id]['parent']:
        id=tree[id]['parent'];out.append(id)
    return list(reversed(out))

def visible_tree(tree,expanded):
    return [id for id in tree if all(p in expanded for p in ancestors(tree,id))]

def find_capabilities(tree,query):
    q=query.strip().casefold()
    return [id for id,n in tree.items() if not q or q in (id+' '+n['name']).casefold()]

def global_counts(bundle):
    count=Counter(n['节点类型'] for n in bundle.nodes)
    return {'候选岗位':count['岗位候选样本'],'典型任务':count['典型工作任务（摘要）'],'三级能力':len(bundle.capabilities),
            '教学项目':len(bundle.project_by_id),'原始节点':len(bundle.nodes),'候选关系':len(bundle.edges),'隔离关系':len(bundle.isolated)}

def evidence_group(edge):
    if edge.get('展示分组')=='隔离关系':return '隔离关系（未确认）'
    if edge.get('展示风险说明') or '研究推断' in str(edge.get('招聘证据类型') or ''):return '研究推断/疑似（未确认）'
    if str(edge.get('关系类型') or '').startswith('能力候选'):return '课程文本匹配（待审核）'
    if '摘要' in str(edge.get('关系类型') or '') or edge.get('招聘证据类型'):return '岗位摘要（未核验）'
    return '结构或归档（待审核）'

def network_slice(bundle,center,options=Options(),types=None,relations=None,evidence=None,hops=2,limit=70):
    nodes,edges=visible_data(bundle,options)
    nodes=[n for n in nodes if types is None or n['节点类型'] in types];ids={n['节点ID'] for n in nodes}
    edges=[e for e in edges if e['起点ID'] in ids and e['终点ID'] in ids and (relations is None or e['关系类型'] in relations) and (evidence is None or evidence_group(e) in evidence)]
    if center not in ids:return [],[],0
    adj=defaultdict(set)
    for e in edges:adj[e['起点ID']].add(e['终点ID']);adj[e['终点ID']].add(e['起点ID'])
    distance={center:0};queue=deque([center])
    while queue:
        id=queue.popleft()
        if distance[id]>=hops:continue
        for next_id in sorted(adj[id]):
            if next_id not in distance:distance[next_id]=distance[id]+1;queue.append(next_id)
    ordered=sorted(distance,key=lambda id:(distance[id],id));total=len(ordered);chosen=set(ordered[:max(1,limit)])
    return [bundle.by_id[id] for id in ordered if id in chosen],[e for e in edges if e['起点ID'] in chosen and e['终点ID'] in chosen],total

def capability_detail(bundle,id):
    if id not in bundle.by_id or bundle.by_id[id]['节点类型']!='职业能力建议':return None
    a=bundle.attrs[id];c=bundle.cap_by_id.get(id,{})
    tasks=[t for t in split_ids(a.get('对应任务ID')) if t in bundle.by_id]
    jobs=[j for j in split_ids(a.get('来源岗位ID')) if j in bundle.by_id]
    cells=[r for r in bundle.project_cells if r['能力ID']==id and r['项目层匹配等级'] in {'直接支撑','部分支撑','待人工审核'}]
    return {'id':id,'name':bundle.by_id[id]['节点名称'],'level':a.get('能力层级') or '暂无记录','parent':a.get('父编码') or '暂无记录',
            'definition':a.get('能力定义') or '暂无记录','behavior':a.get('可观察的工作行为') or '暂无记录',
            'evidence':a.get('证据类型') or '暂无记录','state':a.get('证据状态') or '暂无记录','review':a.get('人工审核状态') or '暂无记录',
            'grade':c.get('最终候选匹配状态') or '暂无记录（上位能力不外推支撑等级）',
            'technical':c.get('课程支撑的具体技术行为') or '暂无记录','uncovered':c.get('课程尚未明确覆盖的工作行为') or '暂无记录',
            'tasks':tasks,'jobs':jobs,'project_cells':cells,'source':bundle.by_id[id]['源文件或条目']}

def project_cell(bundle,cap,project):
    return next((r for r in bundle.project_cells if r['能力ID']==cap and r['课程项目ID']==project),None)

def cell_evidence(bundle,cell):
    # 单元明列的证据ID优先；不借用其他项目的引文或把建议补成正向证据。
    return [bundle.evidence_by_id[eid] for eid in split_ids(cell.get('正向课程证据ID'))+split_ids(cell.get('待审课程提及证据ID')) if eid in bundle.evidence_by_id]

def optimization_rows(bundle):
    return [{'能力ID':r['能力ID'],'能力名称':r['能力名称'],'能力层面文本判断':r['最终候选匹配状态'],
             '研究发现':r.get('课程标准支撑不足或边界') or '暂无记录',
             '建议分类':r.get('未见支撑建议分类') or '暂无记录',
             '建议依据（研究判断）':r.get('分类理由') or '暂无记录',
             '待核实事项':r.get('课程尚未明确覆盖的工作行为') or '暂无记录',
             '来源':'课程能力覆盖与缺口分析.xlsx/逐项支撑与缺口；职业能力课程匹配矩阵.xlsx/34项能力汇总',
             '审核状态':r.get('人工审核状态') or '暂无记录'} for r in bundle.capabilities]
