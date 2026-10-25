"""候选能力图谱只读访问，不调用任何研究生成器。"""
import json
from pathlib import Path
from collections import Counter
class GraphDataError(ValueError):pass
class GraphStore:
 def __init__(self,root):
  self.directory=Path(root)/'stage4_competency_graph/output'
  try:
   def read(name):
    data=json.loads((self.directory/name).read_text(encoding='utf-8'))
    if not isinstance(data,list):raise GraphDataError(name+'必须是数组')
    return data
   self.nodes=read('graph_nodes.json');self.edges=read('graph_edges.json');self.evidence=read('graph_evidence.json');self.reviews=read('graph_review_items.json');self.paths=read('graph_paths.json')
   self.by_id=self.index(self.nodes,'node_id');self.edge_by_id=self.index(self.edges,'edge_id');self.evidence_by_id=self.index(self.evidence,'evidence_id')
   for n in self.nodes:
    for key in ['original_id','node_type','name','source_file','review_status']:n[key]
   for e in self.edges:
    for key in ['relation_code','relation_type','reason','review_status','source_nature','support_function_or_subtype','source_attributes']:e[key]
    if type(e['default_visible']) is not bool:raise GraphDataError('default_visible必须为布尔值')
    if e['source'] not in self.by_id or e['target'] not in self.by_id:raise GraphDataError('关系端点不存在：'+e['edge_id'])
    if any(id not in self.evidence_by_id for id in e['evidence_ids']):raise GraphDataError('证据引用不存在：'+e['edge_id'])
   for p in self.paths:
    if any(id not in self.edge_by_id for id in p['edge_path']):raise GraphDataError('路径引用无效')
  except (OSError,ValueError,KeyError,TypeError) as exc:raise GraphDataError('图谱数据读取失败：'+str(exc)) from exc
 @staticmethod
 def index(rows,key):
  result={r[key]:r for r in rows}
  if len(result)!=len(rows):raise GraphDataError('编号重复：'+key)
  return result
 def search(self,query='',types=None):return [n for n in self.nodes if (not types or n['node_type'] in types) and query.casefold() in (n['original_id']+' '+n['name']).casefold()]
 def filtered(self,explore=False,codes=None,functions=None,statuses=None):return [e for e in self.edges if (explore or e['default_visible']) and (not codes or e['relation_code'] in codes) and (not functions or e['relation_code']!='R03' or e['support_function_or_subtype'] in functions) and (not statuses or e['review_status'] in statuses)]
 def related(self,id,explore=False):return [e for e in self.filtered(explore) if id in [e['source'],e['target']]]
 def parents(self,id):
  result=[]
  while True:
   rs=[e for e in self.edges if e['relation_code']=='R01' and e['target']==id]
   if not rs:return result
   id=rs[0]['source']
   if id in result:raise GraphDataError('能力层级存在循环')
   result.append(id)
 def subgraph(self,center,edges,types=None,depth=1):
  ids={center};selected={}
  for _ in range(depth):
   frontier=set(ids)
   for e in edges:
    if e['source'] in frontier or e['target'] in frontier:
     a,b=self.by_id[e['source']],self.by_id[e['target']]
     if types is not None and (a['node_type'] not in types or b['node_type'] not in types):continue
     selected[e['edge_id']]=e;ids.update([e['source'],e['target']])
  return [n for n in self.nodes if n['node_id'] in ids and (types is None or n['node_type'] in types)],list(selected.values())
 def teaching_paths(self,id,explore=False):return [p for p in self.paths if p['competency_id']==id and (explore or p['default_visible'])]
 def review_for(self,id):
  original=self.by_id[id]['original_id'];edges={e['edge_id'] for e in self.edges if id in [e['source'],e['target']]}
  return [q for q in self.reviews if q['object_id'] in edges or original in str(q['object_id'])]
 def path_to(self,center,target,edges):
  """无向遍历仅用于解释邻接，保留每条实际有向边，不推断层级。"""
  from collections import deque
  queue=deque([(center,[])]);seen={center}
  while queue:
   id,path=queue.popleft()
   if id==target:return path
   for e in edges:
    other=e['target'] if e['source']==id else e['source'] if e['target']==id else None
    if other and other not in seen:seen.add(other);queue.append((other,path+[e['edge_id']]))
  return None
