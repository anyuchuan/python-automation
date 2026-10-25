"""探索视图只选取真实节点/边，不创建研究实体或关系。"""
def overview(g):
 nodes=[n for n in g.nodes if n['node_type'].startswith('COMPETENCY')];ids={n['node_id'] for n in nodes};return nodes,[e for e in g.filtered() if e['relation_code']=='R01' and e['source'] in ids and e['target'] in ids]
def topic(g,cap,kind):
 code='R02' if kind=='岗位关联' else 'R03';edges=[e for e in g.filtered() if e['relation_code']==code and e['target']==cap];ids={cap}|{e['source'] for e in edges};return [n for n in g.nodes if n['node_id'] in ids],edges
def curriculum(g,project=None):
 if project is None:return [n for n in g.nodes if n['node_type']=='TEACHING_PROJECT'],[]
 edges=[e for e in g.filtered() if e['relation_code']=='R04' and e['target']==project];ids={project}|{e['source'] for e in edges};return [n for n in g.nodes if n['node_id'] in ids],edges
def highlight(g,id,nodes,edges):
 ids={id};changed=True
 while changed:
  before=len(ids)
  for e in edges:
   if e['relation_code']=='R01' and e['source'] in ids:ids.add(e['target'])
  changed=len(ids)!=before
 ids.update(g.parents(id) if g.by_id[id]['node_type'].startswith('COMPETENCY') else [])
 return ids
