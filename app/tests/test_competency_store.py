import unittest,json,hashlib,tempfile,shutil
from pathlib import Path
from app.competency_store import GraphStore,GraphDataError
ROOT=Path(__file__).resolve().parents[2]
class CandidateGraphTests(unittest.TestCase):
 def setUp(self):self.g=GraphStore(ROOT)
 def test_source_counts_and_queries(self):
  g=self.g;p=g.directory
  self.assertEqual(len(g.nodes),len(json.loads((p/'graph_nodes.json').read_text())))
  self.assertEqual(len(g.edges),len(json.loads((p/'graph_edges.json').read_text())))
  self.assertEqual(len(g.search(types=['COMPETENCY_L3'])),17);self.assertEqual(len(g.search(types=['COURSE_KS'])),60)
  for i in range(1,18):self.assertTrue(g.search(f'DB{i:03}',['COMPETENCY_L3']))
 def test_inherited_data(self):
  from collections import Counter
  g=self.g;self.assertEqual(sum(e['relation_code']=='R01' for e in g.edges),26);self.assertEqual(Counter(e['support_function_or_subtype'] for e in g.edges if e['relation_code']=='R02'),{'A':11,'B':27,'C':27});self.assertEqual(sum(e['relation_code']=='R03' for e in g.edges),34)
 def test_default_and_exploration(self):
  self.assertEqual(self.g.filtered(),[e for e in self.g.edges if e['default_visible']]);self.assertEqual(len(self.g.filtered(True)),len(self.g.edges))
  self.assertTrue(all(e['source_attributes'].get('审查建议') not in ['不保留','待审'] for e in self.g.filtered() if e['relation_code']=='R09'))
 def test_boundaries(self):
  e=self.g.edge_by_id['R03:034'];self.assertEqual(e['support_function_or_subtype'],'K4');self.assertIn('不要求',e['uncertainty'])
  e=self.g.edge_by_id['R03:030'];self.assertIn('已有质量检查条件',e['uncertainty'])
  for id in ['DB009','DB010','DB014','DB015']:self.assertFalse(any(e['relation_code']=='R03' for e in self.g.related('COMPETENCY_L3:'+id)))
 def test_all_evidence_and_endpoints(self):
  for e in self.g.edges:
   self.assertIn(e['source'],self.g.by_id);self.assertIn(e['target'],self.g.by_id)
   for id in e['evidence_ids']:self.assertIn(id,self.g.evidence_by_id)
 def test_filter_local_paths_and_unlinked(self):
  g=self.g;es=g.filtered(codes=['R03'],functions=['K4']);self.assertTrue(all(e['support_function_or_subtype']=='K4' for e in es))
  ns,edges=g.subgraph('COMPETENCY_L3:DB017',g.filtered());self.assertTrue(edges);self.assertTrue(g.teaching_paths('COMPETENCY_L3:DB017'))
  self.assertTrue(any(not any(e['relation_code']=='R03' for e in g.related(n['node_id'])) for n in g.search(types=['COURSE_KS'])))
 def test_read_only_and_bad_data(self):
  before={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in self.g.directory.glob('*.json')};self.g.related('COMPETENCY_L3:DB017',True);self.assertEqual(before,{p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in self.g.directory.glob('*.json')})
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'stage4_competency_graph/output';shutil.copytree(self.g.directory,p);(p/'graph_nodes.json').write_text('bad')
   with self.assertRaises(GraphDataError):GraphStore(t)
 def test_page_smoke(self):
  from streamlit.testing.v1 import AppTest
  at=AppTest.from_file(str(ROOT/'app/main.py'),default_timeout=30).run();at.radio(key='platform_entry').set_value('研究审查图谱').run();self.assertFalse(at.exception)
  at.radio(key='cg_view').set_value('知识技能反向查询').run();self.assertFalse(at.exception)
  at.radio(key='cg_view').set_value('教学关联路径').run();self.assertFalse(at.exception)
  at.radio(key='cg_view').set_value('局部关系网络').run();self.assertFalse(at.exception)
 def test_acceptance_paths_and_presentation(self):
  from app.competency_display import directed,relation_fields,NODE_LABELS,REL_LABELS
  g=self.g;center='COMPETENCY_L3:DB001'
  for id in ['DB005','DB017']:
   chain=g.path_to(center,'COMPETENCY_L3:'+id,g.filtered());self.assertTrue(chain)
   self.assertTrue(all(g.edge_by_id[x]['relation_code']!='R01' for x in chain))
  for e in g.edges:
   fields=relation_fields(g,e);self.assertIn(g.by_id[e['source']]['name'],fields['起点']);self.assertIn(g.by_id[e['target']]['name'],fields['终点']);self.assertIn(e['relation_code'],REL_LABELS)
   if e['relation_code']=='R09' and e['support_function_or_subtype']=='组合关系':self.assertFalse(directed(e))
  self.assertEqual(set(NODE_LABELS),{n['node_type'] for n in g.nodes})
  self.assertEqual(g.subgraph(center,g.filtered(),[],1),([],[]))

if __name__=='__main__':unittest.main()
