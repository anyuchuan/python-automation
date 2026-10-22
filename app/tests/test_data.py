import copy
import csv
import io
import unittest
from collections import Counter
from app.data import *
from app.graph import build_html

class ResearchDataTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls): cls.b=load_bundle()
 def test_source_counts_and_revised_version(self):
  b=self.b
  self.assertEqual((len(b.nodes),len(b.edges),len(b.isolated)),(314,864,60))
  self.assertEqual(Counter(c['最终候选匹配状态'] for c in b.capabilities),EXPECTED_COUNTS)
  self.assertEqual(len(b.project_cells),272)
 def test_unique_ids_and_valid_endpoints(self):
  b=self.b; ids=set(b.by_id)
  self.assertEqual(len(ids),len(b.nodes))
  self.assertTrue(all(e['起点ID'] in ids and e['终点ID'] in ids for e in b.edges+b.isolated))
 def test_reject_old_five_direct_version(self):
  b=self.b;c=copy.deepcopy(b.capabilities)
  for row in [r for r in c if r['最终候选匹配状态']=='部分支撑'][:4]:row['最终候选匹配状态']='直接支撑'
  with self.assertRaisesRegex(DataError,'禁止回退'):validate_records(b.nodes,b.edges,b.isolated,c,b.coverage,b.project_cells)
 def test_default_hidden_and_explicit_two_step_gate(self):
  b=self.b;n,e=visible_data(b)
  self.assertEqual((len(n),len(e)),(259,450))
  self.assertFalse(any(x['展示风险说明'] for x in e))
  self.assertEqual(visible_data(b,Options(False,True,True)),(n,e))
  alln,alle=visible_data(b,Options(True,True,True))
  self.assertEqual((len(alln),len(alle)),(314,924))
  self.assertEqual(sum(x['展示分组']=='隔离关系' for x in alle),60)
 def test_filter_types_levels_grades(self):
  b=self.b;n,e,_=filter_graph(b,types=['职业能力建议'],levels=['三级'],statuses=['直接支撑'],max_nodes=1000)
  self.assertEqual(len(n),1);self.assertEqual(n[0]['节点ID'],'N113');self.assertEqual(e,[])
  self.assertEqual(filter_graph(b,types=[])[0],[])
  self.assertEqual(filter_graph(b,statuses=[])[0],[])
 def test_local_subgraph_and_limits(self):
  n,e,m=filter_graph(self.b,types=MAIN_TYPES,levels=['三级'],center='N113',hops=2)
  self.assertEqual((len(n),len(e)),(20,39));ids={x['节点ID'] for x in n}
  self.assertTrue(all(x['起点ID'] in ids and x['终点ID'] in ids for x in e))
  self.assertEqual(ids&set(self.b.cap_by_id),{'N113'})
  n,e,m=filter_graph(self.b,max_nodes=5);self.assertEqual(len(n),5);self.assertTrue(m['已截断'])
 def test_search_project_and_empty(self):
  b=self.b
  self.assertTrue(search_nodes(b.nodes,'N113'))
  self.assertTrue(search_nodes(b.nodes,'SQL'))
  self.assertEqual(search_nodes(b.nodes,'不存在的研究节点XYZ'),[])
  project=next(iter(b.project_by_id));n,e,_=filter_graph(b,project=project,max_nodes=1000)
  self.assertIn(project,{r['节点ID'] for r in n})
  self.assertTrue(all(b.cap_by_id[r['节点ID']]['最终候选匹配状态'] in STATUS_ORDER for r in n if r['节点ID'] in b.cap_by_id))
 def test_no_mutation_and_states_preserved(self):
  b=self.b;before=copy.deepcopy((b.nodes,b.edges,b.isolated))
  n,e=visible_data(b,Options(True,True,True));filter_graph(b,center='N313',hops=3)
  self.assertEqual(before,(b.nodes,b.edges,b.isolated))
  original={r['关系ID']:r for r in b.edges+b.isolated}
  for r in e:self.assertEqual(r['审核或置信状态'],original[r['关系ID']]['审核或置信状态'])
 def test_hidden_practice_not_positive_default(self):
  n,e=visible_data(self.b)
  self.assertFalse(any('遮挡' in str(r.get('课程证据可见性')) for r in e))
  self.assertFalse(any('遮挡' in str(r.get('证据或核验状态')) for r in n))
 def test_evidence_pages_and_behavior_boundaries(self):
  for id in ['N113','N313','N321','N322','N341','N331','N332']:
   d=detail(self.b,id,self.b.edges)
   self.assertTrue(d['课程支撑的具体技术行为']);self.assertTrue(d['课程尚未明确覆盖的工作行为'])
  for e in self.b.evidence:self.assertTrue(e['PDF页码']);self.assertTrue(e['原文摘录'])
 def test_csv_same_view_and_formula_safety(self):
  n,e,_=filter_graph(self.b,center='N113',hops=2)
  rows=list(csv.DictReader(io.StringIO(csv_bytes(n).decode('utf-8-sig'))))
  self.assertEqual({r['节点ID'] for r in rows},{r['节点ID'] for r in n})
  self.assertIn("'=1+1",csv_bytes([{'值':'=1+1'}]).decode('utf-8-sig'))
 def test_offline_html_and_escape(self):
  b=self.b;n,e,_=filter_graph(b,center='N113');n=copy.deepcopy(n);n[0]['节点名称']='</script><script>alert(1)</script>'
  h=build_html(b,n,e,'N113')
  self.assertNotIn('</script><script>alert',h);self.assertNotRegex(h,r'<(?:script|link|img)[^>]+(?:src|href)=["\']https?://')
  self.assertNotIn('/*__GRAPH_DATA__*/',h)

if __name__=='__main__':unittest.main()
