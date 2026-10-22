import copy
import unittest
from collections import Counter
from app.data import load_bundle,Options,visible_data
from app.platform_model import *
from app.controller import choose,apply_event

class PlatformTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):cls.b=load_bundle();cls.tree=hierarchy(cls.b)
 def test_exact_real_hierarchy_and_counts(self):
  t=self.tree;self.assertEqual(len(t),53);self.assertEqual(Counter(n['level'] for n in t.values()),{'一级':5,'二级':14,'三级':34})
  self.assertEqual(sum(bool(n['parent']) for n in t.values()),48)
  self.assertEqual({(n['parent'],id) for id,n in t.items() if n['parent']},{(e['起点ID'],e['终点ID']) for e in self.b.edges if e['关系类型']=='建议能力层级包含' and e['起点ID'] in t and e['终点ID'] in t})
  self.assertFalse(any(id.startswith('S') for id in t))
 def test_expand_collapse_search_and_locate(self):
  t=self.tree;self.assertEqual(visible_tree(t,[]),['N1','N2','N3','N4','N5'])
  self.assertEqual(len(visible_tree(t,['N1'])),7)
  self.assertEqual(len(visible_tree(t,[id for id,n in t.items() if n['level']!='三级'])),53)
  self.assertEqual(find_capabilities(t,'n113'),['N113']);self.assertIn('N321',find_capabilities(t,'备份'))
  self.assertEqual(find_capabilities(t,'不存在XYZ'),[])
  s={};choose(s,t,'N113');self.assertEqual(s['tree_expanded'],['N1','N11']);self.assertIn('N113',visible_tree(t,s['tree_expanded']))
 def test_cross_view_and_module_shared_selection(self):
  s={};t=self.tree
  self.assertTrue(apply_event(s,{'action':'select','id':'N113'},t,self.b));self.assertEqual(s['net_center'],'N113')
  self.assertTrue(apply_event(s,{'action':'network_select','id':'N313'},t,self.b));self.assertEqual(s['selected_cap'],'N313')
  apply_event(s,{'action':'tree_locate','id':'N313'},t,self.b);self.assertEqual(s['graph_view'],'树状层级');self.assertEqual(s['tree_expanded'],['N3','N31'])
  apply_event(s,{'action':'matrix_select','id':'N113','project':'P05'},t,self.b);self.assertEqual((s['selected_cap'],s['selected_project']),('N113','P05'))
 def test_partial_missing_and_status_not_invented(self):
  d=capability_detail(self.b,'N313');self.assertEqual(d['grade'],'部分支撑');self.assertEqual(d['definition'],'暂无记录')
  self.assertTrue(d['technical']);self.assertIn('待核验',d['state']);self.assertEqual(capability_detail(self.b,'JOB001'),None)
 def test_network_filters_before_neighborhood_and_limit(self):
  b=self.b;n,e,total=network_slice(b,'N113',hops=2,limit=60)
  self.assertLessEqual(len(n),60);self.assertFalse(any(x['展示风险说明'] for x in e))
  n,e,_=network_slice(b,'N113',relations=[]);self.assertEqual([x['节点ID'] for x in n],['N113']);self.assertEqual(e,[])
  self.assertEqual(network_slice(b,'N113',types=[]),( [],[],0))
  n,e,_=network_slice(b,'N113',evidence=['课程文本匹配（待审核）']);self.assertTrue(all(evidence_group(x)=='课程文本匹配（待审核）' for x in e))
  n,e,_=network_slice(b,'N113',types=['职业能力建议'],limit=3);self.assertLessEqual(len(n),3);ids={x['节点ID'] for x in n};self.assertTrue(all(x['起点ID'] in ids and x['终点ID'] in ids for x in e))
 def test_explicit_risk_gating(self):
  for option in [Options(),Options(False,True,True)]:
   _,e,_=network_slice(self.b,'N113',option,hops=3,limit=1000);self.assertFalse(any(x['展示风险说明'] for x in e))
  _,e,_=network_slice(self.b,'N113',Options(True,True,True),hops=3,limit=1000);self.assertTrue(any(x['展示风险说明'] for x in e))
 def test_matrix_272_cells_and_scoped_evidence(self):
  b=self.b;self.assertEqual(len(b.project_cells),272)
  self.assertEqual(project_cell(b,'N113','P05')['项目层匹配等级'],'直接支撑')
  self.assertTrue(cell_evidence(b,project_cell(b,'N113','P05')))
  for c in b.project_cells:
   for e in cell_evidence(b,c):self.assertIn(e['证据ID'],split_ids(c['正向课程证据ID'])+split_ids(c['待审课程提及证据ID']))
   if c['项目层匹配等级']=='未见明确支撑':self.assertEqual(cell_evidence(b,c),[])
 def test_full_counts_are_independent_of_local_graph(self):
  before=global_counts(self.b);network_slice(self.b,'N111',hops=1,limit=4)
  self.assertEqual(global_counts(self.b),before);self.assertEqual([before[k] for k in ['候选岗位','典型任务','三级能力']],[50,79,34])
 def test_optimization_traceability_and_readonly(self):
  before=copy.deepcopy((self.b.nodes,self.b.edges,self.b.capabilities));rows=optimization_rows(self.b)
  self.assertEqual(len(rows),34);self.assertTrue(all(r['来源'] and r['审核状态']=='待人工审核' for r in rows))
  self.assertEqual(before,(self.b.nodes,self.b.edges,self.b.capabilities))
if __name__=='__main__':unittest.main()
