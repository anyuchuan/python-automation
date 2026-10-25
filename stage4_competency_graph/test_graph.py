import unittest,tempfile,json
from pathlib import Path
from collections import Counter
from openpyxl import load_workbook
import build as b
class GraphTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):
  cls.tmp=tempfile.TemporaryDirectory();cls.out=Path(cls.tmp.name)/'a';cls.n,cls.e,cls.ev,cls.q,cls.s,cls.c,cls.jobs,cls.hs=b.run(cls.out)
 @classmethod
 def tearDownClass(cls):cls.tmp.cleanup()
 def test_01_ability_hierarchy_ids(self):
  c=Counter(n['node_type'] for n in self.n);self.assertEqual([c['COMPETENCY_L1'],c['COMPETENCY_L2'],c['COMPETENCY_L3']],[3,9,17]);self.assertEqual({n['original_id'] for n in self.n if n['node_type']=='COMPETENCY_L3'},{f'DB{i:03}' for i in range(1,18)})
 def test_02_all_source_job_evidence_preserved(self):
  rs=[e for e in self.e if e['relation_code']=='R02'];self.assertEqual(len(rs),65);self.assertEqual(Counter(e['support_function_or_subtype'] for e in rs),{'A':11,'B':27,'C':27});self.assertEqual([e['source_attributes'] for e in rs],self.jobs)
 def test_03_hierarchy_original(self):self.assertEqual([e['source_attributes'] for e in self.e if e['relation_code']=='R01'],self.hs)
 def test_04_logic_default_exclusions(self):
  rs=[e for e in self.e if e['relation_code']=='R09'];self.assertEqual(len(rs),8);self.assertEqual(sum(e['default_visible'] for e in rs),5)
  for r in rs:self.assertEqual(r['default_visible'],r['source_attributes']['审查建议']=='保留候选')
 def test_05_course_ids_occurrences(self):
  self.assertEqual({n['original_id'] for n in self.n if n['node_type']=='COURSE_KS'},{f'CK{i:03}' for i in range(1,61)})
  rs=[e for e in self.e if e['relation_code']=='R04'];self.assertEqual([e['source_attributes'] for e in rs],self.c['occurrences']);self.assertEqual(len(rs),61)
  self.assertEqual(next(e for e in rs if e['edge_id']=='R04:KO056')['source'],'COURSE_KS:CK060')
 def test_06_context_and_interfaces(self):
  rs=[e for e in self.e if e['relation_code']=='R03'];backup=[e for e in rs if e['source']=='COURSE_KS:CK057'];self.assertEqual({e['source_attributes']['interface_id'] for e in backup},{'CTB001','CTB002'})
  for e in backup:self.assertIn('独立',e['source_attributes']['interface_source']['学生操作要求'])
  update=next(e for e in rs if e['source']=='COURSE_KS:CK027' and e['target']=='COMPETENCY_L3:DB006');self.assertEqual(update['source_attributes']['course_context'],'P04')
 def test_07_unique_ids_valid_endpoints(self):
  ids={n['node_id'] for n in self.n};self.assertEqual(len(ids),len(self.n));self.assertEqual(len({e['edge_id'] for e in self.e}),len(self.e))
  for e in self.e:self.assertIn(e['source'],ids);self.assertIn(e['target'],ids)
 def test_08_directions_and_types(self):
  by={n['node_id']:n['node_type'] for n in self.n}
  for e in self.e:
   typ,a,z=b.TYPES[e['relation_code']];self.assertEqual(e['relation_type'],typ)
   if a:self.assertEqual(by[e['source']],a);self.assertEqual(by[e['target']],z)
   else:self.assertIn((by[e['source']],by[e['target']]),[('COMPETENCY_L1','COMPETENCY_L2'),('COMPETENCY_L2','COMPETENCY_L3')])
 def test_09_no_unsupported_duplicates(self):
  rs=[(e['source'],e['target']) for e in self.e if e['relation_code']=='R03'];self.assertEqual(len(set(rs)),len(rs))
  for code in ['R02','R04']:self.assertEqual(len({e['edge_id'] for e in self.e if e['relation_code']==code}),sum(e['relation_code']==code for e in self.e))
 def test_10_supports_inferred_and_evidenced(self):
  ids={e['evidence_id'] for e in self.ev}
  for e in self.e:
   for ref in e['evidence_ids']:self.assertIn(ref,ids)
   self.assertNotIn(e['source_nature'],['A','B','C'])
   if e['relation_code']=='R03':
    self.assertEqual(e['source_nature'],'RESEARCH_INFERENCE');self.assertTrue(e['research_inference']);self.assertTrue(e['reason']);self.assertTrue(e['evidence_ids']);self.assertTrue(e['pdf_pages']);self.assertIn(e['support_function_or_subtype'],['K1','K2','K3','K4']);self.assertIn('待人工审查',e['review_status'])
 def test_11_outputs_and_checks_not_upgraded(self):
  self.assertEqual({r['学习成果ID'] for r in self.c['outcomes'] if '推断' in r['成果证据性质']},{'CO005','CO007','CO009','CO012','CO014'})
  for e in self.e:
   if e['relation_code'] in ['R07','R08']:self.assertEqual(e['source_nature'],'RESEARCH_INFERENCE')
  self.assertEqual(len(self.c['checks']),17);self.assertTrue(all(r['是否构成正式量化评价']=='否' for r in self.c['checks']))
 def test_12_hours_unlinked_retained(self):
  h=self.s['hours'][-1];self.assertEqual((h['总学时'],h['理论学时'],h['实践学时']),(64,32,32));self.assertEqual(self.s['linked_knowledge']+len(self.s['unlinked_knowledge']),60)
 def test_13_historical_files_unchanged(self):
  for f in json.loads((b.HERE/'input_snapshot.json').read_text())['files']:self.assertEqual(b.helper.gitsha(b.ROOT/f['path']),f['sha'],f['path'])
 def test_14_json_excel_consistency(self):
  for name,sheet,rows,key in [('01_能力图谱节点总表.xlsx','节点',self.n,'node_id'),('02_能力图谱关系总表.xlsx','全部关系记录',self.e,'edge_id')]:
   rs=b.read(self.out/name,sheet);self.assertEqual(len(rs),len(rows));self.assertEqual({r[key] for r in rs},{r[key] for r in rows})
  self.assertEqual(json.loads((self.out/'graph_nodes.json').read_text()),self.n)
 def test_15_paths_valid_and_no_direct_attainment(self):
  es={e['edge_id']:e for e in self.e}
  for p in json.loads((self.out/'graph_paths.json').read_text()):
   for id in p['edge_path']:self.assertIn(id,es)
   self.assertIn('非直接',p['nature']);self.assertEqual(p['default_visible'],all(es[id]['default_visible'] for id in p['edge_path']))
 def test_16_hidden_practice_not_positive(self):
  for e in self.e:
   if e['relation_code']=='R06':self.assertFalse(e['default_visible']);self.assertIn('隐藏',e['uncertainty'])
 def test_17_repeatable(self):
  other=Path(self.tmp.name)/'b';b.run(other)
  for name in ['graph_nodes','graph_edges','graph_evidence','graph_review_items','summary']:self.assertEqual(json.loads((self.out/(name+'.json')).read_text()),json.loads((other/(name+'.json')).read_text()))
 def test_18_feedback_traceability_and_counts(self):
  base=json.loads((b.HERE/'feedback_baseline.json').read_text());r=json.loads((self.out/'support_review_comparison.json').read_text());cs=r['comparisons']
  self.assertEqual({x['原关系ID'] for x in cs},{x['edge_id'] for x in base['support']});self.assertEqual(len(cs),34)
  old={x['edge_id']:x for x in base['support']}
  for x in cs:
   self.assertEqual(x['原判断说明'],old[x['原关系ID']]['reason']);self.assertEqual(x['原支撑类型'],old[x['原关系ID']]['support_function_or_subtype']);self.assertTrue(x['修改原因']);self.assertTrue(x['证据引用']);self.assertIn('待研究者确认',x['当前审查状态'])
  self.assertEqual(sum(x['是否实际修改']=='是' for x in cs),13)
  self.assertEqual(r['recommendation_counts'],{'建议保留':21,'建议调整关系说明或适用边界':12,'建议调整支撑类型':1})
  self.assertEqual(len(r['unlinked_review']),4)
 def test_19_review_ids_categories_preserved(self):
  baseline=json.loads((b.HERE/'feedback_baseline.json').read_text())
  self.assertEqual({q['review_id'] for q in self.q},{q['review_id'] for q in baseline['reviews']});self.assertEqual(len(self.q),69)
  for q in self.q:self.assertIn(q['review_category'],'ABCDE');self.assertTrue(q['human_confirmation_required']);self.assertTrue(q['classification_reason']);self.assertTrue(q['handling_suggestion'])
 def test_20_feedback_excel_matches_json(self):
  data=json.loads((self.out/'support_review_comparison.json').read_text());rows=b.read(self.out/'13_第四阶段第二轮候选关系审查修订对照.xlsx','34关系逐条对照')
  for a,z in zip(rows,data['comparisons']):
   for k,v in z.items():self.assertEqual(a[k],json.dumps(v,ensure_ascii=False) if isinstance(v,list) else v)
  self.assertEqual(len(rows),len(data['comparisons']))
  support=[e for e in self.e if e['relation_code']=='R03'];self.assertEqual(len(support),34)
  fk=next(e for e in support if e['edge_id']=='R03:034');self.assertEqual(fk['support_function_or_subtype'],'K4');self.assertIn('不要求',fk['uncertainty'])

if __name__=='__main__':unittest.main()
