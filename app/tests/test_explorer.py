import unittest
from pathlib import Path
from app.competency_store import GraphStore
from app.explorer_model import overview,topic,curriculum,highlight
ROOT=Path(__file__).resolve().parents[2]
class ExplorerTests(unittest.TestCase):
 def setUp(self):self.g=GraphStore(ROOT)
 def test_overview_real_complete(self):
  from collections import Counter
  ns,es=overview(self.g);self.assertEqual(Counter(n['node_type'] for n in ns),{'COMPETENCY_L1':3,'COMPETENCY_L2':9,'COMPETENCY_L3':17});self.assertEqual(len(es),26);self.assertTrue(all(e in self.g.edges for e in es))
 def test_topics_original_status(self):
  for n in self.g.search(types=['COMPETENCY_L3']):
   for kind,code in [('岗位关联','R02'),('课程关联','R03')]:
    ns,es=topic(self.g,n['node_id'],kind);self.assertTrue(all(e in self.g.edges and e['default_visible'] and e['relation_code']==code and e['target']==n['node_id'] for e in es))
 def test_curriculum_unique_shared_retained(self):
  projects,_=curriculum(self.g);self.assertEqual(len(projects),8);allids=set()
  for p in projects:
   ns,es=curriculum(self.g,p['node_id']);allids.update(n['node_id'] for n in ns if n['node_type']=='COURSE_KS');self.assertEqual(len(ns),len({n['node_id'] for n in ns}))
  self.assertEqual(len(allids),60)
 def test_highlight_real_paths(self):
  ns,es=overview(self.g);id='COMPETENCY_L3:DB017';ids=highlight(self.g,id,ns,es);self.assertEqual(ids,{id,*self.g.parents(id)});self.assertNotIn('COMPETENCY_L3:DB001',ids)
 def test_page_default_no_filters(self):
  from streamlit.testing.v1 import AppTest
  at=AppTest.from_file(str(ROOT/'app/main.py'),default_timeout=30).run();at.radio(key='platform_entry').set_value('专业能力图谱').run();self.assertFalse(at.exception);self.assertEqual(len(at.multiselect),0);self.assertEqual(len(at.text_input),0);self.assertTrue(any(b.label=='完整课程知识体系' for b in at.button));self.assertEqual(len(at.metric),3)
if __name__=='__main__':unittest.main()
