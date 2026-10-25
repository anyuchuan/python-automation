import unittest,math
from pathlib import Path
from app.competency_store import GraphStore
from app.explorer_model import overview
from app.radial_layout import radial_layout
ROOT=Path(__file__).resolve().parents[2]
class RadialTests(unittest.TestCase):
 def test_real_nodes_coordinates_and_no_collision(self):
  g=GraphStore(ROOT);nodes,edges=overview(g);r=radial_layout(nodes,edges)
  self.assertEqual(set(r['positions']),{n['node_id'] for n in nodes});self.assertEqual(len(nodes),29);self.assertEqual(len(edges),26)
  ps=list(r['positions'].values())
  for p in ps:
   self.assertTrue(all(math.isfinite(p[k]) for k in ['x','y']));self.assertGreaterEqual(p['x'],0);self.assertGreaterEqual(p['y'],0);self.assertLessEqual(p['x']+p['width'],r['width']);self.assertLessEqual(p['y']+p['height'],r['height'])
  for i,a in enumerate(ps):
   for b in ps[i+1:]:self.assertFalse(abs(a['x']-b['x'])<a['width'] and abs(a['y']-b['y'])<a['height'])
 def test_default_and_switch(self):
  from streamlit.testing.v1 import AppTest
  at=AppTest.from_file(str(ROOT/'app/main.py'),default_timeout=30).run();at.radio(key='platform_entry').set_value('专业能力图谱').run();self.assertEqual(at.radio(key='ex_layout').value,'径向图');self.assertFalse(at.exception)
  at.radio(key='ex_layout').set_value('层级树').run();self.assertFalse(at.exception);at.radio(key='ex_layout').set_value('径向图').run();self.assertFalse(at.exception)
if __name__=='__main__':unittest.main()
