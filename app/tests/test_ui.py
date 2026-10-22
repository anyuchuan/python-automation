import unittest
from streamlit.testing.v1 import AppTest

class PageTests(unittest.TestCase):
 def test_pages_and_filters(self):
  at=AppTest.from_file('app/main.py',default_timeout=30).run()
  self.assertFalse(at.exception)
  self.assertIn('50',[m.value for m in at.metric])
  for page in ['图谱探索','教学项目','支撑分析','来源与局限']:
   at.radio(key='page').set_value(page).run();self.assertFalse(at.exception,page)
  at.radio(key='page').set_value('图谱探索').run()
  self.assertTrue(at.checkbox(key='high_risk').disabled)
  at.checkbox(key='explore').check().run()
  self.assertFalse(at.checkbox(key='high_risk').disabled)
  at.checkbox(key='high_risk').check().run()
  at.checkbox(key='explore').uncheck().run()
  self.assertFalse(at.checkbox(key='high_risk').value)
  at.text_input(key='query').set_value('不存在XYZ').run()
  self.assertTrue(at.info);self.assertFalse(at.exception)
  at.text_input(key='query').set_value('').run()
  at.multiselect(key='node_types').set_value([]).run()
  self.assertTrue(at.info);self.assertFalse(at.exception)
 def test_project_jump_and_analysis_empty(self):
  at=AppTest.from_file('app/main.py',default_timeout=30).run()
  at.radio(key='page').set_value('教学项目').run()
  button=next(b for b in at.button if b.label=='在图谱中查看本项目 →');button.click().run()
  self.assertEqual(at.radio(key='page').value,'图谱探索');self.assertFalse(at.exception)
  at.radio(key='page').set_value('支撑分析').run()
  self.assertEqual([m.value for m in at.metric],['1 / 34','18 / 34','14 / 34','1 / 34'])
  at.multiselect(key='analysis_statuses').set_value([]).run();self.assertFalse(at.exception)
if __name__=='__main__':unittest.main()
