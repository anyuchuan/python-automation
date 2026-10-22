import unittest
from streamlit.testing.v1 import AppTest
from app.platform_model import MODULES

class PageTests(unittest.TestCase):
 def app(self):return AppTest.from_file('app/main.py',default_timeout=30).run()
 def check(self,at):self.assertFalse(at.exception);self.assertFalse(at.error)
 def test_default_tree_and_all_six_modules(self):
  at=self.app();self.check(at)
  self.assertEqual(at.radio(key='module').value,MODULES[0]);self.assertEqual(at.radio(key='graph_view').value,'树状层级')
  self.assertEqual(len(at.metric),0)
  for module in MODULES:
   at.radio(key='module').set_value(module).run();self.check(at)
 def test_search_locate_switch_and_risk_gate(self):
  at=self.app();at.text_input(key='cap_query').set_value('N113').run();self.check(at)
  self.assertEqual(at.session_state['selected_cap'],'N113');self.assertEqual(at.session_state['tree_expanded'],['N1','N11'])
  at.radio(key='graph_view').set_value('网络关系').run();self.check(at);self.assertEqual(at.session_state['net_center'],'N113')
  at.multiselect(key='net_types').set_value([]).run();self.check(at)
  self.assertTrue(at.checkbox(key='high_risk').disabled)
  at.checkbox(key='explore').check().run();at.checkbox(key='high_risk').check().run();self.check(at)
  at.checkbox(key='explore').uncheck().run();self.assertFalse(at.checkbox(key='high_risk').value)
 def test_matrix_and_evaluation_denominators(self):
  at=self.app();at.radio(key='module').set_value(MODULES[2]).run();self.check(at)
  self.assertEqual(len(at.multiselect(key='matrix_columns').value),8)
  at.multiselect(key='matrix_grades').set_value([]).run();self.check(at)
  at.multiselect(key='matrix_columns').set_value([]).run();self.check(at)
  at.radio(key='module').set_value(MODULES[3]).run();self.check(at)
  self.assertEqual([m.value for m in at.metric],['1 / 34','18 / 34','14 / 34','1 / 34'])
  at.multiselect(key='evaluation_grades').set_value([]).run();self.check(at)
 def test_shortcuts_preserve_ability_and_empty_search(self):
  at=self.app();at.text_input(key='cap_query').set_value('N313').run()
  at.button(key='to_mapping').click().run();self.check(at)
  self.assertEqual(at.radio(key='module').value,MODULES[2]);self.assertEqual(at.session_state['selected_cap'],'N313')
  at.button(key='to_evaluation').click().run();self.check(at)
  self.assertEqual(at.session_state['selected_cap'],'N313')
  at.radio(key='module').set_value(MODULES[0]).run();at.text_input(key='cap_query').set_value('不存在XYZ').run();self.assertTrue(at.info);self.check(at)
if __name__=='__main__':unittest.main()
