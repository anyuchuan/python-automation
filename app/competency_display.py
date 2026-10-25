"""中文展示映射，仅转换标签，不改源值。"""
NODE_LABELS={'COMPETENCY_L1':'一级候选能力','COMPETENCY_L2':'二级候选能力','COMPETENCY_L3':'三级候选能力','WORK_TASK':'典型工作任务','COURSE_KS':'课程知识技能点','TEACHING_PROJECT':'教学项目','TEACHING_ACTIVITY':'教学活动','PRACTICAL_TASK':'独立实训','LEARNING_OUTPUT':'成果或潜在产物','ASSESSMENT_EVIDENCE':'评价检查证据'}
REL_LABELS={'R01':'能力层级包含','R02':'任务—能力证据','R03':'知识技能支撑候选','R04':'知识在项目出现','R05':'项目安排活动','R06':'项目—实训内容相关','R07':'活动—潜在成果','R08':'活动—检查证据','R09':'能力非层级候选'}
def directed(edge):return not (edge['relation_code']=='R09' and edge['support_function_or_subtype']=='组合关系')
def relation_fields(g,e):
 def label(id):
  n=g.by_id[id];return n['original_id']+' · '+n['name']
 return {'关系编号':e['edge_id'],'关系类型':REL_LABELS[e['relation_code']]+' / '+e['relation_type'],'起点':label(e['source']),'终点':label(e['target']),'方向':'起点 → 终点' if directed(e) else '组合配对，无先后方向','关系状态':e.get('edge_status','原始资料未提供'),'来源阶段':e['source_stage'],'来源文件':e['source_file'],'岗位A/B/C':e['support_function_or_subtype'] if e['relation_code']=='R02' else '不适用','课程K1—K4':e['support_function_or_subtype'] if e['relation_code']=='R03' else '不适用','判断说明':e['reason'] or '原始资料未提供','适用边界':e['uncertainty'] or '原始资料未提供','是否研究推断':'是' if e['research_inference'] else '按源数据非研究推断标志记录','来源性质':e['source_nature'],'人工审查状态':e['review_status']}
