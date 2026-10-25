"""PR9人工反馈落实：候选审查建议，不代表研究者确认。"""
from collections import Counter
import json

def apply(here,out,helper,support,reviews,course):
 baseline=json.loads((here/'feedback_baseline.json').read_text());old={r['edge_id']:r for r in baseline['support']};rules=json.loads((here/'support_rules.json').read_text());rr={r['relation_id']:r for r in rules};comparisons=[]
 for e in support:
  o=old[e['edge_id']];a=e['source_attributes'];changed=any(e[k]!=o[k] for k in ['reason','uncertainty','support_function_or_subtype']);r=rr[e['edge_id']]
  comparisons.append({'原关系ID':e['edge_id'],'能力编号':a['competency_id'],'能力名称':a['competency_name'],'能力行为':a['competency_behavior'],'知识技能点编号':a['knowledge_id'],'知识技能点名称':a['knowledge_name'],'原支撑类型':o['support_function_or_subtype'],'建议支撑类型':e['support_function_or_subtype'],'原判断说明':o['reason'],'原适用边界':o['uncertainty'],'建议修订说明':e['reason']+'；'+e['uncertainty'],'证据引用':e['evidence_ids'],'PDF页码':e['pdf_pages'],'审查建议':r['review_recommendation'],'是否实际修改':'是' if changed else '否','修改原因':e['uncertainty'] if changed else '对象、操作及预期结果对应具体子行为，保留原K类及候选限制；不按名称认定完整能力','当前审查状态':'Codex候选审查意见，待研究者确认'})
 for q in reviews:
  id=q['review_id'];cat='A';why='涉及技术行为与完整能力边界'
  if id.startswith('STAGE3:') and not id.endswith('015'):cat='B';why='继承岗位摘要明确程度、职责对象及证据来源限制'
  if id=='STAGE3:FR-Q015' or id in ['Q:FR-L007','Q:FR-L008']:cat='C';why='关系类型、方向及默认边集判定'
  if id.startswith('COURSE:'):
   n=int(id[-3:]);cat='D' if n in [1,2,3,4,5,7,8,10,12] else 'A';why='课程原始文本或任务安排限制' if cat=='D' else '术语及课程知识技术边界'
  q.update(review_category=cat,classification_reason=why,original_review_id=id,handling_suggestion='保留问题，结合原依据逐项人工确认；测试通过不关闭研究问题',impact_scope=q['object_id'],human_confirmation_required=True)
  if id.startswith('Q:R03:'):
   cr=next(x for x in comparisons if x['原关系ID']==q['object_id']);q['handling_suggestion']=cr['审查建议']+'；'+cr['建议修订说明']
 cases=[]
 specs=[('DB009',['CK004','CK005'],'P01有安装配置连接及系统概念，未明确数据库实例运行指标/状态检查任务；不是仅总体目标问题。','环境部署不等于运行状态检查'),('DB010',['CK006'],'P01难点含安装错误排查与环境变量配置，属于安装环境故障，未明确数据库运行异常诊断及日志定位。','安装排错不能等同运行异常诊断'),('DB014',['CK025'],'P04含INSERT单条批量插入，未明确外部数据源、字段映射、导入过程及结果对账。','批量INSERT不自动等于数据导入'),('DB015',[],'知识/能力总体目标及P08思政提到权限/DCL；P08具体技术操作为备份恢复与事务，未明确授权主体、对象、权限配置操作。','总体目标/思政提及不足以证明独立配置技能')]
 for id,kids,finding,reason in specs:
  refs=[]
  for k in course['knowledge']:
   if k['知识点ID'] in kids:refs+=k['证据ID'].split(';')
  if id=='DB015':
   for r in course['other_level_mentions']:refs+=r['证据ID'].split(';')
  cases.append({'能力编号':id,'课程相关内容核查':finding,'相近但不能等同':reason,'证据引用':sorted(set(refs)),'新增关系判断':'依据不足，本轮保持未关联','审查状态':'候选核查意见，待研究者确认'})
 counts=Counter(x['审查建议'] for x in comparisons);cats={k:sum(q['review_category']==k for q in reviews) for k in 'ABCDE'}
 helper.write_book(out/'13_第四阶段第二轮候选关系审查修订对照.xlsx',{'34关系逐条对照':comparisons,'四项未关联能力核查':cases,'69问题分类':reviews})
 (out/'support_review_comparison.json').write_text(json.dumps({'comparisons':comparisons,'unlinked_review':cases,'recommendation_counts':dict(counts),'question_categories':cats},ensure_ascii=False,indent=2)+'\n')
 (out/'14_第四阶段第二轮人工审查反馈落实说明.md').write_text('# PR9候选关系人工反馈落实\n\n基线提交'+baseline['commit']+'；34原关系与69问题逐条保留。所有意见为Codex候选意见，未经研究者或专家确认。\n\n建议统计：'+json.dumps(dict(counts),ensure_ascii=False)+'；排除0、证据不足继续待审0（所有保留/调整也仍待人工确认）。实际新增0、删除0、调整13条，其中1条K1→K4、12条边界说明收窄。节点205及关系记录265不变。\n\nDB013限定已有质量条件下筛选待核查记录，保持K4；DB007主外键保持K4，区分设计与SQL约束实现；DB011创建/添加/删除只支持实施子环节，效果对比不证明完整调整闭环；DB017外键关系调整K4，JOIN不以外键为必要前提。DB002不外推完整清洗。原UPDATE/ALTER TABLE、备份/恢复接口及优化目标/索引方法继续区分。\n\n四项未关联能力逐项核查见13第二工作表，均未强行补边。问题分类为'+json.dumps(cats,ensure_ascii=False)+'；空类别不人为填充。69问题均保留原编号与人工确认要求，第三阶段候选关系和课程隐藏文本限制不升级。\n\n源规则、生成脚本、方法及测试同步更新；新增feedback_baseline.json保存原34判断与69问题，13与机器对照保留前后值。输出02/03/07/08等由源生成，既有历史阶段及PDF/网页只读。测试命令及实际结果见上级验证结果.txt。\n\n本轮只完成候选图谱关系的审查反馈修订，不代表专家验证完成，不代表学生能力达成度评价完成。\n')
 return comparisons
