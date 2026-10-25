"""四源阶段成果：冻结输入、先条目审查再比对；不重写基线能力。"""
from pathlib import Path
import json,sys,hashlib,argparse
from collections import Counter
BASE=Path(__file__).resolve().parent
sys.path.insert(0,str(BASE.parent))
import build as legacy

def load(p):return json.loads(Path(p).read_text())
def run(out=None):
 out=Path(out or BASE/'output');out.mkdir(parents=True,exist_ok=True)
 v=load(BASE.parent/'output/v2_research.json');sources=load(BASE/'source_index.json');raw=load(BASE/'raw_requirements.json')
 for j in v['jobs']:
  sources.append(dict(id='S2-'+j['id'],category='S2',title=j['name'],publisher=j['publisher'],url=j['url'],status='继承：'+j['page_state'],version=j['published_at'],locator=j['locator'],access_date=j['collected_at'],limits=j['limitations'],review=j['review_status']))
 sources.append(dict(id='S2-TOURISM',category='S2',title='既有旅游50候选岗位及48任务',publisher='仓库历史研究资料',url=legacy.JOBS,status='继承摘要；招聘原文未核验',version='PR #13基线',locator=legacy.TOURISM,access_date='本轮未重访招聘页面',limits='仅旅游应用情境，不当新增跨行业样本',review='待审'))
 for t in v['tasks']:
  for idx,e in enumerate(t['evidence'],1):
   j=next(j for j in v['jobs'] if j['id']==e['job_id'])
   raw.append(dict(id=f"FR-S2-{t['id']}-{idx}",source_id='S2-'+j['id'],excerpt=j['excerpt'],locator=e['location']+'；'+j['locator'],technical_action=e['fragment'],goal=t['goal'],expected_product=t['result'],level=j['requirements_summary'],decision='部分纳入',excluded_content='不继承岗位全部技术：大数据生态、复杂集群、深度调优等超范围活动排除；SQL实现未明者保留ABC解释',evidence_nature='继承岗位'+e['abc']+'类；'+e['interpretation'],candidate_capabilities=[c['id'] for c in v['capabilities'] if t['id'] in c['task_ids']],review='待研究者审查'))
 si={s['id']:s for s in sources};relations=[]
 for r in raw:
  r['category']=si[r['source_id']]['category'];r['verification_status']=si[r['source_id']]['status']
  r['adaptation_reason']='限关系数据库基础、单机/普通教学环境；成果及评价形式为研究拟定；待人工确认' if r['decision']!='排除' else '涉及明确排除的高级技术，不能进入候选能力关系'
  if r['decision']=='排除':continue
  for cap in r['candidate_capabilities']:
   relations.append(dict(id=f'FE{len(relations)+1:03}',capability_id=cap,requirement_id=r['id'],source_id=r['source_id'],category=r['category'],support_scope=r['technical_action'],nature='研究比对候选；不证明完整能力所有行为',review='待研究者审查'))
 matrix=[];revisions=[]
 for c in v['capabilities']:
  es=[e for e in relations if e['capability_id']==c['id']];s2={e['source_id'] for e in es if e['category']=='S2'}
  pubs=sorted({si[s]['publisher'] for s in s2})
  row=dict(capability_id=c['id'],name=c['name'],**{k:[e['requirement_id'] for e in es if e['category']==k] for k in ['S1','S2','S3','S4']},independent_job_records=len(s2),independent_employers=len(pubs),employers=pubs,source_independence='同一标准多个条款只算1份标准；同雇主岗位不重复计雇主；S1非岗位证据',gaps='S3正文未核验；S4公开副本现行状态待确认；单一来源不代表行业频率',review='待研究者审查')
  matrix.append(row)
  revisions.append(dict(capability_id=c['id'],name=c['name'],actual_change='未修改基线定义与编号',suggestion='补充S4初级条款候选依据；不改写成企业岗位明示' if row['S4'] else '保留现有来源属性与证据缺口',source_requirements=[e['requirement_id'] for e in es],boundary=c['inference'],review='待审建议，未自动采纳'))
 issues=[dict(id='FQ01',problem='S3正文未核验',resolution='保留官方线索，不产生条款或S3能力关系'),dict(id='FQ02',problem='S4发布机构官网直链及现行状态待确认',resolution='已核对版本正文，使用公开托管副本；不可宣称现行有效认证'),dict(id='FQ03',problem='FR009基础用户权限暂未对应现有13项',resolution='保留未关联条目，建议后续审查是否建立独立候选，不能强行补线'),dict(id='FQ04',problem='企业证据小规模且运维集中',resolution='不因新增标准条款自动消除采样限制'),dict(id='FQ05',problem='S4表定义及视图可补充外部规范依据',resolution='不将标准条款视为已核验企业需求；模型设计与完整性验证仍缺明确条款'),dict(id='FQ06',problem='S1只为背景',resolution='不得将政务大数据体系要求改称企业关系数据库技能要求')]
 summary=dict(source_counts=dict(Counter(s['category'] for s in sources)),requirements=len(raw),excluded=sum(r['decision']=='排除' for r in raw),partial=sum(r['decision']=='部分纳入' for r in raw),relations=len(relations),capabilities=len(matrix),s3_verified_body=0,s4_verified_version_body=1,baseline_unchanged=True)
 data=dict(summary=summary,sources=sources,requirements=raw,relations=relations,matrix=matrix,revisions=revisions,issues=issues)
 (out/'four_sources.json').write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n')
 tables={'01_四源资料索引及核验状态.xlsx':{'来源索引':sources},'02_S3_S4正式标准核验记录.xlsx':{'标准核验':[s for s in sources if s['category'] in ['S3','S4']]},'03_四源原始能力需求条目池.xlsx':{'全部条目':raw},'04_高职适配与高级技术排除记录.xlsx':{'逐条筛选':raw},'05_现有13项能力证据比对与建议.xlsx':{'逐项建议':revisions,'四源矩阵':matrix},'06_四源能力证据关系.xlsx':{'候选关系':relations},'07_研究限制与未解决问题.xlsx':{'待审问题':issues}}
 for f,sheets in tables.items():legacy.xlsx(out/f,sheets)
 report=f'''# 四源采集与能力融合阶段报告

2026-10-26，继续PR #13原分支；未新建V3。本轮采用“先原始需求条目池、逐项高职筛选、再与既有能力比对”的流程。基线文件保留，四源结果放在独立子目录，避免覆写历史事实。

## 资料与核验

S1：1份政府公开指南已打开，限数据质量背景。大数据平台、云、AI等技术全部排除，不将政策或趋势当岗位要求。
S2：继承22份跨行业筛查记录及1份旅游资料集合（50岗位/48任务），不新增招聘样本；8份参与原有有限任务分析，保留ABC、摘要及未核验限制。访问日期沿用原采集日，未伪称本轮重访。
S3：2条互相关联线索（标准正文入口和颁布通知），是1项标准，不是两项独立标准。《数据库运行管理员S》2025年版、职业编码4-04-05-04与人社厅发〔2025〕12号仅有官方检索索引线索。通知打开超时，正文接口返回Excel MIME异常，直接公开下载出口403。未获取完整正文，提取0条标准技能要求、生成0条S3关系；未采用2023征求意见稿或商业预览补条款。
S4：2份资料，其中《数据库管理系统职业技能等级标准》2021年1.0版正文16页已核对，封面为武汉达梦数据库股份有限公司制定、2021年3月发布；6.2表1为初级。本轮仅提取初级基础操作，并将高级架构条目排除。公开托管PDF不是发布机构官网直链，尚未确认现行有效性，保持候选依据及待审；没有虚构标准编号。另一份中国通信企业协会文件封面明确征求意见稿、编号日期占位，0条正式标准需求，不使用其中技能条款。

## 条目、筛选和比对

{summary['requirements']}条原始需求，含S1背景/排除项、S2继承片段和S4版本正文条目；排除{summary['excluded']}条，部分纳入{summary['partial']}条。每条保留出处、原文短摘录或明确转述、技术行动、目标、成果、等级、适配理由及状态。S2成果沿用研究拟定，标准也没有被改写为学生评分量规。未核验S3不生成条目。

{summary['relations']}条四源候选比对关系；同一个条款关联多个能力不等于多个来源，同雇主与同标准依赖明确记录。矩阵逐项列出真实需求编号和缺口，不计算得分。S1背景关联仅为研究解释。S4支持V2C001环境、V2C003索引子活动、V2C010记录维护、V2C012表定义、V2C013视图定义的部分行为，不代表完整定义全部有外部依据。尤其表定义不等于完整性验证或关系模型设计。基础用户权限FR009有初级条款，但暂不强制加入现有13项，保留待审。

本轮保留4—9—13候选体系：13项编号和定义均未实际改动，合并/拆分/新增均0。所有比对意见在05表中逐项保留。原43条课程候选支撑、11条旅游应用及21条跨行业关系保持基线，不把S4证据混进ABC招聘关系，也不伪造新增招聘要求。课程60CK不是S1—S4需求来源；没有S5、专业教学标准或平台修改。

## 局限与人工审查

S3正文缺口仍在；S4应确认发布渠道和现行状态；标准指导下操作与独立职业能力不能等同；岗位样本仍少且运维集中；权限能力及视图能力粒度需审查。原始池中保留未关联条目和排除记录。具备阶段人工审查条件，不具备宣称四源完整覆盖或最终能力定稿的条件。

生成命令：`python competency_reconstruction_v2/four_sources/build_four.py`。测试命令：`python -m unittest discover -s competency_reconstruction_v2/four_sources -p 'test_*.py' -v`；实际日志另存测试结果.txt。JSON和7份Excel均可重生成，基线哈希保护覆盖原V2及历史关键输入。测试通过不关闭研究问题。

本轮形成的是适用于高职《数据库原理及应用》课程的四源候选证据比对阶段成果，仍需研究者审查，不代表已经完成行业专家验证或学生能力达成度评价。
'''
 (out/'四源采集与能力融合研究报告.md').write_text(report)
 return data
if __name__=='__main__':
 a=argparse.ArgumentParser();a.add_argument('--output',type=Path);ns=a.parse_args();print(json.dumps(run(ns.output)['summary'],ensure_ascii=False,indent=2))
