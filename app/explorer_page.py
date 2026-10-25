"""以画布探索为中心，无默认筛选/右侧证据栏。"""
import streamlit as st
from pathlib import Path
from collections import Counter
from app.competency_page import cached
from app.competency_store import GraphDataError
from app.explorer_model import overview,topic,curriculum,highlight
from app.competency_display import NODE_LABELS,REL_LABELS,directed,relation_fields
from app.visuals import visual
COLORS={'COMPETENCY_L1':'#9ccbb1','COMPETENCY_L2':'#c2ddcb','COMPETENCY_L3':'#b8d8eb','WORK_TASK':'#ead3b6','COURSE_KS':'#b7ddd9','TEACHING_PROJECT':'#d4cee9'}
def render(root):
 try:
  folder=Path(root)/'stage4_competency_graph/output';signature=tuple((p.name,p.stat().st_mtime_ns,p.stat().st_size) for p in [folder/(n+'.json') for n in ['graph_nodes','graph_edges','graph_evidence','graph_review_items','graph_paths']]);g=cached(str(root),signature)
 except (OSError,GraphDataError) as exc:st.error(str(exc));return
 for k,v in {'ex_space':'总览','ex_cap':None,'ex_node':None,'ex_edge':None,'ex_drawer':False,'ex_project':None,'ex_topic':'岗位关联'}.items():
  if k not in st.session_state:st.session_state[k]=v
 def go(space):
  st.session_state.ex_space=space;st.session_state.ex_edge=None;st.session_state.ex_drawer=False
  if space=='总览' and st.session_state.ex_cap:st.session_state.ex_node=st.session_state.ex_cap
  st.rerun()
 space=st.session_state.ex_space;st.title('专业能力图谱')
 if space=='总览':
  st.caption('点击能力探索真实层级、岗位与课程关联。候选关系仍待研究者审查。')
  counts=Counter(n['node_type'] for n in g.nodes);cols=st.columns(3)
  for col,t,label in zip(cols,['COMPETENCY_L1','COMPETENCY_L2','COMPETENCY_L3'],['一级候选能力','二级候选能力','三级候选能力']):col.metric(label,counts[t])
  if st.button('完整课程知识体系'):go('课程体系')
  nodes,edges=overview(g);layout='hierarchy';center=None
 elif space=='专题':
  if st.button('返回能力总览'):go('总览')
  cap=st.session_state.ex_cap;n=g.by_id[cap];st.subheader(n['original_id']+' · '+n['name']);st.caption(n.get('description') or '暂无明确说明')
  kind=st.radio('专题视角',['岗位关联','课程关联'],horizontal=True,key='ex_topic')
  if st.session_state.get('ex_lasttopic')!=(cap,kind):st.session_state.ex_edge=None;st.session_state.ex_drawer=False;st.session_state.ex_node=cap;st.session_state.ex_lasttopic=(cap,kind)
  nodes,edges=topic(g,cap,kind);center=cap;layout='radial'
  st.caption('A/B/C是岗位摘要证据性质，K1—K4是课程支撑功能；均不表示重要性或学生达成水平。')
  if not edges:st.info('当前源数据默认展示规则下暂无可显示关联；不代表没有教学价值或课程缺陷。隐藏记录可在原研究审查页面查看。')
 elif space=='课程体系':
  if st.button('返回能力总览'):go('总览')
  st.subheader('完整课程知识体系')
  if st.session_state.ex_project and st.button('返回8项目总览'):st.session_state.ex_project=None;st.session_state.ex_node=None;st.rerun()
  nodes,edges=curriculum(g,st.session_state.ex_project);center=st.session_state.ex_project;layout='radial' if center else 'projects'
 selected=st.session_state.ex_node;active=highlight(g,selected,nodes,edges) if selected in {n['node_id'] for n in nodes} else set()
 payload={'nodes':[{'id':n['node_id'],'code':n['original_id'],'name':n['name'],'type':n['node_type'],'color':COLORS.get(n['node_type'],'#d0d8df')} for n in nodes],'edges':[{'id':e['edge_id'],'from':e['source'],'to':e['target'],'label':e['support_function_or_subtype'] if e['relation_code'] in ['R02','R03'] else REL_LABELS[e['relation_code']],'directed':directed(e),'dash':bool(e['research_inference'])} for e in edges]}
 event=visual(mode='explorer',graph=payload,layout=layout,center=center,selected=selected,highlight=list(active),selected_edge=st.session_state.ex_edge,key='explorer_'+space)
 if event and event.get('nonce')!=st.session_state.get('ex_nonce'):
  st.session_state.ex_nonce=event['nonce'];id=event.get('id');action=event.get('action')
  if action=='select' and id in {n['node_id'] for n in nodes}:st.session_state.ex_node=id;st.session_state.ex_edge=None;st.session_state.ex_drawer=False
  if action=='edge' and id in {e['edge_id'] for e in edges}:st.session_state.ex_edge=id;st.session_state.ex_drawer=False
  st.rerun()
 eid=st.session_state.ex_edge
 if eid and eid in {e['edge_id'] for e in edges}:
  e=g.edge_by_id[eid]
  with st.container(border=True):
   st.write(g.by_id[e['source']]['name']+' → '+g.by_id[e['target']]['name']);st.caption(e['edge_id']+' · '+REL_LABELS[e['relation_code']]+' · '+e['support_function_or_subtype']);st.write(e['reason']);st.caption(e['review_status'])
   if st.button('查看完整证据'):st.session_state.ex_drawer=True;st.rerun()
 elif selected in {n['node_id'] for n in nodes}:
  n=g.by_id[selected]
  with st.container(border=True):
   st.subheader(n['original_id']+' · '+n['name']);st.write(n.get('description') or '暂无明确说明');st.caption(n['evidence_state']+' · '+n['review_status'])
   if n['node_type']=='COMPETENCY_L3':
    for kind in ['岗位关联','课程关联']:
     if st.button(kind,key='open_'+kind):st.session_state.ex_cap=selected;st.session_state.ex_topic=kind;go('专题')
   elif n['node_type']=='WORK_TASK':
    links=[e for e in g.filtered() if e['relation_code']=='R02' and e['source']==selected]
    st.caption('岗位来源线索：'+'、'.join(sorted({e['source_attributes'].get('岗位ID','原始资料未提供') for e in links})))
    st.caption('点击真实关系线可查看判断及完整摘要证据；候选岗位尚未核验。')
   elif n['node_type']=='TEACHING_PROJECT' and space=='课程体系':
    if st.button('展开项目知识'):st.session_state.ex_project=selected;st.session_state.ex_node=None;st.rerun()
   elif n['node_type']=='COURSE_KS':
    projects=[g.by_id[e['target']]['name'] for e in g.edges if e['relation_code']=='R04' and e['source']==selected];st.write('所属项目：'+'、'.join(projects))
    # 精确复用第一轮说明/功能标签，不推断新内容。
    import json
    course=json.loads((Path(root)/'stage4_course_extraction/output/课程教学要素数据库.json').read_text());r=next((r for r in course['knowledge'] if r['知识点ID']==n['original_id']),{})
    st.write('原文说明：'+r.get('原文说明','暂无明确说明'));st.caption('课程功能标签：'+r.get('课程功能标签','原始资料未提供'))
    supports=[e for e in g.filtered() if e['relation_code']=='R03' and e['source']==selected]
    if not supports:st.caption('当前未建立三级岗位候选能力关联。')
    for e in supports:
     if st.button(g.by_id[e['target']]['name']+' · '+e['support_function_or_subtype'],key='reverse_'+e['edge_id']):st.session_state.ex_cap=e['target'];st.session_state.ex_node=e['target'];st.session_state.ex_topic='课程关联';go('专题')
 if st.session_state.ex_drawer and eid in g.edge_by_id:
  with st.container(border=True,height=440):
   st.subheader('来源证据')
   if st.button('关闭证据抽屉'):st.session_state.ex_drawer=False;st.rerun()
   e=g.edge_by_id[eid];fields=relation_fields(g,e)
   for heading,keys in [('关系概览',['关系编号','关系类型','起点','终点','关系状态']),('研究判断',['判断说明','适用边界','是否研究推断','人工审查状态'])]:
    st.markdown('**'+heading+'**')
    for k in keys:st.write(k+'：'+str(fields[k]))
   for ref in e['evidence_ids']:
    r=g.evidence_by_id[ref];st.caption(r['system']+' · '+ref+' · 页码/记录 '+str(r.get('page') or r['source_record_id']));st.text(r['quote']);st.caption(r['source_file']+' · '+r['state'])
   if not e['evidence_ids']:st.caption('无单独原文引用；来源及判断：'+e['source_file']+'；'+e['reason'])
