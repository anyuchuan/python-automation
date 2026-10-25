"""新候选专业能力页面，复用既有离线交互组件。"""
from collections import Counter
import json
from pathlib import Path
import streamlit as st
from app.competency_store import GraphStore,GraphDataError
from app.visuals import visual
from app.competency_display import NODE_LABELS,REL_LABELS,directed,relation_fields
COLORS={'COMPETENCY_L1':'#195c4c','COMPETENCY_L2':'#338b71','COMPETENCY_L3':'#69a58c','COURSE_KS':'#477ca7','WORK_TASK':'#ad8654','TEACHING_PROJECT':'#8a71ac','TEACHING_ACTIVITY':'#ba8196','PRACTICAL_TASK':'#d19a58','LEARNING_OUTPUT':'#659fa6','ASSESSMENT_EVIDENCE':'#92936b'}
@st.cache_resource
def cached(root,signature):return GraphStore(root)
def render(root):
 st.title('专业能力图谱')
 st.caption('本图谱为基于岗位任务与课程资料构建的候选专业能力图谱，相关支撑关系仍需研究者及专家进一步审查。')
 try:
  folder=Path(root)/'stage4_competency_graph/output';signature=tuple((p.name,p.stat().st_mtime_ns,p.stat().st_size) for p in [folder/(n+'.json') for n in ['graph_nodes','graph_edges','graph_evidence','graph_review_items','graph_paths']]);g=cached(str(root),signature)
 except (GraphDataError,OSError) as exc:st.error(str(exc));return
 c=Counter(n['node_type'] for n in g.nodes);st.caption(' · '.join(f'{name} {value}' for name,value in [('一级候选',c['COMPETENCY_L1']),('二级候选',c['COMPETENCY_L2']),('三级候选',c['COMPETENCY_L3']),('知识技能',c['COURSE_KS']),('节点',len(g.nodes)),('默认候选边',len(g.filtered())),('审查问题',len(g.reviews))]))
 st.info('先搜索或展开能力树，再切换局部网络；点击节点查看详情，点击边查看原文。K1—K4是支撑功能，不是分数或达成等级。')
 with st.sidebar:
  explore=st.checkbox('研究审查模式：包括默认隐藏和排除记录',key='cg_explore')
  if explore:st.warning('研究审查模式：含未核验、待审和排除记录，不能作为有效确认关系。')
  query=st.text_input('搜索能力编号或名称',key='cg_query')
  l1=st.selectbox('一级能力',key='cg_l1',options=['全部']+[n['node_id'] for n in g.search(types=['COMPETENCY_L1'])],format_func=lambda x:g.by_id[x]['name'] if x in g.by_id else x)
  l2=st.selectbox('二级能力',key='cg_l2',options=['全部']+[n['node_id'] for n in g.search(types=['COMPETENCY_L2']) if l1=='全部' or l1 in g.parents(n['node_id'])],format_func=lambda x:g.by_id[x]['name'] if x in g.by_id else x)
  if st.button('重置筛选'):
   for k in ['cg_query','cg_explore','cg_l1','cg_l2','cg_codes','cg_functions','cg_statuses','cg_types']:st.session_state.pop(k,None)
   st.rerun()
  with st.expander('高级筛选：类型、关系与状态'):
   codes=st.multiselect('关系类型',sorted(REL_LABELS),format_func=lambda x:REL_LABELS[x],key='cg_codes')
   functions=st.multiselect('K1—K4支撑功能',['K1','K2','K3','K4'],key='cg_functions')
   statuses=st.multiselect('关系审查状态',sorted({e['review_status'] for e in g.edges}),key='cg_statuses')
   types=st.multiselect('节点类型',list(c),default=list(c),format_func=lambda x:NODE_LABELS[x],key='cg_types')

 options=[n['node_id'] for n in g.search(query,['COMPETENCY_L3']) if (l1=='全部' or l1 in g.parents(n['node_id'])) and (l2=='全部' or l2 in g.parents(n['node_id']))]
 if not options:st.warning('当前筛选无能力，请调整条件。');return
 if st.session_state.get('cg_cap') not in options:st.session_state.cg_cap=options[0]
 cap=st.selectbox('三级候选能力',options,format_func=lambda x:g.by_id[x]['original_id']+' · '+g.by_id[x]['name'],key='cg_cap')
 if 'cg_detail' not in st.session_state:st.session_state.cg_detail=cap
 if st.session_state.get('cg_lastcap')!=cap:st.session_state.cg_detail=cap;st.session_state.cg_center=cap;st.session_state.cg_lastcap=cap
 left,right=st.columns([2,1]);edges=g.filtered(explore,codes,functions,statuses)
 with left:
  tab=st.radio('查询视图',['能力层级','局部关系网络','知识技能反向查询','教学关联路径'],horizontal=True,key='cg_view')
  event=None
  if tab=='能力层级':
   parents={e['target']:e['source'] for e in g.edges if e['relation_code']=='R01'}
   tree=[{'id':n['node_id'],'name':n['name'],'level':{'COMPETENCY_L1':'一级','COMPETENCY_L2':'二级','COMPETENCY_L3':'三级'}[n['node_type']],'parent':parents.get(n['node_id'])} for n in g.nodes if n['node_type'].startswith('COMPETENCY') and (l1=='全部' or n['node_id']==l1 or l1 in g.parents(n['node_id'])) and (l2=='全部' or n['node_id']==l2 or l2 in g.parents(n['node_id']) or n['node_id'] in g.parents(l2))]
   event=visual(mode='tree',tree=tree,expanded=st.session_state.get('cg_expanded',[]),selected_cap=cap,locate=True,key='cg_tree')
  elif tab=='局部关系网络':
   depth=st.slider('局部展开层数',1,3,1);center=st.session_state.get('cg_center',cap);ns,es=g.subgraph(center,edges,types,depth)
   payload={'nodes':[{'id':n['node_id'],'name':n['name'],'type':n['node_type'],'lane':0 if n['node_type']=='WORK_TASK' else 1 if n['node_type']=='COURSE_KS' else 2 if n['node_type'].startswith('COMPETENCY') else 3 if n['node_type']=='TEACHING_PROJECT' else 4,'code':n['original_id'],'capability':n['node_type'].startswith('COMPETENCY'),'color':COLORS[n['node_type']]} for n in ns],'edges':[{'id':e['edge_id'],'from':e['source'],'to':e['target'],'type':e['relation_type'],'color':'#ba725c' if not e['default_visible'] else '#648b82','dash':'5 4' if e['research_inference'] else '','directed':directed(e),'label':REL_LABELS[e['relation_code']]} for e in es]}
   event=visual(mode='network',candidate_graph=True,graph=payload,center=center,selected_node=st.session_state.cg_detail,key='cg_network')
   st.caption('实线：来源结构关联；虚线：研究推断；棕色：默认隐藏记录。所有连线均保留候选/待审状态。节点颜色按类型区分。')
   st.dataframe([{'类型':NODE_LABELS[k],'颜色':v} for k,v in COLORS.items()],hide_index=True)
  elif tab=='知识技能反向查询':
   q=st.text_input('CK编号或知识名称');ks=g.search(q,['COURSE_KS'])
   if ks:
    kid=st.selectbox('课程知识技能点',[n['node_id'] for n in ks],format_func=lambda x:g.by_id[x]['original_id']+' · '+g.by_id[x]['name']);st.session_state.cg_detail=kid
    rel=[e for e in edges if e['source']==kid and e['relation_code'] in ['R03','R04']]
    if not any(e['relation_code']=='R03' for e in rel):st.info('当前未建立三级岗位候选能力关联。筛选可能隐藏已有关系，请同时检查筛选条件。')
    st.dataframe(rel,hide_index=True)
   else:st.info('无匹配知识点。')
  else:
   ps=g.teaching_paths(cap,explore);st.caption('基于知识点路径推导的候选关联，不是课程原文明示的直接能力培养或正式考核。');st.dataframe(ps,hide_index=True)
  if tab=='局部关系网络':
   with st.expander('关联路径说明（图谱路径计算结果）'):
    st.write('中心：'+g.by_id[center]['original_id']+' · '+g.by_id[center]['name'])
    target=st.selectbox('解释节点出现原因',[n['node_id'] for n in ns],format_func=lambda x:g.by_id[x]['original_id']+' · '+g.by_id[x]['name'])
    chain=g.path_to(center,target,es)
    if chain is None:st.warning('没有有效路径，需检查展示或数据。')
    elif not chain:st.caption('当前中心节点。')
    else:
     st.caption(('直接关联' if len(chain)==1 else '多跳关联')+'；邻接遍历不改变边的方向，非层级边不表示下级能力。')
     for eid in chain:st.write(relation_fields(g,g.edge_by_id[eid]))
  if event and event.get('nonce')!=st.session_state.get('cg_nonce'):
   st.session_state.cg_nonce=event.get('nonce');action=event.get('action');id=event.get('id')
   if action=='expand':st.session_state.cg_expanded=event.get('expanded',[])
   elif action=='edge':st.session_state.cg_edge=id
   elif id in g.by_id:
    st.session_state.cg_detail=id
    if action=='local_expand':st.session_state.cg_center=id
    if action in ['select','tree_locate']:
     if g.by_id[id]['node_type']=='COMPETENCY_L3':st.session_state.cg_pending=id
     if action=='tree_locate':st.session_state.cg_pending_view='能力层级'
   st.rerun()
  related=g.related(cap,explore);support=[e for e in related if e['relation_code']=='R03']
  if not support:st.info('当前尚未建立充分依据的课程知识技能支撑候选关系。')
  with st.expander('岗位任务依据及A/B/C说明'):
   st.caption('A：候选岗位摘要明确技术活动；B：数据库相关但具体技术未明确；C：业务任务的关系数据库实现推断。均不是已核验招聘原文，不代表频率或重要性。')
   st.dataframe([e for e in related if e['relation_code']=='R02'],hide_index=True)
  with st.expander('课程知识支撑及K1—K4'):
   st.caption('K1基础知识、K2直接技术操作、K3综合应用、K4条件性支撑；不等于完整能力培养。');st.dataframe(support,hide_index=True)
 with right:
  id=st.session_state.cg_detail;n=g.by_id[id];st.subheader(n['original_id']+' · '+n['name']);st.write(n['description']);st.caption(n['source_stage']+' · '+n['review_status'])
  st.write('能力归属：'+' → '.join(g.by_id[x]['name'] for x in reversed(g.parents(id))) if n['node_type'].startswith('COMPETENCY') else '课程/教学要素节点')
  behavior=next((e['source_attributes'].get('competency_behavior') for e in g.edges if e['relation_code']=='R03' and e['target']==id),None)
  st.write('关键技术行为：'+(behavior or '原始资料未提供'))
  rel=g.related(id,explore)
  ids=[e['edge_id'] for e in rel]
  clicked=st.session_state.get('cg_edge')
  if clicked in {e['edge_id'] for e in edges} and clicked not in ids:ids.append(clicked)
  if ids:
   selected=st.session_state.get('cg_edge');eid=st.selectbox('关系与证据',[*ids],index=ids.index(selected) if selected in ids else 0);e=g.edge_by_id[eid]
   for label,value in relation_fields(g,e).items():st.markdown('**'+label+'**');st.write(value)
   with st.expander('高级数据查看（原始JSON）'):st.json(e,expanded=False)
   for ref in e['evidence_ids']:
    item=g.evidence_by_id[ref];st.caption(item['system']+' · '+ref+' · PDF页 '+str(item['page'])+' · '+item['state']);st.text(item['quote']);st.caption(item['source_file'])
  with st.expander('相关人工审查问题'):
   for q in g.review_for(id):
    st.write({'问题编号':q['review_id'],'类别':q.get('review_category','原始资料未提供'),'描述':q['issue'],'状态':q['state'],'建议或待确认':q.get('handling_suggestion',q.get('restriction','原始资料未提供'))})
 st.download_button('下载当前筛选关系JSON',json.dumps(edges,ensure_ascii=False,indent=2),'filtered_graph_edges.json','application/json')
