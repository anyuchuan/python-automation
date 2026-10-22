"""课程能力图谱与教学优化分析平台；只读前三阶段成果。"""
import html
import platform
import sys
from collections import Counter
from pathlib import Path
from time import perf_counter
ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:sys.path.insert(0,str(ROOT))
import altair as alt
import pandas as pd
import streamlit as st
from streamlit.logger import get_logger
from app.data import DataError,Options,MAIN_TYPES,STATUS_ORDER,load_bundle,visible_data,csv_bytes,detail
from app.platform_model import (MODULES,hierarchy,ancestors,find_capabilities,global_counts,evidence_group,network_slice,
                                capability_detail,project_cell,cell_evidence,optimization_rows,split_ids)
from app.controller import choose,apply_event
from app.visuals import visual,graph_payload
from app.graph import STATUS_COLORS

st.set_page_config(page_title='课程能力图谱与教学优化分析平台',page_icon='◈',layout='wide')
st.markdown('''<style>
.stApp{background:#f7f9f7;color:#24444b}[data-testid="stHeader"]{background:#f7f9f7e8}.block-container{padding-top:2.5rem;padding-bottom:2rem;max-width:1550px}[data-testid="stSidebar"]{background:#eaf0ed;border-right:1px solid #dfe7e2}h1{font-size:26px!important;letter-spacing:-.025em;color:#23463e;line-height:1.45!important}h2{font-size:20px!important;color:#2b5048}h3{font-size:16px!important;color:#355b52}.brand{font-size:11px;letter-spacing:.15em;color:#7d948a;font-weight:650}.brand-title{font-size:20px;color:#294d42;font-weight:650;margin:8px 0}.eyebrow{font:11px ui-monospace,monospace;color:#78938a;letter-spacing:.15em}.intro{font-size:13px;color:#7a908a;line-height:1.9;margin:0 0 15px}.scope{font-size:11px;line-height:1.9;color:#6b8980;background:#edf4ef;border:1px solid #dce9df;padding:9px 12px;border-radius:7px;margin:12px 0}.detail-head{font-size:11px;letter-spacing:.1em;color:#8b9d94;margin:12px 0 8px}.detail-name{font-size:19px;color:#254e43;font-weight:650;line-height:1.5}.code{font:12px ui-monospace,monospace;color:#428673}.stMetric{background:white;border:1px solid #e0e8e2;padding:12px;border-radius:9px}[data-testid="stMetricValue"]{font-size:25px}[data-testid="stMetricLabel"]{font-size:11px}.stButton>button,.stDownloadButton>button{border-radius:7px;font-size:12px}[data-testid="stAlert"]{border-radius:8px}.stCaption{line-height:1.8}button[kind="primary"]{background:#15755f;border-color:#15755f}div[data-testid="stVerticalBlockBorderWrapper"]{border-radius:10px}
</style>''',unsafe_allow_html=True)

@st.cache_data(show_spinner='正在读取第三阶段已生成的图谱数据…',max_entries=2)
def cached_bundle(signature):
    started=perf_counter();get_logger('course_graph').info('Course graph: loading existing Excel on Python %s; no research pipeline',platform.python_version())
    data=load_bundle(ROOT);hierarchy(data)
    get_logger('course_graph').info('Course graph data ready: %s nodes, %s candidate edges, %s isolated edges in %.3fs',len(data.nodes),len(data.edges),len(data.isolated),perf_counter()-started)
    return data

def load():
    try:
        paths=[ROOT/'phase3/output'/name for name in ['知识图谱节点表.xlsx','知识图谱关系表.xlsx','职业能力课程匹配矩阵.xlsx','课程能力覆盖与缺口分析.xlsx']]+[ROOT/'app/input_snapshot.json']
        return cached_bundle(tuple((str(p),p.stat().st_mtime_ns,p.stat().st_size) for p in paths))
    except (DataError,FileNotFoundError) as exc:st.error('未能加载修订版数据：'+str(exc));st.stop()

bundle=load();tree=hierarchy(bundle);counts=global_counts(bundle)
for key,value in {'module':MODULES[0],'selected_cap':'N1','tree_expanded':[],'locate':False,'graph_view':'树状层级','net_center':'N1','net_selected':'N1','selected_project':'P01','selected_edge':None}.items():
    if key not in st.session_state:st.session_state[key]=value
if st.session_state.get('pending_event'):
    apply_event(st.session_state,st.session_state.pop('pending_event'),tree,bundle)

def process_event(event,key):
    if isinstance(event,dict) and event.get('nonce') and event['nonce']!=st.session_state.get('seen_'+key):
        st.session_state['seen_'+key]=event['nonce'];st.session_state.pending_event=event;st.rerun()

def navigate(module):st.session_state.module=module

def select_cap():choose(st.session_state,tree,st.session_state.cap_picker)

def network_select():choose(st.session_state,tree,st.session_state.network_cap)

def locate_search():
    matches=find_capabilities(tree,st.session_state.cap_query)
    if st.session_state.cap_query.strip() and matches:choose(st.session_state,tree,matches[0])

def cap_select(label='选择三级能力'):
    current=st.session_state.selected_cap
    ids=list(bundle.cap_by_id)
    if current not in ids:
        current='N113';choose(st.session_state,tree,current)
    if st.session_state.get('cap_picker')!=current:st.session_state.cap_picker=current
    return st.selectbox(label,ids,format_func=lambda id:id+' · '+tree[id]['name'],key='cap_picker',on_change=select_cap)

def heading(kicker,title,note):
    st.markdown('<div class="eyebrow">'+html.escape(kicker)+'</div>',unsafe_allow_html=True);st.title(title);st.markdown('<div class="intro">'+html.escape(note)+'</div>',unsafe_allow_html=True)

def frame_table(records,columns=None,height=None):
    frame=pd.DataFrame(records)
    if columns:frame=frame.reindex(columns=columns)
    kwargs={'hide_index':True,'width':'stretch'}
    if height:kwargs['height']=height
    st.dataframe(frame,**kwargs)

def export(records,label,name,key):st.download_button(label,csv_bytes(records),name,'text/csv',key=key,disabled=not records,width='stretch')

def evidence_panel(evidence):
    if not evidence:st.caption('暂无记录：本单元没有可引用的正向课程证据。');return
    ids=[e['证据ID'] for e in evidence]
    key='quote_'+st.session_state.module
    # 同一能力的多项目证据均保留；默认仅展开一条原文。
    eid=st.selectbox('课程标准摘录',ids,format_func=lambda id:id+' · PDF物理页'+str(bundle.evidence_by_id[id]['PDF页码']),key=key)
    e=bundle.evidence_by_id[eid];st.caption(str(e['课程证据状态'])+' · '+str(e['人工审核状态']))
    st.text_area('课程标准原文（保留原摘录）',str(e['原文摘录']),height=130,disabled=True,key='quote_text_'+key)

def right_detail(cap=None,cell=None,options=Options()):
    id=cap or st.session_state.selected_cap;d=capability_detail(bundle,id)
    with st.container(border=True,height=690,key='ability_details'):
        st.markdown('<div class="detail-head">ABILITY / 能力与证据</div>',unsafe_allow_html=True)
        if not d:st.write('请选择一个能力节点。');return
        st.markdown('<div class="code">'+html.escape(id)+' · '+html.escape(d['level'])+'</div><div class="detail-name">'+html.escape(d['name'])+'</div>',unsafe_allow_html=True)
        st.caption('所属上位编码：'+d['parent']);st.write('**能力层面文本判断：** '+d['grade'])
        st.caption(d['state']+' · '+d['review'])
        with st.expander('定义与可观察行为',expanded=True):
            st.caption('能力定义');st.write(d['definition']);st.caption('可观察工作行为（既有研究表述）');st.write(d['behavior'])
        if id in bundle.cap_by_id:
            with st.expander('课程支撑行为与边界',expanded=True):
                st.caption('课程支撑的具体技术行为');st.write(d['technical']);st.caption('课程尚未明确覆盖的工作行为');st.write(d['uncovered'])
        if cell:
            st.divider();st.markdown('**当前项目单元 · '+str(cell['课程项目ID'])+'**');st.write('项目层判断：'+cell['项目层匹配等级'])
            st.caption('匹配理由（研究判断）');st.write(cell.get('等级说明') or cell.get('匹配理由') or '暂无记录')
            st.caption('知识点：'+str(cell.get('知识点ID') or '暂无记录')+' · 实践任务：'+str(cell.get('实践任务ID') or '暂无记录'))
            with st.expander('项目单元完整记录（含查阅范围）'):st.json(cell,expanded=False)
            evidence_panel(cell_evidence(bundle,cell))
        else:
            ev_ids=set()
            for m in bundle.matches:
                if m['能力ID']==id:ev_ids.update(split_ids(m.get('课程证据ID')))
            with st.expander('课程标准证据与页码'):
                evidence_panel([bundle.evidence_by_id[x] for x in sorted(ev_ids) if x in bundle.evidence_by_id])
        _,visible_edges=visible_data(bundle,options)
        task_edges=[e for e in visible_edges if e['关系类型']=='任务候选关联新能力' and e['终点ID']==id]
        task_ids=sorted({e['起点ID'] for e in task_edges});job_ids=sorted({bundle.attrs[t].get('岗位ID') for t in task_ids if bundle.attrs[t].get('岗位ID')})
        with st.expander('关联候选岗位与典型任务'):
            st.caption(f'当前展示规则下：{len(job_ids)}候选岗位、{len(task_ids)}摘要任务；不是全量统计或正式需求频次。')
            st.caption('历史来源字段共指向'+str(len(d['jobs']))+'岗位/'+str(len(d['tasks']))+'任务，包含研究推断；隐含/高风险关系默认不展开。')
            if task_ids:frame_table([{'任务ID':t,'典型任务':bundle.by_id[t]['节点名称'],'岗位ID':bundle.attrs[t].get('岗位ID'),'证据':'摘要证据；待核验'} for t in task_ids])
            else:st.write('暂无记录（当前展示规则下）。')
            if job_ids:frame_table([{'岗位ID':j,'岗位名称':bundle.by_id[j]['节点名称'],'状态':bundle.attrs[j].get('核验状态') or '暂无记录'} for j in job_ids])
        with st.expander('关联教学项目'):
            if d['project_cells']:frame_table([{'项目':r['课程项目ID'],'名称':bundle.by_id[r['课程项目ID']]['节点名称'],'项目层判断':r['项目层匹配等级']} for r in d['project_cells']])
            else:st.write('暂无记录；未见明确支撑不等于课程绝对不涉及。')
        st.caption('证据类型：'+d['evidence']);st.caption('来源：'+str(d['source']))
        st.button('查看能力—教学内容映射',on_click=navigate,args=(MODULES[2],),key='to_mapping',width='stretch')
        st.button('查看课程支撑评价',on_click=navigate,args=(MODULES[3],),key='to_evaluation',width='stretch')

with st.sidebar:
    st.markdown('<div class="brand">CURRICULUM RESEARCH</div><div class="brand-title">课程能力与教学优化</div>',unsafe_allow_html=True)
    st.caption('《数据库原理及应用》 · 高职大数据技术专业')
    module=st.radio('研究模块',MODULES,key='module',label_visibility='collapsed')
    st.divider();st.caption('研究边界')
    st.caption('以旅游行业候选岗位为依据，以课程能力为核心，服务教学内容映射与优化分析。')
    st.caption('候选样本与研究推断保留原状态；文本支撑不是学生能力达成。')
    with st.expander('研究探索选项'):
        explore=st.checkbox('研究探索模式',key='explore')
        if not explore:st.session_state.high_risk=False;st.session_state.isolated=False
        high=st.checkbox('显示高风险推断（未核验）',key='high_risk',disabled=not explore)
        iso=st.checkbox('显示隔离关系（未确认）',key='isolated',disabled=not explore)
        st.caption('探索只改变可见性，不升级证据。')
options=Options(explore,high,iso)
if explore and (high or iso):st.warning('研究探索已开启：显示内容可能包含高风险推断或隔离记录，均未核验/未确认。')

if module==MODULES[0]:
    heading('01 / ABILITY STRUCTURE','《数据库原理及应用》课程能力图谱与教学优化分析平台','从能力层级出发，连接候选岗位任务、课程教学内容与可追溯的优化线索。')
    view=st.radio('图谱视图',['树状层级','网络关系'],horizontal=True,key='graph_view')
    left,right=st.columns([2.15,1],gap='large')
    with left:
        if view=='树状层级':
            st.text_input('搜索能力编号或名称',placeholder='输入 N113、备份、业务数据…，回车定位',key='cap_query',on_change=locate_search)
            matches=find_capabilities(tree,st.session_state.get('cap_query',''))
            if st.session_state.get('cap_query','').strip():
                if matches:
                    target=st.selectbox('搜索结果',matches,format_func=lambda id:id+' · '+tree[id]['name'],key='tree_result')
                    st.button('定位所选能力',on_click=choose,args=(st.session_state,tree,target),key='locate_result')
                else:st.info('没有匹配能力；未生成占位节点。')
            event=visual(mode='tree',tree=[{**n,'grade':bundle.cap_by_id.get(id,{}).get('最终候选匹配状态')} for id,n in tree.items()],selected_cap=st.session_state.selected_cap,expanded=st.session_state.tree_expanded,locate=st.session_state.locate,key='ability_visual')
            process_event(event,'ability_visual')
            st.caption('首次使用：点击＋逐级展开，点击名称查看右侧依据；选中三级能力后切换网络视图，中心能力保持不变。')
        else:
            current=st.session_state.selected_cap
            c1,c2=st.columns([3,1])
            with c1:
                if st.session_state.get('network_cap')!=current:st.session_state.network_cap=current
                st.selectbox('能力中心',list(tree),format_func=lambda id:id+' · '+tree[id]['name'],key='network_cap',on_change=network_select)
            with c2:st.caption('选择后自动定位')
            center=st.session_state.net_center
            with st.expander('节点、关系与证据筛选',expanded=False):
                types=st.multiselect('节点类型',sorted({n['节点类型'] for n in bundle.nodes}),default=list(MAIN_TYPES),key='net_types')
                _,all_edges=visible_data(bundle,Options(True,True,True))
                rels=sorted({e['关系类型'] for e in all_edges});groups=sorted({evidence_group(e) for e in all_edges})
                rel=st.multiselect('关系类型',rels,default=rels,key='net_relations')
                evid=st.multiselect('证据状态组',groups,default=groups,key='net_evidence')
                a,b=st.columns(2)
                with a:hops=st.slider('局部邻域深度',1,3,2,key='net_hops')
                with b:limit=st.slider('绘制节点上限',20,100,60,10,key='net_limit')
            nodes,edges,total=network_slice(bundle,center,options,types,rel,evid,hops,limit)
            st.caption(f'当前子图（非全量）：{len(nodes)}节点 / {len(edges)}关系；中心 {center}。'+(f'匹配{total}节点，已按上限截断。' if total>len(nodes) else ''))
            event=visual(mode='network',graph=graph_payload(bundle,nodes,edges),center=center,selected_node=st.session_state.net_selected,key='ability_visual');process_event(event,'ability_visual')
            a,b=st.columns(2)
            with a:export(nodes,'导出当前子图节点','当前子图节点.csv','graph_nodes')
            with b:export(edges,'导出当前子图关系','当前子图关系.csv','graph_edges')
            e=next((e for e in edges if e['关系ID']==st.session_state.selected_edge),None)
            if e:
                with st.expander('选中关系的证据',expanded=True):
                    st.write(e['关系ID']+' · '+e['关系类型']);st.caption(str(e['审核或置信状态']))
                    st.write(e.get('课程标准原文摘录') or e.get('岗位摘要证据') or '暂无记录')
                    st.caption('来源：'+str(e['来源文件'])+' · 原记录 '+str(e['源记录ID'])+' · PDF页码 '+str(e.get('课程PDF页码') or '暂无记录'))
            selected_node=st.session_state.net_selected
            if selected_node in bundle.by_id and selected_node not in tree:
                with st.expander('选中关联节点详情',expanded=True):st.json(detail(bundle,selected_node,edges),expanded=False)
        st.markdown('<div class="scope">全量研究数据：50候选岗位 · 79摘要任务 · 34三级能力 · 8教学项目。树含5/14/34项能力；局部子图计数始终单独标注。</div>',unsafe_allow_html=True)
    with right:right_detail(options=options)

elif module==MODULES[1]:
    heading('02 / CANDIDATE DEMAND','岗位能力需求分析','沿着已有候选关系查看岗位摘要与任务，保留研究推断边界；不计算正式行业需求频次。')
    st.caption('全量：50候选岗位 / 79摘要任务 / 34三级能力；下方筛选和局部计数不能替代全量。')
    left,right=st.columns([2.15,1],gap='large')
    with left:
        cap=cap_select();_,edges=visible_data(bundle,options)
        links=[e for e in edges if e['关系类型']=='任务候选关联新能力' and e['终点ID']==cap]
        query=st.text_input('检索岗位、任务或证据',key='demand_query')
        records=[]
        for e in links:
            task=e['起点ID'];a=bundle.attrs[task];job=a.get('岗位ID');j=bundle.attrs.get(job,{})
            record={'能力ID':cap,'任务ID':task,'典型任务':bundle.by_id[task]['节点名称'],'岗位ID':job,'候选岗位':j.get('岗位名称') or '暂无记录','招聘单位':j.get('招聘单位') or '暂无记录','证据文本':a.get('证据文本') or '暂无记录','证据类型':e.get('招聘证据类型'),'关系状态':e['审核或置信状态'],'岗位状态':j.get('核验状态') or '暂无记录','原始招聘链接':j.get('原始招聘链接') or '暂无记录','原关系ID':e['源记录ID'],'风险说明':e.get('展示风险说明') or '默认候选展示；非已确认'}
            if not query.strip() or query.casefold() in str(record).casefold():records.append(record)
        st.caption(f'当前筛选：{len({r["岗位ID"] for r in records})}候选岗位 / {len({r["任务ID"] for r in records})}任务 / {len(records)}条关系；含疑似重复可能，非独立需求频次。')
        frame_table(records,['任务ID','典型任务','岗位ID','候选岗位','证据类型','关系状态'],height=340)
        if records:
            selected=st.selectbox('查看候选记录证据',range(len(records)),format_func=lambda i:records[i]['任务ID']+' · '+records[i]['候选岗位'])
            r=records[selected];st.write(r['证据文本']);st.caption('摘要证据，非招聘原文；'+str(r['岗位状态']));st.write('原始招聘链接：'+str(r['原始招聘链接']))
        else:st.info('当前展示规则下暂无关联；可切换能力或主动查看研究探索选项。')
        export(records,'导出当前候选关联','候选任务能力关联.csv','demand_csv')
    with right:right_detail(cap,options=options)

elif module==MODULES[2]:
    heading('03 / CONTENT MAPPING','能力—教学内容映射','点击能力×项目单元，查看原判断、理由及课程引文。项目单元等级与能力整体等级分别呈现。')
    left,right=st.columns([2.15,1],gap='large')
    with left:
        query=st.text_input('筛选能力编号或名称',key='matrix_query')
        grades=st.multiselect('能力整体支撑等级（行筛选）',list(STATUS_ORDER),default=list(STATUS_ORDER),key='matrix_grades')
        projects=st.multiselect('教学项目（列筛选）',list(bundle.project_by_id),default=list(bundle.project_by_id),format_func=lambda id:id+' · '+bundle.by_id[id]['节点名称'],key='matrix_columns')
        caps=[id for id in bundle.cap_by_id if id in find_capabilities(tree,query) and bundle.cap_by_id[id]['最终候选匹配状态'] in grades]
        cells=[r for r in bundle.project_cells if r['能力ID'] in caps and r['课程项目ID'] in projects]
        st.caption(f'完整矩阵34×8＝272单元；当前显示{len(caps)}×{len(projects)}＝{len(cells)}单元。行筛选使用能力等级，单元颜色使用项目等级。')
        event=visual(mode='matrix',abilities=[tree[id] for id in caps],projects=[{'id':id,'name':bundle.by_id[id]['节点名称']} for id in projects],cells=[{'cap':r['能力ID'],'project':r['课程项目ID'],'grade':r['项目层匹配等级'],'reason':r.get('等级说明') or r.get('匹配理由') or '暂无记录'} for r in cells],selected_cap=st.session_state.selected_cap,selected_project=st.session_state.selected_project,key='mapping_matrix');process_event(event,'mapping_matrix')
        export(cells,'导出当前映射单元','能力教学项目映射.csv','matrix_csv')
        with st.expander('项目编码与名称'):frame_table([{'项目ID':id,'教学项目':bundle.by_id[id]['节点名称'],'总学时':bundle.attrs[id].get('总学时'),'内含实践学时':bundle.attrs[id].get('内含实践学时')} for id in projects]);st.caption('内含实践已经计入项目总学时，不另行叠加32学时口径。')
        cap=cap_select('通过编号查看单元（也可直接点击矩阵）')
        project=st.selectbox('查看教学项目单元',list(bundle.project_by_id),index=list(bundle.project_by_id).index(st.session_state.selected_project),format_func=lambda id:id+' · '+bundle.by_id[id]['节点名称'],key='matrix_project')
        st.session_state.selected_project=project
    with right:right_detail(cap,project_cell(bundle,cap,project),options)

elif module==MODULES[3]:
    heading('04 / TEXT SUPPORT EVALUATION','课程能力支撑评价','评价对象为课程标准文本与候选职业能力的研究匹配；不是学生能力达成率或经核验行业需求覆盖率。')
    left,right=st.columns([2.15,1],gap='large')
    with left:
        counts_by=Counter(c['最终候选匹配状态'] for c in bundle.capabilities)
        cols=st.columns(4)
        for col,status in zip(cols,STATUS_ORDER):col.metric(status,f'{counts_by[status]} / 34',f'{counts_by[status]/34:.2%}',delta_color='off')
        chart_data=pd.DataFrame([{'支撑等级':s,'数量':counts_by[s]} for s in STATUS_ORDER])
        chart=alt.Chart(chart_data).mark_bar(cornerRadiusEnd=4,size=25).encode(x=alt.X('数量:Q',title='能力数量（分母34）',axis=alt.Axis(tickMinStep=1)),y=alt.Y('支撑等级:N',sort=list(STATUS_ORDER),title=None),color=alt.Color('支撑等级:N',scale=alt.Scale(domain=list(STATUS_ORDER),range=[STATUS_COLORS[s] for s in STATUS_ORDER]),legend=None),tooltip=['支撑等级','数量']).properties(height=180)
        st.altair_chart(chart,use_container_width=True)
        st.caption('直接或部分文本对应19/34，深度不同，不能表述为19项能力已完整培养。')
        selected=st.multiselect('筛选能力等级',list(STATUS_ORDER),default=list(STATUS_ORDER),key='evaluation_grades')
        rows=[r for r in bundle.capabilities if r['最终候选匹配状态'] in selected]
        frame_table(rows,['能力ID','能力名称','最终候选匹配状态','课程支撑的具体技术行为','课程尚未明确覆盖的工作行为'])
        export(rows,'导出当前能力评价','课程能力文本评价.csv','evaluation_csv')
        with st.expander('能力×教学项目单元统计（独立口径）'):
            counter=Counter(r['项目层匹配等级'] for r in bundle.project_cells)
            frame_table([{'项目单元判断':s,'单元数':counter[s],'分母':272,'比例':f'{counter[s]/272:.2%}'} for s in STATUS_ORDER]);st.caption('272项目单元不是272项能力；不得与34项能力分母混用。')
        cap=cap_select()
    with right:right_detail(cap,options=options)

elif module==MODULES[4]:
    heading('05 / TEACHING OPTIMIZATION','课程教学优化','从既有行为边界与建议分类中提取优化线索，分别呈现研究发现、建议与待核实事项；不自动认定课程缺失。')
    left,right=st.columns([2.15,1],gap='large')
    with left:
        rows=optimization_rows(bundle);categories=sorted({r['建议分类'] for r in rows})
        category=st.multiselect('建议分类',categories,default=categories,key='optimization_categories')
        filtered=[r for r in rows if r['建议分类'] in category]
        frame_table(filtered,['能力ID','能力名称','能力层面文本判断','建议分类','研究发现'])
        st.caption('14项未见明确支撑：4项潜在改进、5项其他课程可能承担、4项旅游业务背景、1项证据不足。分类均为研究建议。')
        cap=cap_select();r=next(r for r in rows if r['能力ID']==cap)
        for label in ['研究发现','建议分类','建议依据（研究判断）','待核实事项']:
            st.markdown('**'+label+'**');st.write(r[label])
        st.caption('上述待核实事项为教学边界判断，不新增招聘逐条核验任务，不作为进入后续阶段前置条件。')
        export(filtered,'导出当前优化线索','教学优化线索.csv','optimization_csv')
    with right:right_detail(cap,options=options)

else:
    heading('06 / RESEARCH PROVENANCE','研究依据与数据管理','只读展示来源、状态和统计口径。探索、筛选与导出不会更新候选岗位、推断关系或课程证据的研究状态。')
    st.write('**全量来源数据**：'+' · '.join(str(v)+' '+k for k,v in counts.items()))
    st.caption('当前页面的全量计数始终由全部输入记录计算，不使用某个能力的局部子图。')
    st.write('研究文件快照：`'+bundle.provenance['main_commit']+'`（PR #3修订版）；平台代码基于已合并PR #5的main。')
    frame_table([{'文件':name,'仓库相对路径':str(p.relative_to(ROOT)),'SHA256':bundle.provenance['loaded_sha256'][p.relative_to(ROOT).as_posix()]} for name,p in bundle.files.items()])
    st.subheader('证据口径与局限')
    st.write('50岗位均为待核验候选样本；任务来自摘要；能力及部分关联为AI辅助研究归纳。109条历史疑似任务能力关系仍保留原状态，不新增核验任务。课程引文有物理页码但未经专家或人工确认。文本支撑不等于教学实施、学生能力达成或已核验行业需求覆盖。')
    st.write('未见明确支撑表示现有课程文本未见明确对应，不证明课程绝对不涉及。潜在改进、其他课程可能承担等均为建议，不伪称课标结论。隔离课程候选不作为已确认正向实训证据。')
    st.subheader('默认与探索视图')
    nodes,edges=visible_data(bundle,options)
    st.caption(f'当前风险展示规则：{len(nodes)}节点 / {len(edges)}关系。默认259/450；主动开启全部探索后314/924（含60隔离）。所有默认关系仍是候选/研究关系。')
    a,b=st.columns(2)
    with a:export(nodes,'下载当前规则节点','研究图谱节点.csv','source_nodes')
    with b:export(edges,'下载当前规则关系','研究图谱关系.csv','source_edges')
    with st.expander('隔离关系目录（仅显示说明，不作为确认关系）'):
        if options.explore and options.isolated:
            frame_table(bundle.isolated,['关系ID','起点ID','终点ID','关系类型','审核或置信状态','课程证据可见性']);export(bundle.isolated,'导出隔离关系原记录','隔离关系_未确认.csv','isolated_csv')
        else:st.caption('共60条隔离关系默认隐藏。需在左侧研究探索选项主动开启隔离显示。')
    with st.expander('局部计数4/2/2的复核说明'):
        st.write('最新main旧首页代码按完整bundle计数，应为50/79/34；未复现4/2/2为全量首页统计。局部能力、关系筛选、绘制上限可产生小计数，亦可能与部署分支/入口或旧缓存有关。缺少现场部署快照，不能确认具体原因。此版明确标注全量、当前筛选、当前子图三种口径，并有测试隔离。')
