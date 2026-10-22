"""启动：streamlit run app/main.py"""
import hashlib
import html
import sys
from collections import Counter
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:sys.path.insert(0,str(ROOT))
import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from app.data import (DataError,Options,MAIN_TYPES,STATUS_ORDER,load_bundle,filter_graph,search_nodes,visible_data,detail,csv_bytes)
from app.graph import build_html,STATUS_COLORS

st.set_page_config(page_title='数据库原理及应用 · 课程能力图谱',page_icon='◈',layout='wide',initial_sidebar_state='expanded')
st.markdown('''<style>
:root{--ink:#233f49;--teal:#117862}html,body,[class*="css"]{font-family:Inter,"PingFang SC","Microsoft YaHei",sans-serif}.stApp{background:#f5f8f7;color:var(--ink)}[data-testid="stHeader"]{background:#f5f8f7e8}.block-container{padding-top:3rem;padding-bottom:2rem;max-width:1580px}h1{font-weight:700!important;letter-spacing:-.035em!important;color:#203b45!important}h2,h3{color:#2b4650!important}[data-testid="stSidebar"]{background:#eaf0ed;border-right:1px solid #dfe7e3}[data-testid="stSidebar"] .block-container{padding-top:1.5rem}.brand{font-size:12px;letter-spacing:.2em;font-weight:700;color:#648176;margin-bottom:8px}.brand-title{font-size:21px;font-weight:700;color:#25463e;margin-bottom:6px}.subtle{color:#7f939b;font-size:12px}.eyebrow{font-size:11px;letter-spacing:.18em;color:#688078;font-weight:600;margin-bottom:10px}.hero-note{font-size:14px;color:#6b838b;max-width:880px;line-height:1.9;margin-bottom:24px}.pill{display:inline-block;background:#e0efe7;border:1px solid #cfe2d7;color:#24765b;padding:5px 10px;border-radius:20px;font-size:11px;letter-spacing:.02em;margin-right:7px}.flow{display:flex;align-items:center;gap:10px;margin:22px 0 28px;flex-wrap:wrap}.flow span{padding:9px 13px;background:white;border:1px solid #e0e8e5;border-radius:9px;font-size:12px;color:#47666a}.flow b{font-weight:400;color:#9caaa7}[data-testid="stMetric"]{padding:18px 18px;background:white;border:1px solid #e0e9e5;border-radius:12px}[data-testid="stMetricLabel"]{font-size:12px;color:#7b9195}[data-testid="stMetricValue"]{font-size:29px;line-height:1.5;font-weight:600;color:#294c47}.stButton>button{border-radius:8px}.caption-card{background:white;border:1px solid #e0e8e4;border-radius:13px;padding:20px 22px;margin-top:16px}.caption-card h4{margin:0 0 7px;font-size:14px;color:#38615b}.caption-card p{margin:0;color:#778f92;font-size:12px;line-height:1.9}[data-testid="stAlert"]{border-radius:9px;font-size:12px}[data-testid="stDataFrame"]{border-radius:10px;overflow:hidden}.section-note{font-size:12px;color:#7d9397;line-height:1.8}.stDownloadButton>button{border-radius:8px}.step{font-size:10px;letter-spacing:.14em;color:#869c95}
</style>''',unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def cached_bundle(signature):
    return load_bundle(ROOT)


def load():
    names=['知识图谱节点表.xlsx','知识图谱关系表.xlsx','职业能力课程匹配矩阵.xlsx','课程能力覆盖与缺口分析.xlsx']
    paths=[ROOT/'phase3/output'/n for n in names]+[ROOT/'app/input_snapshot.json']
    try:
        sig=tuple((str(p),p.stat().st_mtime_ns,p.stat().st_size) for p in paths)
        return cached_bundle(sig)
    except (DataError,FileNotFoundError) as exc:
        st.error('未能加载当前修订版研究数据：'+str(exc));st.stop()


bundle=load()
with st.sidebar:
    st.markdown('<div class="brand">COURSE RESEARCH LAB</div><div class="brand-title">数据库能力图谱</div><div class="subtle">旅游行业岗位需求 · 本地研究展示</div>',unsafe_allow_html=True)
    st.divider()
    page=st.radio('工作区',['研究总览','图谱探索','教学项目','支撑分析','来源与局限'],key='page',label_visibility='collapsed')
    st.divider()
    st.caption('研究数据快照')
    st.markdown('**34** 三级能力 · **8** 教学项目')
    st.caption('修订版支撑等级 1 / 18 / 14 / 1')
    st.caption('main · '+bundle.provenance['main_commit'][:10])
    st.caption('候选岗位与研究关系保留原状态；本原型不开展招聘核验。')


def title(kicker,heading,note):
    st.markdown('<div class="eyebrow">'+html.escape(kicker)+'</div>',unsafe_allow_html=True)
    st.title(heading);st.markdown('<div class="hero-note">'+html.escape(note)+'</div>',unsafe_allow_html=True)


def chart_counts(rows,label='匹配状态',number='能力数',order=None):
    frame=pd.DataFrame(rows)
    color=alt.Color(f'{label}:N',scale=alt.Scale(domain=list(STATUS_ORDER),range=[STATUS_COLORS[x] for x in STATUS_ORDER]),legend=None) if label=='匹配状态' else alt.value('#5a9180')
    bars=alt.Chart(frame).mark_bar(cornerRadiusEnd=5,size=25).encode(y=alt.Y(f'{label}:N',sort=order,axis=alt.Axis(title=None,labelFontSize=12)),x=alt.X(f'{number}:Q',title='能力数量',axis=alt.Axis(tickMinStep=1)),color=color,tooltip=list(frame.columns))
    text=bars.mark_text(align='left',dx=7,color='#49646b',fontSize=12).encode(text=f'{number}:Q')
    st.altair_chart((bars+text).properties(height=200).configure_view(stroke=None),use_container_width=True)


def set_graph(id='N113'):
    st.session_state.page='图谱探索';st.session_state.focus=id;st.session_state.graph_mode='单节点局部子图'
    st.session_state.query='';st.session_state.node_types=list(MAIN_TYPES);st.session_state.levels=['三级']
    st.session_state.grades=list(STATUS_ORDER);st.session_state.graph_project=''


def download_pair(nodes,edges,prefix,key=''):
    a,b=st.columns(2)
    a.download_button('下载'+prefix+'节点 CSV',csv_bytes(nodes),prefix+'节点.csv','text/csv',key='nodes'+key,disabled=not nodes,width='stretch')
    b.download_button('下载'+prefix+'关系 CSV',csv_bytes(edges),prefix+'关系.csv','text/csv',key='edges'+key,disabled=not edges,width='stretch')


def node_evidence(id,edges):
    d=detail(bundle,id,edges)
    with st.expander('中心节点的课程证据与完整属性',expanded=False):
        a,b=st.columns([1,1])
        with a:
            for k in ['编码','名称','类型','来源','证据类型','课程引用物理页码','审核状态','课程支撑等级']:
                st.caption(k);st.write(d[k])
        with b:
            if d['课程支撑的具体技术行为']:
                st.caption('课程支撑的具体技术行为');st.write(d['课程支撑的具体技术行为'])
                st.caption('课程尚未明确覆盖的工作行为');st.write(d['课程尚未明确覆盖的工作行为'])
            else:st.json(d['源属性'],expanded=False)
        eids=set(filter(None,str(bundle.by_id[id].get('课程证据ID') or '').split(';')))
        for m in bundle.matches:
            if m['能力ID']==id:eids.update(filter(None,str(m.get('课程证据ID') or '').split(';')))
        eids=sorted(eids)
        if eids:
            eid=st.selectbox('课程原文摘录',eids,format_func=lambda x:f"{x} · PDF第{bundle.evidence_by_id[x]['PDF页码']}页",key='evidence_select')
            e=bundle.evidence_by_id[eid]
            st.caption(e['课程证据状态']+' · 物理页码 '+str(e['PDF页码']))
            st.text_area('课程标准原文',str(e['原文摘录']),height=160,disabled=True)
        else:st.caption('本节点未单列课程引文；可查看相关节点或来源属性。')


if page=='研究总览':
    title('RESEARCH OVERVIEW / 研究总览','《数据库原理及应用》\n课程能力知识图谱','从旅游行业候选岗位与工作任务出发，查看职业能力与课程项目、知识点和实操要求的可追溯联系。')
    st.markdown('<span class="pill">PR #3 修订版</span><span class="pill">只读研究数据</span><span class="pill">默认隐藏高风险与隔离关系</span>',unsafe_allow_html=True)
    st.markdown('<div class="flow"><span>岗位候选</span><b>→</b><span>典型任务</span><b>→</b><span>三级能力</span><b>→</b><span>教学项目</span><b>→</b><span>知识点 / 实操</span></div>',unsafe_allow_html=True)
    counts=Counter(n['节点类型'] for n in bundle.nodes)
    cols=st.columns(5)
    for col,label,value in zip(cols,['岗位候选样本','典型工作任务','三级能力建议','教学项目','图谱原始节点'],[counts['岗位候选样本'],counts['典型工作任务（摘要）'],len(bundle.capabilities),len(bundle.project_by_id),len(bundle.nodes)]):col.metric(label,value)
    st.write('')
    left,right=st.columns([1.4,1],gap='large')
    with left:
        st.subheader('课程文本匹配快照');chart_counts([r for r in bundle.coverage if r['匹配状态'] in STATUS_ORDER],order=list(STATUS_ORDER))
        st.caption('分母34项三级候选能力。不是学生能力达成率，也不是经核验的行业需求覆盖率。')
    with right:
        st.subheader('从一个能力开始')
        st.markdown('<div class="caption-card"><h4>N113 · 查询并提取业务数据</h4><p>课程整体直接支撑的研究建议，与各项目的技术子环节分开。点击节点，回到原文、页码和判断理由。</p></div>',unsafe_allow_html=True)
        st.write('');st.button('打开 N113 局部子图 →',on_click=set_graph,width='stretch',type='primary')
        st.caption(f'源关系：{len(bundle.edges)} 条候选 · {len(bundle.isolated)} 条隔离。显示关系不等于已人工确认。')
        safe_n,safe_e=visible_data(bundle)
        st.caption(f'默认规则下可用 {len(safe_n)} 节点 / {len(safe_e)} 关系；图谱按当前中心和筛选进一步收束。')

elif page=='图谱探索':
    title('RELATION EXPLORER / 图谱探索','沿着关系，查看证据','先选一个能力、岗位或任务，从局部关系理解课程支撑。图内点击节点或连线可查来源与证据；未见支撑不生成正向连线。')
    with st.sidebar:
        st.divider();st.subheader('研究视图')
        explore=st.checkbox('研究探索模式',key='explore')
        if not explore:
            st.session_state.high_risk=False;st.session_state.isolated=False
        high=st.checkbox('显示高风险推断（未核验）',key='high_risk',disabled=not explore)
        iso=st.checkbox('显示隔离关系（未确认）',key='isolated',disabled=not explore)
        options=Options(explore,high,iso)
        if explore:st.caption('探索仅放开展示，不提升原证据或审核状态。')
        types=st.multiselect('节点类型',sorted({n['节点类型'] for n in bundle.nodes}),default=list(MAIN_TYPES),key='node_types')
        levels=st.multiselect('能力层级',['一级','二级','三级'],default=['三级'],key='levels')
        project=st.selectbox('教学项目',['']+sorted(bundle.project_by_id),format_func=lambda id:'全部教学项目' if not id else id+' · '+bundle.project_by_id[id]['节点名称'],key='graph_project')
        grades=st.multiselect('课程整体支撑等级',list(STATUS_ORDER),default=list(STATUS_ORDER),key='grades')
        mode=st.radio('视图范围',['单节点局部子图','能力局部子图','筛选范围图'],key='graph_mode')
        query=st.text_input('检索编码、名称或来源',placeholder='例如 N113、SQL、票务、备份',key='query')
        hops=st.slider('邻域深度',1,3,2,key='hops');limit=st.slider('绘制节点上限',20,120,70,10,key='limit')
        peers=st.checkbox('能力中心允许显示相邻能力',value=False,key='peers')
    base_nodes,_,_=filter_graph(bundle,options,types,levels,project or None,grades,max_nodes=10000)
    center=None
    if mode!='筛选范围图':
        choices=search_nodes(base_nodes,query)
        if mode=='能力局部子图':choices=[n for n in choices if n['节点类型']=='职业能力建议']
        ids=[n['节点ID'] for n in sorted(choices,key=lambda n:n['节点ID'])]
        if not ids:
            st.info('当前筛选没有可选中心节点，请调整检索、类型或层级。');st.stop()
        if st.session_state.get('focus') not in ids:st.session_state.focus='N113' if 'N113' in ids else ids[0]
        center=st.selectbox('中心节点（也可直接输入名称搜索）',ids,format_func=lambda id:id+' · '+bundle.by_id[id]['节点名称'],key='focus')
    nodes,edges,meta=filter_graph(bundle,options,types,levels,project or None,grades,center,hops,limit,peers,query if mode=='筛选范围图' else None)
    a,b,c=st.columns(3);a.metric('当前绘制节点',len(nodes));b.metric('当前候选关系',len(edges));c.metric('筛选匹配节点',meta['匹配节点数'])
    if explore and (high or iso):st.warning('研究探索视图包含未核验或隔离数据。它们不代表已确认招聘事实或完整课程培养。')
    if meta['已截断']:st.info(f"为保持清晰，仅绘制最近的{len(nodes)}个节点（匹配{meta['匹配节点数']}个）。当前CSV导出与画面一致；可提高上限或收紧筛选。")
    if nodes:components.html(build_html(bundle,nodes,edges,center),height=760,scrolling=False)
    else:st.info('当前筛选无可显示节点。未添加任何占位节点或关系。')
    download_pair(nodes,edges,'当前视图',key='filtered')
    if center and nodes:node_evidence(center,edges)
    with st.expander('当前视图的数据表'):
        st.dataframe(pd.DataFrame(nodes)[['节点ID','节点类型','节点名称','源文件或条目','证据或核验状态']] if nodes else pd.DataFrame(),hide_index=True,width='stretch')
        st.dataframe(pd.DataFrame(edges)[['关系ID','起点ID','终点ID','关系类型','审核或置信状态','展示分组']] if edges else pd.DataFrame(),hide_index=True,width='stretch')

elif page=='教学项目':
    title('TEACHING PROJECTS / 教学项目','以项目查看课程联系','项目层支撑等级与课程整体判断分开。下表依据第三阶段272个项目单元，不从关键词新造课程关系。')
    project=st.selectbox('选择教学项目',sorted(bundle.project_by_id),format_func=lambda id:id+' · '+bundle.project_by_id[id]['节点名称'],key='project_page')
    raw=bundle.attrs[project];st.subheader(bundle.by_id[project]['节点名称'])
    a,b,c=st.columns(3);a.metric('项目总学时',raw['总学时']);b.metric('理论学时',raw['理论学时']);c.metric('内含实践学时',raw['内含实践学时'])
    st.caption('内含实践已计入总学时，不再次叠加。')
    cells=[r for r in bundle.project_cells if r['课程项目ID']==project]
    matched=[r for r in cells if r['项目层匹配等级'] in {'直接支撑','部分支撑'}]
    st.markdown(f'**{len(matched)} 项能力存在项目文本支撑** · 分母34项候选能力，不表示完整培养或学生达成。')
    cols=['能力ID','能力名称','项目层匹配等级','知识点ID','实践任务ID','课程支撑的具体技术行为','课程尚未明确覆盖的工作行为']
    st.dataframe(pd.DataFrame(matched)[cols] if matched else pd.DataFrame(columns=cols),hide_index=True,width='stretch')
    with st.expander('全部34项项目单元（含未见支撑、待审核）'):st.dataframe(pd.DataFrame(cells),hide_index=True,width='stretch')
    st.button('在图谱中查看本项目 →',on_click=set_graph,args=(project,),type='primary')
    left,right=st.columns(2)
    for col,kind,title_text in [(left,'知识条目','项目知识条目'),(right,'实训任务','项目内实操要求')]:
        with col:
            st.subheader(title_text)
            items=[{'编码':n['节点ID'],'名称':n['节点名称'],'PDF页码':n['PDF页码'],'状态':n['证据或核验状态']} for n in bundle.nodes if n['节点类型']==kind and bundle.attrs[n['节点ID']].get('课程项目ID')==project]
            st.dataframe(pd.DataFrame(items),hide_index=True,width='stretch')
    st.subheader('潜在课程改进方向 · 研究建议')
    st.caption('以下4项为全课程层面的建议，尚未分配到本项目；不是课标已有内容，也不形成新的项目支撑边。')
    improvements=[r for r in bundle.improvements if r['未见支撑建议分类']=='本课程潜在改进方向']
    st.dataframe(pd.DataFrame(improvements)[['能力ID','能力名称','课程标准支撑不足或边界','分类理由']],hide_index=True,width='stretch')
    st.download_button('下载本项目34项复核 CSV',csv_bytes(cells),project+'项目复核.csv','text/csv')

elif page=='支撑分析':
    title('TEXT SUPPORT / 支撑分析','看支撑深度，而非达成结论','统计只描述课程标准文本与34项三级能力的研究匹配。不把文本支撑率称为学生能力达成率，也不作为经核验行业需求覆盖率。')
    counts=Counter(r['最终候选匹配状态'] for r in bundle.capabilities)
    cols=st.columns(4)
    for col,status in zip(cols,STATUS_ORDER):col.metric(status,f'{counts[status]} / 34',f'{counts[status]/34:.2%}',delta_color='off')
    left,right=st.columns([1.3,1])
    with left:chart_counts([r for r in bundle.coverage if r['匹配状态'] in STATUS_ORDER],order=list(STATUS_ORDER))
    with right:
        st.markdown('**分母与解释**')
        st.write('能力等级分母：34。项目单元分母：272。未见支撑建议分类分母：14。三种统计不能混用。')
        st.write('直接或部分文本对应为19/34，深度不同，不表示19项能力已完整培养。4项能力已由直接支撑调整为部分支撑。')
    selected=st.multiselect('筛选能力状态',list(STATUS_ORDER),default=list(STATUS_ORDER),key='analysis_statuses')
    cap_rows=[r for r in bundle.capabilities if r['最终候选匹配状态'] in selected]
    st.dataframe(pd.DataFrame(cap_rows)[['能力ID','能力名称','最终候选匹配状态','课程支撑的具体技术行为','课程尚未明确覆盖的工作行为','课程支撑判断']] if cap_rows else pd.DataFrame(),hide_index=True,width='stretch')
    st.download_button('下载当前能力分析 CSV',csv_bytes(cap_rows),'能力文本支撑分析.csv','text/csv',disabled=not cap_rows)
    st.subheader('未见明确支撑：14项建议分类')
    grouped=Counter(r['未见支撑建议分类'] for r in bundle.improvements)
    st.dataframe(pd.DataFrame([{'研究建议类别':k,'数量':v,'分母':14,'比例':f'{v/14:.2%}'} for k,v in grouped.items()]),hide_index=True,width='stretch')
    st.caption('分类为研究建议，不代表课程标准已证明；“其他课程可能承担”并未核实具体课程开设或覆盖。')

else:
    title('PROVENANCE / 来源与局限','每条联系，都保留来处','本地应用只读第三阶段修订版四份Excel，不生成新岗位、知识点、能力或课程关系，也不开展招聘核验。')
    st.markdown('**来源版本**：main `'+bundle.provenance['main_commit']+'` ；PR #3已合并，修订头 `'+bundle.provenance['pr3_head']+'`。')
    st.dataframe(pd.DataFrame([{'文件':name,'工作目录':str(path.relative_to(ROOT)),'SHA256':bundle.provenance['loaded_sha256'][path.relative_to(ROOT).as_posix()]} for name,path in bundle.files.items()]),hide_index=True,width='stretch')
    st.subheader('默认展示策略')
    st.write('默认隐藏隔离关系、含疑似标记的关系、任务到能力含研究推断的关系，以及旧能力归档和编码映射。课程文本支持的研究匹配可显示，但仍未人工确认。研究探索开关只改变展示，不升级证据状态。')
    st.write('岗位候选、摘要证据、研究推断、课程文本支撑及待审核状态保持原值。未见明确支撑不绘制虚构的支撑边。历史候选课程内容不作为已确认实训支撑。')
    st.subheader('研究局限')
    st.write('招聘样本及推断关系保留历史未核验状态，不代表独立真实行业需求分布。课程能力联系是研究判断，不是专家结论、实际教学实施或学习成效证据。本阶段不以招聘核验作为网页建设或后续阶段条件。')
    st.write('图中保留旧能力归档、上位能力、目标等辅助节点，因此原始314个节点不等于314项三级能力。图谱筛选、局部深度和绘制上限会改变显示数量；当前视图CSV始终与绘制结果一致。')
    st.subheader('CSV数据下载')
    explore=st.checkbox('研究探索下载（保留高风险及隔离状态）',key='source_explore')
    nodes,edges=visible_data(bundle,Options(explore,explore,explore))
    st.caption(f"{'研究探索全部数据' if explore else '默认候选视图'}：{len(nodes)}节点 / {len(edges)}关系。风险与原始审核字段随CSV保留。")
    download_pair(nodes,edges,'研究探索全部' if explore else '默认候选',key='sources')
