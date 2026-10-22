"""无CDN的双向Streamlit组件；不增加可视化依赖。"""
from pathlib import Path
import streamlit.components.v1 as components
from app.graph import LANES,TYPE_COLORS,STATUS_COLORS
from app.data import edge_style

_component=components.declare_component('course_platform',path=str(Path(__file__).parent/'frontend'))

def visual(**kwargs):return _component(**kwargs,default=None)

def graph_payload(bundle,nodes,edges):
    return {'nodes':[{'id':n['节点ID'],'name':n['节点名称'],'type':n['节点类型'],'lane':LANES.get(n['节点类型'],3),
                      'capability':n['节点类型']=='职业能力建议',
                      'color':STATUS_COLORS.get(bundle.cap_by_id.get(n['节点ID'],{}).get('最终候选匹配状态'),TYPE_COLORS.get(n['节点类型'],'#7a8c9e'))} for n in nodes],
            'edges':[{'id':e['关系ID'],'from':e['起点ID'],'to':e['终点ID'],'type':e['关系类型'],
                      'color':edge_style(e)[1],'dash':edge_style(e)[2]} for e in edges]}
