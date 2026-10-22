"""离线原生SVG交互图：无需CDN、JavaScript包或图数据库。"""
import json
from pathlib import Path
from app.data import detail,edge_style

TYPE_COLORS={'岗位候选样本':'#377fa5','典型工作任务（摘要）':'#6292a2','职业能力建议':'#12806b','教学项目':'#c58a31','知识条目':'#7587b4','实训任务':'#bd788d'}
STATUS_COLORS={'直接支撑':'#12806b','部分支撑':'#c68a28','未见明确支撑':'#93a0a9','待人工审核':'#8b68b0'}
LANES={'岗位候选样本':0,'典型工作任务（摘要）':1,'职业能力建议':2,'原职业能力（归档）':2,'课程建议能力（无招聘证据）':2,'教学项目':3,'知识条目':4,'实训任务':5,'课程目标':3,'课程能力要求':3,'考核相关条目':5,'课程标准文件':3}


def build_html(bundle,nodes,edges,center=None):
    payload={'nodes':[],'edges':[],'center':center}
    for n in nodes:
        info=detail(bundle,n['节点ID'],edges);status=bundle.cap_by_id.get(n['节点ID'],{}).get('最终候选匹配状态')
        payload['nodes'].append({'id':n['节点ID'],'name':n['节点名称'],'type':n['节点类型'],'lane':LANES.get(n['节点类型'],3),
            'color':STATUS_COLORS.get(status,TYPE_COLORS.get(n['节点类型'],'#8b68b0')),'info':info})
    for e in edges:
        label,color,dash=edge_style(e)
        payload['edges'].append({'id':e['关系ID'],'from':e['起点ID'],'to':e['终点ID'],'name':e['关系类型'],'color':color,'dash':dash,
            'state':e['审核或置信状态'],'risk':e.get('展示风险说明',''),'quote':e.get('课程标准原文摘录') or e.get('岗位摘要证据') or '',
            'source':e['来源文件'],'sourceId':e['源记录ID']})
    # 研究原文按JSON转义后仅经textContent进入DOM，不插入可执行HTML。
    data=json.dumps(payload,ensure_ascii=False).replace('<','\\u003c').replace('>','\\u003e').replace('&','\\u0026').replace('\u2028','\\u2028').replace('\u2029','\\u2029')
    template=(Path(__file__).parent/'graph.html').read_text()
    return template.replace('/*__GRAPH_DATA__*/',data)
