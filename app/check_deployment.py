"""只读部署诊断：python app/check_deployment.py；不运行研究脚本或解析PDF。"""
import json
import platform
import sys
from collections import Counter
from importlib.metadata import version
from pathlib import Path
from time import perf_counter
ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:sys.path.insert(0,str(ROOT))
from app.data import load_bundle,visible_data

def check():
    started=perf_counter();data=load_bundle();nodes,edges=visible_data(data)
    return {'python':platform.python_version(),'packages':{p:version(p) for p in ['streamlit','pandas','openpyxl','altair']},
            'source_files':[str(p.relative_to(ROOT)) for p in data.files.values()],
            'nodes':len(data.nodes),'candidate_edges':len(data.edges),'isolated_edges':len(data.isolated),
            'default_nodes':len(nodes),'default_edges':len(edges),
            'support_counts':dict(Counter(r['最终候选匹配状态'] for r in data.capabilities)),
            'load_seconds':round(perf_counter()-started,3)}

if __name__=='__main__':
    try:print(json.dumps(check(),ensure_ascii=False,indent=2))
    except Exception as exc:
        print('部署只读检查失败：'+str(exc),file=sys.stderr);sys.exit(1)
