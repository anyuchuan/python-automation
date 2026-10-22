"""共享选择与双向组件事件状态；不改研究数据。"""
from app.platform_model import ancestors

def choose(state,tree,id):
    if id not in tree:return False
    state['selected_cap']=id;state['net_center']=id;state['net_selected']=id
    state['tree_expanded']=sorted(set(state.get('tree_expanded',[]))|set(ancestors(tree,id)))
    state['locate']=True
    return True

def apply_event(state,event,tree,bundle):
    if not isinstance(event,dict):return False
    action=event.get('action');id=event.get('id')
    if action=='expand':
        state['tree_expanded']=[i for i in event.get('expanded',[]) if i in tree and tree[i]['level']!='三级'];state['locate']=False;return True
    if action in {'select','tree_locate','matrix_select'}:
        if not choose(state,tree,id):return False
        if action=='select':state['locate']=False
        if action=='tree_locate':state['graph_view']='树状层级'
        if action=='matrix_select':
            if event.get('project') not in bundle.project_by_id:return False
            state['selected_project']=event['project'];state['matrix_project']=event['project']
        return True
    if action in {'network_select','local_expand'} and id in bundle.by_id:
        state['net_selected']=id;state['selected_edge']=None
        if id in tree:
            state['selected_cap']=id;state['tree_expanded']=ancestors(tree,id);state['locate']=True
        if action=='local_expand':state['net_center']=id
        return True
    if action=='edge' and id in {e['关系ID'] for e in bundle.edges+bundle.isolated}:
        state['selected_edge']=id;return True
    return False
