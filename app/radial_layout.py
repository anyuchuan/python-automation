"""真实森林的径向坐标；圆心不创建节点或边。"""
import math

def radial_layout(nodes,edges):
 by={n['node_id']:n for n in nodes};children={id:[] for id in by};incoming=set()
 for e in edges:
  if e['relation_code']!='R01':raise ValueError('径向能力总览只接受R01')
  children[e['source']].append(e['target']);incoming.add(e['target'])
 roots=[id for id in by if id not in incoming];ranges={};cursor=0
 def assign(id,trail):
  nonlocal cursor
  if id in trail:raise ValueError('层级循环')
  start=cursor
  if children[id]:
   for child in children[id]:assign(child,trail|{id})
  else:cursor+=1
  ranges[id]=(start,cursor)
 for root in roots:assign(root,set())
 angles={id:-math.pi/2+2*math.pi*((a+b)/2)/cursor for id,(a,b) in ranges.items()}
 levels={'COMPETENCY_L1':(220,150),'COMPETENCY_L2':(500,310),'COMPETENCY_L3':(850,520)}
 width,height=280,84
 def boxes(scale):
  return {id:{'x':levels[n['node_type']][0]*scale*math.cos(angles[id])-width/2,'y':levels[n['node_type']][1]*scale*math.sin(angles[id])-height/2,'width':width,'height':height,'angle':angles[id]} for id,n in by.items()}
 def overlap(a,b):return abs((a['x']+width/2)-(b['x']+width/2))<width+16 and abs(a['y']-b['y'])<height+12
 scale=1.0
 for _ in range(100):
  coords=boxes(scale);values=list(coords.values())
  if not any(overlap(a,b) for i,a in enumerate(values) for b in values[i+1:]):break
  scale*=1.04
 else:raise ValueError('径向标签间距无法满足')
 minx=min(p['x'] for p in values)-40;miny=min(p['y'] for p in values)-40;maxx=max(p['x']+width for p in values)+40;maxy=max(p['y']+height for p in values)+40
 for p in coords.values():p['x']-=minx;p['y']-=miny
 return {'positions':coords,'width':maxx-minx,'height':maxy-miny,'center':[-minx,-miny],'node_width':width,'node_height':height}
