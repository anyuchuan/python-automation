# 第四阶段第二轮候选图谱

本目录独立新增，原能力体系、课程提取、PDF及网页只读。原版本见input_snapshot.json，所有既有文件逐一哈希验证。技术判断规则见support_rules.json，误连案例见boundary_cases.json。新增/修改规则需检查技术对象、行为、结果及课程原文，不用关键词自动补边。

仓库根目录运行：

```bash
python -m pip install -r stage4_competency_graph/requirements.txt
python stage4_competency_graph/build.py
python -m unittest discover -s stage4_competency_graph -p test_graph.py -v
```

默认输出本目录output，支持--output /tmp/graph_review。原第三阶段目录由实际结构定位，课程以第一轮机器JSON为主并读取Excel库存/字段；不运行历史生成器。复用第一轮build.py的Excel写入与哈希辅助函数，不调用其提取流程。

输出01—12为9个Excel、3个Markdown；另有graph_nodes/edges/evidence/review_items/paths.json及summary.json。总关系表保留审查/排除记录，消费时必须按default_visible筛选；默认展示也只是候选关系，未经专家确认。多跳路径不是原文明示的直接能力培养或评分。未关联课程知识完整保留。实际测试日志见验证结果.txt。
