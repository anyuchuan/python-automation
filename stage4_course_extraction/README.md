# 第四阶段第一轮：课程标准结构化提取与教学要素建库

唯一课程原件：仓库根目录[课标0913.pdf](../课标0913.pdf)，完整13页；不重复存放，不修改原件。基线main及138既有文件哈希见input_snapshot.json。所有新增内容仅在本目录，不变更PR7能力、应用或历史教学映射。

建议Python 3.11/3.12，在仓库根目录：

```bash
python -m pip install -r stage4_course_extraction/requirements.txt
python stage4_course_extraction/build.py
python -m unittest discover -s stage4_course_extraction -p test_extraction.py -v
python stage4_course_extraction/render_review.py --output /tmp/course_pdf_review
```

build支持`--output /tmp/course_extraction_check`；仓库内仅允许本阶段目录，避免覆盖旧成果。不调用历史阶段生成器、不访问网络、不提取招聘信息。输入PDF有变更/缺失或摘录不匹配时停止，不能以旧表或教材替代。

可编辑extraction_manifest.json是本轮阅读原PDF后的人工定位清单，含固定CK/KO/CA/CO/EC编号、页码和摘录，不是旧课程数据复制。后续增补分配新ID，勿整体重排编号。build从PDF重新核对并保留原文；项目正文与目标由PDF章节结构直接切分。

output/包含要求的01—11项成果（9Excel＋2Markdown），另附完整机器可读数据库JSON、计数JSON和逐页文本。07证据索引及JSON含实体、关联、数据字典。09方法说明包括双向互补和学时同源原则。研究方法.md为可编辑源稿，重跑时复制到09。11总结由本次实际数据动态生成。

重要边界：第13页末尾实训和考核文本层渲染不可见，相关片段保留为待审隐藏文本。研究者本轮明确确认的实训名称及10/12/10学时、同源32规则另列来源，不能据此称全部技术内容已可视核验。INERT INTO保留原文。课程64=32理论+32实践；8项目实践32与3实训32同源，不相加。学习成果中的潜在SQL脚本、备份文件等标研究推断，未明示正式提交物不补写。检查反馈与任务完成要求不等于正式评分。

本轮没有17能力正式支持矩阵，也没有学生评价结果。知识与出现记录、活动与成果、检查与证据均分表，为下一轮审查后双向映射准备接口。阅读顺序建议：11总结→02项目→03知识→04活动成果→05实训→06评价→07证据→08学时→10问题。实际测试结果见验证结果.txt。
