# 第三阶段最终规范化候选审查稿

只继续PR #7，不更改原数据、网页或教学映射。研究者认可3—9—17粒度，所有结果仍为候选。

在仓库根目录运行（Python 3.11/3.12）：

```bash
python -m pip install -r research_revision/database_competencies_v1/stage3_final_review/requirements.txt
python research_revision/database_competencies_v1/stage3_final_review/build.py
python -m unittest discover -s research_revision/database_competencies_v1/stage3_final_review -p 'test_final.py' -v
```

无需网络、招聘核验或PDF重新解析。输入按既有目录中的实际工作表/字段唯一定位。输入快照锁定PR7此前121文件；任何偏离将停止生成，不能通过改原数据绕过检查。--output可指定仓库外临时目录；仓库内只允许本轮目录，避免覆盖历史成果。

output/01、02、03、04、05、07为6个Excel；06为方法说明、08为总结，共8项研究成果。方法说明.md与修订总结.md为可编辑源稿，build.py复制到output对应文件，可重复生成。修改说明时也应同步重跑生成及测试。非层级关系表包括不保留的判断记录，不能将全部8行当作图谱边；6个保留/待审候选也默认不启用，已确认0。岗位摘要A类不等于原公告核验，历史109疑似关系不在本轮重新审核范围。

全部生成文件清单可见output；测试覆盖输入保护、稳定层级、证据不升级、逻辑关系类型及重跑内容一致性。验证结果.txt保存实际命令和结果。
