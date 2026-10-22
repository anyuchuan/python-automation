"""可选真实浏览器检查。先启动8501端口应用，再运行本文件。"""
import csv
import io
import os
from pathlib import Path
from playwright.sync_api import sync_playwright

out=Path('app/artifacts');out.mkdir(exist_ok=True)
Path('work').mkdir(exist_ok=True)
with sync_playwright() as p:
 browser=p.chromium.launch(executable_path=os.environ.get('CHROMIUM_PATH','/usr/bin/chromium'),args=['--no-sandbox'])
 page=browser.new_page(viewport={'width':1440,'height':1050},accept_downloads=True)
 errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
 page.goto(os.environ.get('APP_URL','http://127.0.0.1:8501'))
 page.get_by_text('课程文本匹配快照',exact=True).wait_for()
 page.screenshot(path=str(out/'overview.png'),full_page=True)
 page.get_by_text('图谱探索',exact=True).click()
 frame=page.frame_locator('iframe[data-testid="stIFrame"]')
 frame.locator('.node[data-id="N113"]').wait_for()
 assert frame.locator('.node').count()==20
 assert frame.locator('.edge').count()==39
 frame.locator('.node[data-id="N113"]').click()
 assert '课程尚未明确覆盖的工作行为' in frame.locator('#detail').inner_text()
 before=frame.locator('#graph').get_attribute('viewBox')
 frame.locator('#zoom-in').click();assert frame.locator('#graph').get_attribute('viewBox')!=before
 frame.locator('#fit').click();assert frame.locator('#graph').get_attribute('viewBox')==before
 frame.locator('.node').first.focus();frame.locator('.node').first.press('Enter')
 assert frame.locator('#detail h2').inner_text()
 # 路径位于节点下方，DOM点击验证其事件及证据面板。
 frame.locator('.edge').first.dispatch_event('click')
 assert '关系证据' in frame.locator('#detail').inner_text()
 page.screenshot(path=str(out/'graph.png'),full_page=True)
 with page.expect_download() as event:page.get_by_text('下载当前视图节点 CSV',exact=True).click()
 downloaded=event.value;downloaded.save_as('work/phase4_download.csv')
 with open('work/phase4_download.csv',encoding='utf-8-sig',newline='') as f:assert len(list(csv.DictReader(f)))==20
 page.get_by_text('研究探索模式',exact=True).click()
 page.get_by_text('显示高风险推断（未核验）',exact=True).click()
 page.get_by_text('研究探索视图包含未核验或隔离数据。',exact=False).wait_for()
 page.get_by_text('研究探索模式',exact=True).click()
 page.get_by_text('支撑分析',exact=True).click()
 page.get_by_text('未见明确支撑：14项建议分类',exact=True).wait_for()
 page.screenshot(path=str(out/'analysis.png'),full_page=True)
 assert not page.locator('[data-testid="stException"]').count()
 page.set_viewport_size({'width':390,'height':844})
 page.get_by_text('未见明确支撑：14项建议分类',exact=True).wait_for()
 page.screenshot(path='work/phase4_mobile.png',full_page=True)
 assert not page.locator('[data-testid="stException"]').count()
 assert not errors,errors
 browser.close()
print('PASS: 浏览器节点点击、键盘、缩放、连线证据、CSV下载、风险开关、统计页、移动视口；无JavaScript错误')
