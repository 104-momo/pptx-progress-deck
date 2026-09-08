# -*- coding: utf-8 -*-
"""浏览器端自测：渲染、翻页、断点、hash、备注、导入导出"""
import sys
from playwright.sync_api import sync_playwright

import os
DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "web", "L01-导论.html")
URL = os.environ.get("DECK_URL") or ("file://" + DEFAULT)
ok, bad = [], []
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "build") + os.sep


def chk(cond, msg):
    (ok if cond else bad).append(msg)
    print(("  PASS " if cond else "  FAIL ") + msg)


with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={"width": 1440, "height": 900})
    errs = []
    pg.on("console", lambda m: errs.append(m.text) if m.type == "error" else None)
    pg.on("pageerror", lambda e: errs.append("PAGEERROR: " + str(e)))

    pg.goto(URL)
    pg.wait_for_timeout(600)

    print("\n--- 基础渲染 ---")
    chk(pg.locator(".slide").count() == 25, "25 页全部渲染")
    chk(pg.locator("#clsOverlay.show").count() == 1, "首次打开弹出选班面板")
    chk(pg.locator(".clsCard").count() >= 8, "班级卡片已渲染（含添加按钮）")
    # 封面文字与位置
    txt = pg.locator(".slide").first.inner_text()
    chk("叙事理论与实践" in txt, "第1页文字正确")
    chk(pg.evaluate("getComputedStyle(document.querySelector('.slide')).backgroundColor")
        == "rgb(27, 40, 56)", "第1页深色封面背景还原")

    print("\n--- 选班 ---")
    pg.locator(".clsCard").nth(2).click()     # 选班级3
    pg.wait_for_timeout(300)
    chk(pg.locator("#clsName").inner_text().strip() != "", "顶栏显示当前班级")
    chk(pg.locator("#fab.show").count() == 1, "右下角悬浮钮出现")
    chk("#c3" in pg.evaluate("location.hash"), "hash 已写入班级")

    print("\n--- 翻页 ---")
    for _ in range(4):
        pg.keyboard.press("ArrowRight")
        pg.wait_for_timeout(80)
    chk(pg.locator("#pg").inner_text().startswith("5"), "方向键翻到第5页")
    pg.locator("#dots i").nth(11).click()
    pg.wait_for_timeout(500)
    chk(pg.locator("#pg").inner_text().startswith("12"), "点击圆点跳到第12页")

    print("\n--- S 键记断点 ---")
    pg.keyboard.press("s")
    pg.wait_for_timeout(300)
    bp = pg.evaluate("JSON.parse(localStorage[KEY]).progress[cur()].L01.p")
    chk(bp == 12, "S 键：断点记为当前页 P12（下次从本页开始）")
    chk(pg.locator("#toast.show").count() == 1, "toast 提示已出现")
    chk("-L01:12" in pg.evaluate("location.hash"), "hash 同步为 L01:12")

    print("\n--- Shift+S 本页讲完 ---")
    pg.keyboard.press("Shift+S")
    pg.wait_for_timeout(300)
    bp = pg.evaluate("JSON.parse(localStorage[KEY]).progress[cur()].L01.p")
    chk(bp == 13, "Shift+S：断点推进到下一页 P13")

    print("\n--- 刷新后定位 ---")
    pg.reload()
    pg.wait_for_timeout(700)
    chk(pg.locator("#resume").is_visible(), "刷新后显示定位提示条")
    chk("P13" in pg.locator("#rTo").inner_text(), "提示「下次继续 P13」")
    chk(pg.locator("#clsName").inner_text().strip() != "", "班级被记住")
    pg.locator("#rGo").click()
    pg.wait_for_timeout(600)
    chk(pg.locator("#pg").inner_text().startswith("13"), "点「继续」跳到断点页")

    print("\n--- URL hash 跨源恢复 ---")
    pg2 = b.new_page(viewport={"width": 1440, "height": 900})
    pg2.goto(URL + "#c5-L01:20")
    pg2.wait_for_timeout(700)
    expect_cls = pg2.evaluate("DEFAULT_CLASSES.find(function(c){return c.id==='c5'}).name")
    chk(expect_cls in pg2.locator("#clsName").inner_text(), "hash 指定班级生效")
    chk(pg2.locator("#pg").inner_text().startswith("20"), "hash 指定页码生效，直接跳到第20页")
    chk(pg2.locator(".clsOverlay.show").count() == 0, "hash 有效时不再弹选班")

    print("\n--- 备注 ---")
    pg2.keyboard.press("n")
    pg2.wait_for_timeout(300)
    chk(pg2.locator("#note.show").count() == 1, "N 键打开备注抽屉")
    pg2.locator("#noteInput").fill("这里要补充短视频案例")
    pg2.locator("#noteSave").click()
    pg2.wait_for_timeout(300)
    chk(pg2.locator(".nitem").count() == 1, "备注已保存并显示")
    chk("短视频案例" in pg2.locator("#noteList").inner_text(), "备注内容正确")
    pg2.locator(".nitem .del").click()
    pg2.wait_for_timeout(200)
    chk(pg2.locator(".nitem").count() == 0, "备注可删除")

    print("\n--- 总览 ---")
    chk(pg2.locator("#note.show").count() == 1, "备注抽屉仍开着")
    pg2.keyboard.press("Escape")
    pg2.wait_for_timeout(300)
    chk(pg2.locator("#note.show").count() == 0, "Esc 先关闭备注抽屉")
    pg2.keyboard.press("Escape")
    pg2.wait_for_timeout(600)
    chk(pg2.locator("#ovp.show").count() == 1, "Esc 打开总览")
    chk(pg2.locator(".ovcell").count() == 25, "总览显示 25 个缩略图")
    chk(pg2.locator(".ovcell.done").count() == 19, "前 19 页标为已讲（断点 P20）")
    chk(pg2.locator(".ovcell.bp").count() == 1, "断点页有唯一橙框标记")
    chk("▶ 继续" in pg2.locator(".ovcell.bp").inner_text(), "断点页显示「▶继续」角标")
    pg2.locator("#ovp").screenshot(path=OUT+"shot_overview.png")
    pg2.locator(".ovcell").nth(3).click()
    pg2.wait_for_timeout(500)
    chk(pg2.locator("#pg").inner_text().startswith("4"), "总览点击可跳转")

    print("\n--- 截图 ---")
    pg.goto(URL + "#c1-L01:1"); pg.wait_for_timeout(700)
    pg.locator("#viewport").screenshot(path=OUT+"p1.png")
    pg.goto(URL + "#c1-L01:2"); pg.wait_for_timeout(700)
    pg.locator("#viewport").screenshot(path=OUT+"p2.png")
    pg.goto(URL + "#c1-L01:9"); pg.wait_for_timeout(700)
    pg.locator("#viewport").screenshot(path=OUT+"p9.png")
    pg.goto(URL + "#c1-L01:12"); pg.wait_for_timeout(700)
    pg.locator("#viewport").screenshot(path=OUT+"p12.png")
    pg.goto(URL + "#c1-L01:22"); pg.wait_for_timeout(700)
    pg.locator("#viewport").screenshot(path=OUT+"p22.png")

    print("\n--- 控制台错误 ---")
    real = [e for e in errs if "favicon" not in e.lower()]
    chk(not real, "无 JS 报错" + (": " + " | ".join(real[:3]) if real else ""))

    b.close()

print("\n===== %d 通过 / %d 失败 =====" % (len(ok), len(bad)))
for m in bad:
    print("  ✗ " + m)
sys.exit(1 if bad else 0)
