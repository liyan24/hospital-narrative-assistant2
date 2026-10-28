# -*- coding: utf-8 -*-
"""补拍针灸需求对照文档所需截图（一次性工具）。"""
import json
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs" / "images" / "research"
BASE = "http://localhost:8501"

auth = json.load(open(ROOT / "auth_tmp.json", encoding="utf-8"))
INIT_JS = f"""
localStorage.setItem('token', {json.dumps(auth["access_token"])});
localStorage.setItem('user', {json.dumps(json.dumps(auth["user"], ensure_ascii=False))});
"""


def run_skill(page, skill_name, shot_name, wait_text="AI 解读"):
    """专家模式→分析Skills→选算子→运行→截图"""
    page.goto(f"{BASE}/portal/research-assistant", wait_until="networkidle")
    page.wait_for_timeout(2000)
    page.get_by_role("tab", name="专家模式").click()
    page.wait_for_selector("text=表格数据", timeout=120000)
    page.get_by_role("button", name="下一步").first.click()
    page.wait_for_timeout(2000)
    page.locator(".ant-menu-item", has_text=skill_name).first.click()
    page.wait_for_timeout(1000)
    page.get_by_role("button", name="运行").first.click()
    page.wait_for_selector(f"text={wait_text}", timeout=300000)
    page.wait_for_timeout(2000)
    page.screenshot(path=str(OUT / shot_name), full_page=True)
    print(f"[shot] {shot_name}")


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(viewport={"width": 1680, "height": 1000},
                                  device_scale_factor=1.5)
        ctx.add_init_script(INIT_JS)
        page = ctx.new_page()

        # 1. REQ-218 数据追踪：后台管理-缓存管理页（LLM 缓存条目含命名空间/时间/元数据）
        page.goto(f"{BASE}/admin/cache", wait_until="networkidle")
        page.wait_for_timeout(3000)
        page.screenshot(path=str(OUT / "req218_缓存追踪.png"), full_page=True)
        print("[shot] req218_缓存追踪.png")

        # 2. REQ-219 海量数据：专家模式数据资产页（含图谱统计）
        page.goto(f"{BASE}/portal/research-assistant", wait_until="networkidle")
        page.wait_for_timeout(2000)
        page.get_by_role("tab", name="专家模式").click()
        page.wait_for_selector("text=表格数据", timeout=120000)
        page.wait_for_timeout(2500)
        page.screenshot(path=str(OUT / "req219_数据资产.png"), full_page=True)
        print("[shot] req219_数据资产.png")

        # 3. REQ-221 疗效趋势预判：分类预测（再入院）
        run_skill(page, "分类预测", "req221_分类预测.png")

        # 4. REQ-222 疗效数据分析：趋势分析
        run_skill(page, "趋势分析", "req222_趋势分析.png")

        browser.close()
    print("补拍完成")


if __name__ == "__main__":
    main()
