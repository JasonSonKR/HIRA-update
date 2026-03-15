from __future__ import annotations

import argparse
import json

from playwright.sync_api import sync_playwright


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Search HIRA middle-category codes by keyword.")
    parser.add_argument("query", help="Search term, for example Anchor")
    parser.add_argument("--browser", choices=["chromium", "chrome"], default="chromium")
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--headed", action="store_true")
    return parser


def launch_browser(playwright, browser_name: str, headless: bool):
    if browser_name == "chrome":
        return playwright.chromium.launch(channel="chrome", headless=headless)
    return playwright.chromium.launch(headless=headless)


def main() -> int:
    parser = build_argument_parser()
    args = parser.parse_args()
    headless = not args.headed

    with sync_playwright() as playwright:
        browser = launch_browser(playwright, args.browser, headless)
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        page.goto("https://opendata.hira.or.kr/op/opc/olapMaterialTab3.do", wait_until="networkidle", timeout=120000)
        page.locator("#searchWrd").click()
        page.wait_for_timeout(500)
        page.fill("#searchWrd1", args.query)
        page.locator("#popSearchBtn").click()
        page.wait_for_timeout(1500)

        rows = []
        for row in page.locator("#layerSearchTbody tr").all():
            cells = [cell.inner_text().strip() for cell in row.locator("td").all()]
            if len(cells) >= 2:
                rows.append({"code": cells[0], "name": cells[1]})

        print(json.dumps(rows, ensure_ascii=False, indent=2))
        context.close()
        browser.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
