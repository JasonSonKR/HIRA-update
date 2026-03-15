from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from playwright.sync_api import sync_playwright

from process_hira_mhtml_xls import ensure_directory, load_app_config, normalize_hira_export, resolve_config_paths


@dataclass
class Category:
    code: str
    name: str
    slug: str


def shift_month(today: date, months_back: int) -> tuple[str, str]:
    year = today.year
    month = today.month - months_back
    while month <= 0:
        month += 12
        year -= 1
    return f"{year:04d}-{month:02d}", f"{year:04d}{month:02d}"


def target_month(config: dict) -> tuple[str, str]:
    override = config["download"].get("target_month_override", "").strip()
    if override:
        return override, override.replace("-", "")
    return shift_month(date.today(), int(config["download"].get("months_lag", 2)))


def load_categories(config: dict) -> list[Category]:
    categories = []
    for item in config.get("categories", []):
        categories.append(
            Category(
                code=str(item["code"]).strip(),
                name=str(item["name"]).strip(),
                slug=str(item.get("slug") or item["code"]).strip(),
            )
        )
    return categories


def browser_launch(playwright, browser_name: str, headless: bool):
    if browser_name == "chrome":
        return playwright.chromium.launch(channel="chrome", headless=headless)
    return playwright.chromium.launch(headless=headless)


def query_and_download(page, category: Category, ym_dash: str, ym_compact: str, raw_directory: Path) -> dict:
    page.goto("https://opendata.hira.or.kr/op/opc/olapMaterialTab3.do", wait_until="networkidle", timeout=120000)
    page.evaluate(
        """({code, name}) => {
            document.querySelector('#mcatMdivCd').value = code;
            document.querySelector('#mcatMdivCdNm').value = name;
            document.querySelector('#searchWrd').value = name;
            const searchDate = document.querySelector('#searchDate');
            if (searchDate) {
              searchDate.style.display = 'block';
            }
        }""",
        {"code": category.code, "name": category.name},
    )
    page.fill("#sYm", ym_dash)
    page.fill("#eYm", ym_dash)
    page.locator("#searchBtn").click()
    page.wait_for_load_state("networkidle", timeout=120000)
    page.wait_for_timeout(1000)

    rows = []
    for row in page.locator("div.tblType02.data table tr").all()[:10]:
        cells = [cell.inner_text().strip() for cell in row.locator("th, td").all()]
        if cells:
            rows.append(cells)

    file_name = f"{ym_compact}__{category.slug}__{category.code}.xlsx"
    target_path = raw_directory / file_name
    with page.expect_download(timeout=20000) as download_info:
        page.locator("#exlBtn").click()
    download = download_info.value
    download.save_as(str(target_path))

    return {
        "category_code": category.code,
        "category_name": category.name,
        "target_month": ym_dash,
        "raw_file": str(target_path),
        "table_preview": rows,
    }


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Download HIRA treatment-material claim stats and normalize them.")
    parser.add_argument("--config", type=Path, default=None, help="Path to config.json")
    parser.add_argument("--browser", choices=["chromium", "chrome"], default="chromium", help="Browser runtime")
    parser.add_argument("--headless", action="store_true", help="Run browser headless")
    parser.add_argument("--headed", action="store_true", help="Run browser with a visible window")
    parser.add_argument("--force", action="store_true", help="Download even if the raw file already exists")
    parser.add_argument("--skip-process", action="store_true", help="Only download raw xlsx files")
    return parser


def main() -> int:
    parser = build_argument_parser()
    args = parser.parse_args()

    config, config_root = load_app_config(args.config)
    config = resolve_config_paths(config, config_root)

    paths = config["paths"]
    processing = config["processing"]
    raw_directory = ensure_directory(paths["raw_directory"])
    output_directory = ensure_directory(paths["output_directory"])
    log_directory = ensure_directory(paths["log_directory"])
    run_log_path = log_directory / "last_download_run.json"

    ym_dash, ym_compact = target_month(config)
    categories = load_categories(config)
    if not categories:
        print("No categories were defined in config.json")
        return 1

    if args.headed:
        headless = False
    elif args.headless:
        headless = True
    else:
        headless = bool(config["download"].get("headless", True))

    results = []
    with sync_playwright() as playwright:
        browser = browser_launch(playwright, args.browser, headless)
        context = browser.new_context(locale="ko-KR", accept_downloads=True)
        page = context.new_page()

        for category in categories:
            target_raw = raw_directory / f"{ym_compact}__{category.slug}__{category.code}.xlsx"
            if target_raw.exists() and not args.force:
                item = {
                    "category_code": category.code,
                    "category_name": category.name,
                    "target_month": ym_dash,
                    "raw_file": str(target_raw),
                    "skipped_download": True,
                }
            else:
                item = query_and_download(page, category, ym_dash, ym_compact, raw_directory)
                item["skipped_download"] = False

            if not args.skip_process:
                item["normalized"] = normalize_hira_export(
                    source_path=Path(item["raw_file"]),
                    output_directory=output_directory,
                    required_keywords=processing["required_header_keywords"],
                    minimum_keyword_matches=int(processing["minimum_keyword_matches"]),
                    write_csv_file=bool(processing["write_csv"]),
                    write_xlsx_file=bool(processing["write_xlsx"]),
                )
            results.append(item)

        context.close()
        browser.close()

    run_log_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
