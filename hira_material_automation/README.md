# HIRA Material Automation

This project now uses three separate execution tracks.

1. `HIRA Backfill 2020-2022`
2. `HIRA Backfill 2023-2025`
3. `HIRA Material Rolling Sync`

## Why the flows are split

- Backfill is safer in smaller year blocks.
- Rolling stays focused on 2026 and later.
- Master files are rebuilt from monthly files every run, so overlap does not accumulate.

## Included categories

- `250099` 관절경 수술시 사용하는 활액 임시대체재
- `040021` ALL SUTURE ANCHOR
- `040019` 비금속성 ANCHOR
- `041001` 관절경 CANNULA
- `040023` 십자인대고정용 일차 고정재 - SCREW,BUTTON 등(금속류)
- `040012` 십자인대고정용-INTERFERENCE SCREW(흡수성)
- `040028` 1회용 관절 봉합용 NEEDLE(분리형)
- `040017` 반월상연골봉합술용-DOUBLE ARM
- `040029` 1회용 관절 봉합용 NEEDLE(일체형)
- `040026` 반월상연골봉합술용-SCREW,ANCHOR
- `900086` 연조직 재건용
- `900129` 척추경막외 유착방지제

## Range rules

- Backfill batch 1: `2020-01` to `2022-12`
- Backfill batch 2: `2023-01` to `2025-12`
- Rolling strategy: `current_year_replace`
- Rolling refresh range: `current year January` through `current month`
- Rolling floor month: `2026-01`
- Before every rolling run, files that overlap the target year are deleted and downloaded again
- Rolling schedule: the 5th, 15th, and 25th at `06:00` Asia/Seoul

## Output structure

- `raw`: original HIRA downloads, one file per category per queried range
- `output/monthly`: normalized CSV/XLSX files produced from each range download
- `output/master`: transformed master outputs
- `output/master/hira_material_summary.xlsx`: one workbook with a `통합` sheet and one sheet per category
- `logs/last_download_run.json`: detailed run history
- `logs/latest_master_report.json`: duplicate-removal and master rebuild summary

## Data shaping rules

- `-` values are converted to numeric `0`
- `연도` is added as a separate column
- `건강보험` and `의료급여` are kept separately and also rolled up into `청구량 합계` and `청구금액 합계`
- Master files are deduplicated by `기간 + 중분류코드`
- Rolling refresh replaces the entire current-year dataset so HIRA revisions are picked up cleanly

## Local commands

Rolling refresh:

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\run_hira_material_sync.ps1" -Mode rolling
```

Manual range run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\run_hira_material_sync.ps1" -Mode range -StartMonth 2020-01 -EndMonth 2022-12
```

Task Scheduler registration:

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\register_hira_material_task.ps1"
```

Pull the latest GitHub results back to the local PC:

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\sync_hira_results_from_github.ps1"
```

## GitHub Actions

- `HIRA Backfill 2020-2022`: one-time manual backfill for 2020-2022
- `HIRA Backfill 2023-2025`: one-time manual backfill for 2023-2025
- `HIRA Material Rolling Sync`: ongoing rolling sync for 2026 and later
- All workflows upload logs and partial outputs even if the run fails
- Run the local GitHub sync script after a workflow if you want the same files copied into this PC
