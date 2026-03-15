# HIRA Material Automation

This project supports two execution modes.

1. Off-PC mode via GitHub Actions.
2. Local fallback via Windows Task Scheduler.

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

## Sync rules

- Initial one-time backfill target: `2020-01` to `2025-12`
- Rolling refresh window: `current - 2 months` through `current month`
- Rolling schedule target: the 5th, 15th, and 25th at `06:00` Asia/Seoul
- Existing monthly files are overwritten in place when HIRA values change
- Master files are rebuilt from monthly files every run, so values are not accumulated twice

## Output structure

- `raw`: original HIRA downloads, one file per category per month
- `output/monthly`: normalized monthly CSV/XLSX files from the raw HIRA files
- `output/master`: transformed master outputs
- `output/master/hira_material_summary.xlsx`: one workbook with a `통합` sheet and one sheet per category
- `logs/last_download_run.json`: detailed run history
- `logs/latest_master_report.json`: duplicate-removal and master rebuild summary

## Data shaping rules

- `-` values are converted to numeric `0`
- `연도` is added as a separate column
- `건강보험` and `의료급여` rows are kept separately and also rolled up into `청구량 합계` and `청구금액 합계`

## Common commands

Rolling refresh:

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\run_hira_material_sync.ps1" -Mode rolling
```

One-time backfill:

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\run_hira_material_sync.ps1" -Mode backfill
```

Custom month range:

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\run_hira_material_sync.ps1" -Mode range -StartMonth 2024-01 -EndMonth 2024-12
```

Task Scheduler registration:

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\register_hira_material_task.ps1"
```

## GitHub Actions

- Scheduled run: every month on the 5th, 15th, and 25th at `06:00` Asia/Seoul
- Manual run: choose `rolling`, `backfill`, or `range`
- Workflow file: `.github/workflows/hira-material-sync.yml`
