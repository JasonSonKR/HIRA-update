# HIRA Material Automation

This folder supports two execution modes.

1. Off-PC candidate:
   - Run the Playwright downloader in GitHub Actions.
   - This works only after the repo is pushed to GitHub and Actions are enabled.
2. Local fallback:
   - Run the same downloader on a powered-on Windows PC with Task Scheduler.

## What is automated

- Open the HIRA `중분류별 청구현황` page.
- Query configured middle-category codes.
- Download the official HIRA `.xlsx`.
- Normalize the main table into `.csv` and `.xlsx`.

## Current default category

- `040019`
- `비금속성 ANCHOR`

## Key files

- `config.json`: category list, date lag, and paths
- `download_hira_material.py`: browser download plus normalization
- `process_hira_mhtml_xls.py`: normalize downloaded HIRA exports
- `lookup_hira_category.py`: search middle-category codes by keyword
- `run_hira_material_sync.ps1`: one-shot local sync
- `run_hira_material_pipeline.ps1`: process-only runner for files already downloaded
- `register_hira_material_task.ps1`: local Task Scheduler registration helper

## Local folders

- `inbox`: manual files to normalize
- `raw`: downloaded original HIRA files
- `output`: normalized outputs
- `archive`: archived manual inbox files
- `logs`: last run logs and manifests

## Local run

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\run_hira_material_sync.ps1"
```

## Local category lookup

```powershell
python \
  ".\hira_material_automation\lookup_hira_category.py" \
  Anchor --browser chrome --headless
```

## Local fallback scheduling

```powershell
powershell -ExecutionPolicy Bypass -File ".\hira_material_automation\register_hira_material_task.ps1"
```

Default schedule:

- Every Monday
- 10:15 AM Asia/Seoul
- Runs only when the PC is on and the user can log in

## Off-PC path

The repo now includes a GitHub Actions workflow file under:

- `.github/workflows/hira-material-sync.yml`

That path becomes active only after:

1. This repo is pushed to GitHub.
2. GitHub Actions is enabled.
3. The workflow is allowed to push content updates back to the repo.

## Official-source notes

- HIRA Open API guide: `https://opendata.hira.or.kr/op/opc/selectOpenApiInfoView.do`
- Treatment-material info Open API: `https://www.data.go.kr/data/3074384/openapi.do`
- Monthly treatment-material claim stats page: `https://opendata.hira.or.kr/op/opc/olapMaterialTab3.do`
