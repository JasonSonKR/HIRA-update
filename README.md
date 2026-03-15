# HIRA-update

This repository runs the HIRA treatment-material middle-category sync in two ways.

1. GitHub Actions: runs even when the local PC is off.
2. Windows Task Scheduler: fallback when local execution is preferred.

Main folder:

- `hira_material_automation`

Main workflow:

- `.github/workflows/hira-material-sync.yml`
