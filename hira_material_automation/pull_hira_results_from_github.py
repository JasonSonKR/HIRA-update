from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


SYNC_DIRECTORIES = [
    "raw",
    "output/monthly",
    "output/master",
    "logs",
]


def github_request_json(url: str) -> object:
    request = Request(url, headers={"User-Agent": "Codex-HIRA-Sync"})
    with urlopen(request) as response:
        return json.load(response)


def github_download(download_url: str, target_path: Path) -> None:
    request = Request(download_url, headers={"User-Agent": "Codex-HIRA-Sync"})
    with urlopen(request) as response:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(response.read())


def list_repo_files(owner: str, repo: str, branch: str, remote_path: str) -> list[dict]:
    api_url = f"https://api.github.com/repos/{owner}/{repo}/contents/{remote_path}?ref={branch}"
    payload = github_request_json(api_url)
    if isinstance(payload, dict):
        payload = [payload]

    files: list[dict] = []
    for item in payload:
        item_type = item.get("type")
        if item_type == "dir":
            files.extend(list_repo_files(owner, repo, branch, item["path"]))
        elif item_type == "file":
            files.append(item)
    return files


def iter_local_files(base_directory: Path) -> Iterable[Path]:
    for path in base_directory.rglob("*"):
        if path.is_file():
            yield path


def sync_directory(owner: str, repo: str, branch: str, app_root: Path, relative_directory: str, clean: bool) -> dict:
    remote_prefix = f"hira_material_automation/{relative_directory}"
    local_directory = app_root / Path(relative_directory)
    local_directory.mkdir(parents=True, exist_ok=True)

    remote_files = list_repo_files(owner, repo, branch, remote_prefix)
    expected_local_files: set[Path] = set()
    downloaded = 0

    for remote_file in remote_files:
        remote_path = remote_file["path"]
        if remote_path.endswith("/.gitkeep"):
            continue
        local_relative = Path(remote_path).relative_to("hira_material_automation")
        target_path = app_root / local_relative
        expected_local_files.add(target_path.resolve())
        github_download(remote_file["download_url"], target_path)
        downloaded += 1

    deleted = 0
    if clean:
        for local_file in iter_local_files(local_directory):
            if local_file.name == ".gitkeep":
                continue
            if local_file.resolve() not in expected_local_files:
                local_file.unlink()
                deleted += 1

        # Remove empty directories left after cleanup.
        for directory in sorted(local_directory.rglob("*"), reverse=True):
            if directory.is_dir() and not any(directory.iterdir()):
                directory.rmdir()

    return {
        "directory": relative_directory,
        "downloaded": downloaded,
        "deleted": deleted,
    }


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Pull HIRA workflow results from GitHub into the local workspace.")
    parser.add_argument("--owner", default="JasonSonKR", help="GitHub owner")
    parser.add_argument("--repo", default="HIRA-update", help="GitHub repository name")
    parser.add_argument("--branch", default="main", help="Git branch or ref")
    parser.add_argument("--no-clean", action="store_true", help="Do not remove local files missing from GitHub")
    return parser


def main() -> int:
    args = build_argument_parser().parse_args()
    app_root = Path(__file__).resolve().parent
    reports = []

    try:
        for relative_directory in SYNC_DIRECTORIES:
            reports.append(
                sync_directory(
                    owner=args.owner,
                    repo=args.repo,
                    branch=args.branch,
                    app_root=app_root,
                    relative_directory=relative_directory,
                    clean=not args.no_clean,
                )
            )
    except (HTTPError, URLError) as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        return 1

    print(
        json.dumps(
            {
                "ok": True,
                "owner": args.owner,
                "repo": args.repo,
                "branch": args.branch,
                "reports": reports,
                "summary_workbook": str(app_root / "output" / "master" / "hira_material_summary.xlsx"),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
