from __future__ import annotations

import json
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple

ROOT = Path(__file__).resolve().parent
TOKEN_DIR = ROOT / ".watch"
TOKEN_FILE = TOKEN_DIR / "reload-token.json"
POLL_SECONDS = 2.0

EXCLUDED_DIRS = {".git", "__pycache__", ".watch", ".vscode"}
EXCLUDED_FILES = {"reload-token.json"}


def is_excluded(path: Path) -> bool:
    rel = path.relative_to(ROOT)
    if any(part in EXCLUDED_DIRS for part in rel.parts):
        return True
    if path.name in EXCLUDED_FILES:
        return True
    return False


def snapshot_files() -> Dict[str, Tuple[float, int]]:
    state: Dict[str, Tuple[float, int]] = {}
    for p in ROOT.rglob("*"):
        if not p.is_file():
            continue
        if is_excluded(p):
            continue
        rel = str(p.relative_to(ROOT)).replace("\\", "/")
        stat = p.stat()
        state[rel] = (stat.st_mtime, stat.st_size)
    return state


def write_reload_token(counter: int) -> None:
    TOKEN_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "token": f"{counter}-{int(time.time())}",
        "updatedAt": datetime.now().isoformat(timespec="seconds"),
    }
    TOKEN_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def run_git(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", *args],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
    )


def git_ready() -> tuple[bool, str]:
    inside = run_git("rev-parse", "--is-inside-work-tree")
    if inside.returncode != 0:
        return False, "git 저장소가 아닙니다. (git init 필요)"

    remote = run_git("remote")
    if remote.returncode != 0 or not remote.stdout.strip():
        return False, "원격 저장소가 없습니다. (git remote add origin <url> 필요)"

    branch = run_git("branch", "--show-current")
    if branch.returncode != 0 or not branch.stdout.strip():
        return False, "현재 브랜치를 확인할 수 없습니다."

    return True, branch.stdout.strip()


def auto_commit_and_push() -> None:
    ready, detail = git_ready()
    if not ready:
        print(f"[AUTO-SYNC] Git skip: {detail}")
        return

    branch = detail

    add_result = run_git("add", "-A")
    if add_result.returncode != 0:
        print("[AUTO-SYNC] git add 실패")
        print(add_result.stderr.strip())
        return

    diff_result = run_git("diff", "--cached", "--quiet")
    if diff_result.returncode == 0:
        print("[AUTO-SYNC] 변경사항 없음 (commit skip)")
        return

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    commit_result = run_git("commit", "-m", f"auto: sync changes {ts}")
    if commit_result.returncode != 0:
        print("[AUTO-SYNC] git commit 실패")
        print(commit_result.stderr.strip())
        return

    print("[AUTO-SYNC] commit 완료")
    push_result = run_git("push", "origin", branch)
    if push_result.returncode != 0:
        print("[AUTO-SYNC] git push 실패")
        print(push_result.stderr.strip())
        return

    print(f"[AUTO-SYNC] push 완료 (origin/{branch})")


def main() -> None:
    print("[AUTO-SYNC] 감시 시작: C:/Copilot/RND-Dashboard")
    print(f"[AUTO-SYNC] polling: {POLL_SECONDS}초")

    previous = snapshot_files()
    counter = 0
    write_reload_token(counter)

    while True:
        time.sleep(POLL_SECONDS)
        current = snapshot_files()

        if current == previous:
            continue

        added = sorted(set(current) - set(previous))
        removed = sorted(set(previous) - set(current))
        maybe_changed = sorted(set(current).intersection(previous))
        changed = [p for p in maybe_changed if current[p] != previous[p]]

        total_changes = len(added) + len(removed) + len(changed)
        print(f"[AUTO-SYNC] 변경 감지: {total_changes}건 (추가 {len(added)}, 수정 {len(changed)}, 삭제 {len(removed)})")

        previous = current
        counter += 1
        write_reload_token(counter)
        auto_commit_and_push()


if __name__ == "__main__":
    main()
