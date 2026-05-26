"""
projects 폴더의 xlsx 파일 변경/추가를 감시하다가
빌드 스크립트를 자동으로 실행합니다.

실행 방법:
    py watch_projects.py

종료:
    Ctrl+C
"""

import subprocess
import sys
import time
from pathlib import Path

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

ROOT = Path(__file__).resolve().parent
PROJECTS_DIR = ROOT / "projects"
BUILD_SCRIPT = ROOT / "build_project_dashboard.py"

DEBOUNCE_SECONDS = 2.0
POLL_SECONDS = 2.0


def snapshot_xlsx_mtime_ns() -> dict[str, int]:
    snapshot: dict[str, int] = {}
    for path in PROJECTS_DIR.rglob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        try:
            snapshot[str(path.resolve())] = path.stat().st_mtime_ns
        except FileNotFoundError:
            continue
    return snapshot


class XlsxHandler(FileSystemEventHandler):
    def __init__(self):
        self._last_run: float = 0

    def _should_handle_xlsx(self, path: str) -> bool:
        target = Path(path)
        return target.suffix.lower() == ".xlsx" and not target.name.startswith("~$")

    def _should_handle_build_script(self, path: str) -> bool:
        try:
            return Path(path).resolve() == BUILD_SCRIPT.resolve()
        except FileNotFoundError:
            return False

    def _rebuild(self, reason: str) -> None:
        now = time.monotonic()
        if now - self._last_run < DEBOUNCE_SECONDS:
            return
        self._last_run = now
        print(f"\n[감시] {reason}")
        print("[빌드] build_project_dashboard.py 실행 중...")
        result = subprocess.run(
            [sys.executable, str(BUILD_SCRIPT)],
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        if result.returncode == 0:
            print("[완료] 데이터 갱신 성공")
        else:
            print("[오류] 빌드 실패:")
            print(result.stderr or result.stdout)

    def on_created(self, event):
        if not event.is_directory and self._should_handle_xlsx(event.src_path):
            self._rebuild(f"새 파일 감지: {Path(event.src_path).name}")

    def on_modified(self, event):
        if event.is_directory:
            return
        if self._should_handle_xlsx(event.src_path):
            self._rebuild(f"파일 수정 감지: {Path(event.src_path).name}")
            return
        if self._should_handle_build_script(event.src_path):
            self._rebuild(f"빌드 스크립트 수정 감지: {Path(event.src_path).name}")

    def on_moved(self, event):
        if not event.is_directory and self._should_handle_xlsx(event.dest_path):
            self._rebuild(f"파일 이동/이름변경 감지: {Path(event.dest_path).name}")


def main() -> None:
    print(f"[감시 시작] {PROJECTS_DIR}")
    print("projects 폴더의 xlsx 파일이 변경/추가되면 자동으로 data를 갱신합니다.")
    print("종료하려면 Ctrl+C 를 누르세요.\n")

    # 시작 시 한 번 빌드
    print("[초기 빌드] 시작 시 데이터 최신화...")
    subprocess.run([sys.executable, str(BUILD_SCRIPT)], encoding="utf-8")

    handler = XlsxHandler()
    observer = Observer()
    observer.schedule(handler, str(PROJECTS_DIR), recursive=True)
    observer.schedule(handler, str(ROOT), recursive=False)
    observer.start()

    last_snapshot = snapshot_xlsx_mtime_ns()

    try:
        while True:
            time.sleep(POLL_SECONDS)
            current_snapshot = snapshot_xlsx_mtime_ns()
            if current_snapshot != last_snapshot:
                handler._rebuild("폴링 감지: xlsx 파일 변경")
                last_snapshot = current_snapshot
    except KeyboardInterrupt:
        observer.stop()
        print("\n[감시 종료]")

    observer.join()


if __name__ == "__main__":
    main()
