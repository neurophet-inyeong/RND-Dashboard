from __future__ import annotations

import os
import time
from pathlib import Path

import pandas as pd

SOURCE_PATH = Path(
    os.environ.get(
        "SHAREPOINT_LEDGER_PATH",
        r"C:\Users\박인영(InyeongPark)\OneDrive - 뉴로핏 주식회사\R&D\00. 연구개발과제 관리대장\뉴로핏_연구과제_통합관리.xlsx",
    )
)
OUT_DIR = Path(__file__).resolve().parent / "ledger"
OUT_PATH = OUT_DIR / "총괄표_원본데이터.xlsx"
SHEETS = ["총괄표", "원본데이터"]
POLL_SECONDS = 5.0


def export_ledger() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
        for sheet in SHEETS:
            df = pd.read_excel(SOURCE_PATH, sheet_name=sheet, header=0)
            df.to_excel(writer, sheet_name=sheet, index=False)


def main() -> None:
    print(f"[LEDGER-SYNC] 감시 시작: {SOURCE_PATH}")
    print(f"[LEDGER-SYNC] polling: {POLL_SECONDS}초, 출력: {OUT_PATH}")

    last_mtime: float | None = None
    while True:
        if not SOURCE_PATH.exists():
            print(f"[LEDGER-SYNC] 원본 파일을 찾을 수 없습니다: {SOURCE_PATH}")
        else:
            mtime = SOURCE_PATH.stat().st_mtime
            if mtime != last_mtime:
                try:
                    export_ledger()
                    print(f"[LEDGER-SYNC] 갱신 완료 -> {OUT_PATH}")
                except Exception as exc:
                    print(f"[LEDGER-SYNC] 갱신 실패: {exc}")
                last_mtime = mtime

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
