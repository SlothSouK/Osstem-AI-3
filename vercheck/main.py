"""
VerCheck — Excel 버전 비교 도구

실행:
  python vercheck/main.py
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from vercheck.src.utils import get_config, setup_logger
from vercheck.src.excel_reader import ExcelReader
from vercheck.src.comparator import WorkbookComparator
from vercheck.src.report_writer import ReportWriter


# ---------------------------------------------------------------------------
# 인터랙티브 입력
# ---------------------------------------------------------------------------

def _clean(text: str) -> str:
    return text.strip().strip('"').strip("'").strip()


def _resolve_path(dir_input: str, name_input: str) -> Path:
    """
    경로와 파일명을 조합해 Path 반환.
    - 경로 입력에 파일명(.xlsx)이 포함된 경우 그대로 사용
    - 파일명에 확장자 없으면 .xlsx 자동 추가
    """
    d = _clean(dir_input)
    n = _clean(name_input)

    # 경로 입력이 이미 .xlsx로 끝나는 경우 (파일명까지 함께 입력한 경우)
    if d.lower().endswith(".xlsx"):
        return Path(d)

    # 파일명에 확장자 없으면 .xlsx 추가
    if n and not n.lower().endswith(".xlsx"):
        n = n + ".xlsx"

    return Path(d) / n if n else Path(d)


def _prompt_inputs() -> tuple[Path, Path]:
    """4개 조건 입력 → y/r/n 확인 루프. 검증 통과한 두 파일 경로 반환."""

    while True:
        print()
        print("=" * 55)
        print("  VerCheck — Excel 버전 비교 도구")
        print("=" * 55)
        print("  ※ 파일 경로에 파일명까지 입력해도 됩니다.")

        # ── 파일 1 ────────────────────────────────────────────
        print("\n[비교 파일 1 (구버전)]")
        file1_dir  = input("  파일 경로  : ")
        file1_name = input("  파일명     : ")

        # ── 파일 2 ────────────────────────────────────────────
        print("\n[비교 파일 2 (신버전)]")
        file2_dir  = input("  파일 경로  : ")
        file2_name = input("  파일명     : ")

        path1 = _resolve_path(file1_dir, file1_name)
        path2 = _resolve_path(file2_dir, file2_name)

        # ── 입력 요약 ─────────────────────────────────────────
        print()
        print("-" * 55)
        print("  입력 확인")
        print("-" * 55)
        print(f"  구버전 파일 : {path1}")
        print(f"  신버전 파일 : {path2}")
        print(f"  결과 저장   : {path2.parent}")
        print("-" * 55)

        # ── 오류 사전 점검 ────────────────────────────────────
        errors = []
        if not path1.exists():
            errors.append(f"  [오류] 파일 1을 찾을 수 없습니다: {path1}")
        elif path1.suffix.lower() != ".xlsx":
            errors.append(f"  [오류] 파일 1이 .xlsx가 아닙니다: {path1.name}")
        if not path2.exists():
            errors.append(f"  [오류] 파일 2를 찾을 수 없습니다: {path2}")
        elif path2.suffix.lower() != ".xlsx":
            errors.append(f"  [오류] 파일 2가 .xlsx가 아닙니다: {path2.name}")

        if errors:
            for e in errors:
                print(e)
            print()

        # ── y / r / n 확인 ────────────────────────────────────
        while True:
            ans = input("  실행하시겠습니까?  [y=실행 / r=재입력 / n=취소] : ").strip().lower()
            if ans in ("y", "r", "n"):
                break
            print("  y, r, n 중 하나를 입력해주세요.")

        if ans == "y":
            if errors:
                print("  파일 경로 오류를 먼저 수정하고 재입력해주세요.")
                continue
            return path1, path2

        elif ans == "r":
            continue

        else:  # n
            print("\n  취소되었습니다.")
            sys.exit(0)


# ---------------------------------------------------------------------------
# 메인
# ---------------------------------------------------------------------------

def main() -> None:
    config = get_config()
    logger = setup_logger("vercheck", config)

    # 인터랙티브 입력
    old_path, new_path = _prompt_inputs()

    # 결과 파일: 파일 2 경로에 저장
    out_dir = new_path.parent
    stem = f"{old_path.stem}_vs_{new_path.stem}"
    output_path = out_dir / f"비교_{stem}.xlsx"

    print()
    logger.info("=" * 55)
    logger.info("VerCheck 비교 시작")
    logger.info(f"  구버전: {old_path}")
    logger.info(f"  신버전: {new_path}")
    logger.info(f"  출력:   {output_path}")
    logger.info("=" * 55)

    start = time.time()

    reader     = ExcelReader(config, logger)
    comparator = WorkbookComparator(config, logger)
    writer     = ReportWriter(config, logger)

    logger.info("[STEP 1] 구버전 파일 로드")
    old_sheets = reader.load(old_path)

    logger.info("[STEP 2] 신버전 파일 로드")
    new_sheets = reader.load(new_path)

    logger.info("[STEP 3] 비교 수행")
    diff = comparator.compare(old_sheets, new_sheets)

    logger.info("[STEP 4] 리포트 생성")
    writer.write(diff, old_sheets, new_sheets, output_path)

    elapsed = time.time() - start
    total_changes = sum(len(sd.cell_changes) for sd in diff.sheet_diffs)
    total_added   = sum(len(sd.added_rows)   for sd in diff.sheet_diffs)
    total_deleted = sum(len(sd.deleted_rows) for sd in diff.sheet_diffs)

    logger.info("=" * 55)
    logger.info("비교 완료")
    logger.info(f"  변경 셀   : {total_changes}개")
    logger.info(f"  추가 행   : {total_added}개")
    logger.info(f"  삭제 행   : {total_deleted}개")
    logger.info(f"  리포트    : {output_path}")
    logger.info(f"  소요 시간 : {elapsed:.1f}초")
    logger.info("=" * 55)


if __name__ == "__main__":
    main()
