"""
FBL5N 채권 미결항목 다운로드 스크립트

사용법:
  python sapost/fbl5n_download.py --keydate 202503

  --keydate  조회 기준 년월 (YYYYMM). 해당 월 말일로 자동 변환됩니다.

동작 순서:
  1. source_dir 의 파일 목록에서 고객계정(파일명 앞 7자리) 수집
  2. 각 고객계정마다 FBL5N 실행
     - 미결항목 / 특별G/L거래 / 임시항목 선택
     - 기준일 = 해당 월 말일
  3. 전기일자 오름차순 정렬 후 엑셀 로컬 저장
  4. raw_dir 에 {계정코드}-{YYYYMM}.xlsx 로 저장
"""
import sys
import time
import shutil
import argparse
import calendar
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from sapost.src.utils import get_config, setup_logger

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32가 설치되어 있지 않습니다. pip install pywin32 를 실행하세요.")
    sys.exit(1)

from dotenv import load_dotenv
import os


# ──────────────────────────────────────────────────────────
# 헬퍼
# ──────────────────────────────────────────────────────────

def month_end(yyyymm: str) -> str:
    """'202603' → '2026.03.31' (SAP 날짜 형식 YYYY.MM.DD)"""
    year  = int(yyyymm[:4])
    month = int(yyyymm[4:6])
    last_day = calendar.monthrange(year, month)[1]
    return f"{year}.{month:02d}.{last_day:02d}"


def get_customer_accounts(source_dir: Path, logger) -> list[str]:
    """source_dir 의 파일명이 7자리 숫자로 시작하는 파일에서 고객계정 수집"""
    import re
    accounts = []
    seen = set()
    for f in sorted(source_dir.iterdir()):
        if not f.is_file():
            continue
        stem = f.stem
        # 파일명이 7자리 숫자로 시작하는 경우만 추출
        match = re.match(r'^(\d{7})', stem)
        if not match:
            logger.debug(f"고객계정 아님 — 건너뜀: {f.name}")
            continue
        account = match.group(1)
        if account not in seen:
            seen.add(account)
            accounts.append(account)
            logger.info(f"고객계정 추출: {account}  ← {f.name}")
    return accounts


# ──────────────────────────────────────────────────────────
# FBL5N 실행 클래스
# ──────────────────────────────────────────────────────────

class FBL5NDownloader:
    def __init__(self, config, logger):
        self.config = config
        self.logger = logger
        self.session = None

        env_path = Path(__file__).parent / "config" / ".env"
        load_dotenv(dotenv_path=env_path)

        # config.ini 에서 필드 ID 로드
        self.transaction      = config.get("SAP", "transaction", fallback="FBL5N")
        self.customer_field   = config.get("SAP", "customer_field_id")
        self.company_code     = config.get("SAP", "company_code", fallback="1000")
        self.company_code_field = config.get("SAP", "company_code_field")
        self.keydate_field    = config.get("SAP", "keydate_field_id")
        self.open_items_radio = config.get("SAP", "open_items_radio")
        self.special_gl_chk   = config.get("SAP", "special_gl_chk")
        self.noted_items_chk  = config.get("SAP", "noted_items_chk")
        self.posting_date_col = config.get("SAP", "posting_date_col", fallback="BUDAT")
        self.execute_vkey     = config.getint("SAP", "execute_vkey", fallback=8)

        self.raw_dir = Path(config.get("PATHS", "raw_dir"))
        self.raw_dir.mkdir(parents=True, exist_ok=True)

    def connect(self):
        """실행 중인 SAP GUI 세션에 연결"""
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application  = sap_gui_auto.GetScriptingEngine
        connection   = application.Children(0)
        self.session = connection.Children(0)
        self.logger.info("SAP GUI 세션 연결 완료")

    def run_all(self, accounts: list[str], keydate_str: str, yyyymm: str):
        """모든 고객계정에 대해 FBL5N 실행 → 다운로드"""
        success, failed = [], []

        for i, account in enumerate(accounts, 1):
            self.logger.info(f"[{i}/{len(accounts)}] 계정: {account}  기준일: {keydate_str}")
            try:
                dest = self._run_single(account, keydate_str, yyyymm)
                success.append(account)
                self.logger.info(f"  → 저장 완료: {dest.name}")
            except Exception as e:
                self.logger.error(f"  → 실패 ({account}): {e}")
                failed.append(account)
                self._go_back_to_start()

        self.logger.info("=" * 50)
        self.logger.info(f"완료: 성공 {len(success)}건 / 실패 {len(failed)}건")
        if failed:
            self.logger.warning(f"실패 계정: {failed}")

    def _run_single(self, account: str, keydate_str: str, yyyymm: str) -> Path:
        """단일 고객계정 FBL5N 실행 → 저장 → 파일 경로 반환"""
        # 1) FBL5N 트랜잭션 이동
        self._navigate_to_fbl5n()

        # 2) 선택 화면 입력
        self._fill_selection_screen(account, keydate_str)

        # 3) 실행 (F8)
        self.session.findById("wnd[0]").sendVKey(self.execute_vkey)
        time.sleep(3)
        self.logger.info("  FBL5N 조회 실행")

        # 4) ALV 그리드에서 직접 읽기 → pandas 정렬 → 엑셀 저장
        dest = self.raw_dir / f"{account}-{yyyymm}.xlsx"
        self._read_grid_and_save(dest)

        return dest

    def _navigate_to_fbl5n(self):
        """FBL5N 선택 화면으로 이동"""
        cmd = self.session.findById("wnd[0]/tbar[0]/okcd")
        cmd.text = f"/n{self.transaction}"
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

    def _fill_selection_screen(self, account: str, keydate_str: str):
        """FBL5N 선택 화면: 고객계정, 기준일, 체크박스 입력"""
        s = self.session

        # 고객계정
        s.findById(self.customer_field).text = account

        # 회사코드
        try:
            s.findById(self.company_code_field).text = self.company_code
        except Exception as e:
            self.logger.warning(f"  회사코드 입력 실패 (ID 확인 필요): {e}")

        # 미결항목 라디오 선택
        try:
            s.findById(self.open_items_radio).select()
        except Exception as e:
            self.logger.warning(f"  미결항목 라디오 선택 실패 (ID 확인 필요): {e}")

        # 기준일자
        s.findById(self.keydate_field).text = keydate_str

        # 특별G/L거래 체크
        try:
            s.findById(self.special_gl_chk).selected = True
        except Exception as e:
            self.logger.warning(f"  특별G/L 체크 실패 (ID 확인 필요): {e}")

        # 임시항목 체크
        try:
            s.findById(self.noted_items_chk).selected = True
        except Exception as e:
            self.logger.warning(f"  임시항목 체크 실패 (ID 확인 필요): {e}")

        self.logger.info(f"  선택 화면 입력 완료 (계정: {account}, 기준일: {keydate_str})")

    def _read_grid_and_save(self, dest: Path):
        """ALV 그리드 직접 읽기 → 전기일자 오름차순 정렬 → 엑셀 저장"""
        import pandas as pd

        grid_id = self.config.get("SAP", "grid_id")
        grid = self.session.findById(grid_id)

        row_count = grid.RowCount
        self.logger.info(f"  그리드 행 수: {row_count}")

        if row_count == 0:
            raise ValueError("조회 결과가 없습니다.")

        # 컬럼 목록
        columns = list(grid.ColumnOrder)
        self.logger.info(f"  컬럼 수: {len(columns)}")

        # 컬럼 헤더(화면 표시명) 수집
        headers = {}
        for col in columns:
            try:
                headers[col] = grid.GetDisplayedColumnTitle(col)
            except Exception:
                headers[col] = col

        # 전체 데이터 읽기
        records = []
        for row in range(row_count):
            record = {}
            for col in columns:
                try:
                    record[col] = grid.GetCellValue(row, col)
                except Exception:
                    record[col] = ""
            records.append(record)

        df = pd.DataFrame(records, columns=columns)

        # 컬럼명을 화면 표시명으로 변경
        df = df.rename(columns=headers)

        # 전기일자 오름차순 정렬 (BUDAT 또는 표시명)
        budat_col = headers.get(self.posting_date_col, self.posting_date_col)
        if budat_col in df.columns:
            df = df.sort_values(budat_col).reset_index(drop=True)
            self.logger.info(f"  전기일자 오름차순 정렬 완료")
        else:
            self.logger.warning(f"  전기일자 컬럼 '{budat_col}' 없음 — 정렬 생략")

        # 엑셀 저장
        dest.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(dest, index=False)
        self.logger.info(f"  엑셀 저장 완료: {dest}  ({len(df)}행)")

    def _go_back_to_start(self):
        """오류 발생 시 초기 화면으로 복귀"""
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
        except Exception:
            pass

    def close(self):
        self._go_back_to_start()


# ──────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(description="FBL5N 채권 미결항목 다운로드")
    parser.add_argument(
        "--keydate",
        required=True,
        help="조회 기준 년월 (예: 202503). 해당 월 말일로 자동 변환됩니다.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    yyyymm = args.keydate
    keydate_str = month_end(yyyymm)

    config = get_config()
    logger = setup_logger("sapost.fbl5n", config)

    logger.info("=" * 60)
    logger.info(f"FBL5N 다운로드 시작 | 기준일: {keydate_str}")
    logger.info("=" * 60)

    # 고객계정 목록 수집
    source_dir = Path(config.get("PATHS", "source_dir"))
    if not source_dir.exists():
        logger.error(f"source_dir 를 찾을 수 없습니다: {source_dir}")
        sys.exit(1)

    accounts = get_customer_accounts(source_dir, logger)
    if not accounts:
        logger.error("고객계정을 추출할 파일이 없습니다.")
        sys.exit(1)

    logger.info(f"총 {len(accounts)}개 계정 추출: {accounts}")

    # SAP 연결 및 실행
    downloader = FBL5NDownloader(config, logger)
    try:
        downloader.connect()
        downloader.run_all(accounts, keydate_str, yyyymm)
    finally:
        downloader.close()


if __name__ == "__main__":
    main()
