"""
충당금 계산 자동화 파이프라인 진입점

사용법:
  python automation/main.py --month 202503

  --month  기준월 (YYYYMM 형식, 생략 시 현재 월 사용)
  --skip-erp  ERP 조작을 건너뛰고 raw/ 에 이미 있는 파일로 처리
"""
import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from automation.src.utils import get_config, setup_logger
from automation.src.erp_controller import ERPController
from automation.src.downloader import Downloader
from automation.src.data_processor import DataProcessor
from automation.src.template_writer import TemplateWriter


def parse_args():
    parser = argparse.ArgumentParser(description="충당금 계산 자동화")
    parser.add_argument(
        "--month",
        default=datetime.now().strftime("%Y%m"),
        help="기준월 (예: 202503). 기본값: 현재 월",
    )
    parser.add_argument(
        "--skip-erp",
        action="store_true",
        help="ERP 조작 생략 (raw/ 폴더에 이미 다운로드된 파일이 있을 때 사용)",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    month = args.month

    # ── 초기화 ──────────────────────────────────────
    config = get_config()
    logger = setup_logger("main", config)

    logger.info("=" * 60)
    logger.info(f"충당금 자동화 시작 | 기준월: {month}")
    logger.info("=" * 60)

    start_time = time.time()
    result_summary = {"month": month, "erp": None, "download": None, "process": None, "write": None}

    # ── STEP 1-2: ERP 조작 & 다운로드 ───────────────
    raw_dir = Path(config.get("PATHS", "raw_dir"))
    raw_dir.mkdir(parents=True, exist_ok=True)

    if args.skip_erp:
        logger.info("[SKIP] ERP 조작 생략 — raw/ 폴더의 기존 파일 사용")
        file_paths = sorted(raw_dir.glob(f"{month}_*.xlsx"))
        if not file_paths:
            logger.error(f"raw/ 폴더에 {month}_*.xlsx 파일이 없습니다.")
            sys.exit(1)
        logger.info(f"기존 파일 {len(file_paths)}건 발견")
    else:
        erp = None
        try:
            logger.info("[STEP 1] ERP 연결 및 로그인")
            erp = ERPController(config, logger)
            erp.connect()
            erp.login()
            result_summary["erp"] = "success"

            logger.info("[STEP 2] 기준월 조회 및 다운로드")
            erp.set_month_and_search(month)
            item_count = erp.get_item_count()
            logger.info(f"조회된 항목 수: {item_count}건")

            downloader = Downloader(config, logger)
            file_paths = downloader.download_all(erp, item_count, month)
            result_summary["download"] = f"{len(file_paths)}건"

            if not file_paths:
                logger.error("다운로드된 파일이 없습니다. 종료합니다.")
                sys.exit(1)

        except Exception as e:
            logger.critical(f"ERP/다운로드 단계 오류: {e}")
            result_summary["erp"] = f"failed: {e}"
            sys.exit(1)
        finally:
            if erp:
                erp.close()

    # ── STEP 3: 데이터 정제 ──────────────────────────
    try:
        logger.info("[STEP 3] 데이터 정제")
        processor = DataProcessor(config, logger)
        df = processor.process(file_paths, month)
        result_summary["process"] = f"{len(df)}행"
        logger.info(f"정제 완료: {len(df)}행")
    except Exception as e:
        logger.critical(f"데이터 정제 오류: {e}")
        result_summary["process"] = f"failed: {e}"
        sys.exit(1)

    # ── STEP 4: 양식 붙여넣기 ────────────────────────
    try:
        logger.info("[STEP 4] 양식 붙여넣기")
        writer = TemplateWriter(config, logger)
        output_path = writer.write(df, month)
        result_summary["write"] = str(output_path)
    except Exception as e:
        logger.critical(f"양식 기입 오류: {e}")
        result_summary["write"] = f"failed: {e}"
        sys.exit(1)

    # ── 완료 요약 ────────────────────────────────────
    elapsed = time.time() - start_time
    logger.info("=" * 60)
    logger.info("자동화 완료 요약")
    logger.info(f"  기준월    : {result_summary['month']}")
    logger.info(f"  다운로드  : {result_summary['download']}")
    logger.info(f"  정제 결과 : {result_summary['process']}")
    logger.info(f"  산출 파일 : {result_summary['write']}")
    logger.info(f"  소요 시간 : {elapsed:.1f}초")
    logger.info("=" * 60)
    logger.info("※ 최종 결과 파일을 열어 이상값 여부를 검수하세요.")


if __name__ == "__main__":
    main()
