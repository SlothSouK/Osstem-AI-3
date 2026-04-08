"""
영업수금(Collection) 집계 실행 진입점

사용법:
  python collection/main.py

  --file    자금일보 엑셀 파일 경로 (미지정 시 config.ini source_file 사용)
  --output  검증 엑셀 저장 경로 (미지정 시 source_file 동일 폴더에 자동 생성)
  --label   검증 파일 제목 (기본: '영업수금 2026년 3월')
"""
import sys
import argparse
import logging
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import configparser

from collection.src.collector import CollectionCollector


def get_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    cfg_path = Path(__file__).parent / "config" / "config.ini"
    cfg.read(cfg_path, encoding="utf-8")
    return cfg


def setup_logger(level_str: str) -> logging.Logger:
    level = getattr(logging, level_str.upper(), logging.INFO)
    logging.basicConfig(
        level=level,
        format="[%(asctime)s] %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    return logging.getLogger("collection")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="영업수금(Collection) 집계")
    parser.add_argument(
        "--file",
        help="자금일보 엑셀 파일 경로 (미지정 시 config.ini source_file 사용)",
    )
    parser.add_argument(
        "--output",
        help="검증 엑셀 저장 경로 (미지정 시 source_file 동일 폴더에 자동 생성)",
    )
    parser.add_argument(
        "--label",
        default="영업수금",
        help="검증 파일 제목 접두사 (기본: '영업수금')",
    )
    return parser.parse_args()


def main() -> None:
    args   = parse_args()
    config = get_config()
    log    = setup_logger(config.get("LOGGING", "level", fallback="INFO"))

    # 소스 파일 결정
    if args.file:
        source_file = Path(args.file)
    else:
        source_file = Path(config.get("PATHS", "source_file"))

    if not source_file.exists():
        log.error(f"파일을 찾을 수 없습니다: {source_file}")
        sys.exit(1)

    # 출력 경로 결정
    if args.output:
        out_path = Path(args.output)
    else:
        out_dir_cfg = config.get("PATHS", "output_dir", fallback="").strip()
        out_dir = Path(out_dir_cfg) if out_dir_cfg else source_file.parent
        stem = source_file.stem.replace(" ", "_")[:30]
        out_path = out_dir / f"영업수금_Collection_{stem}_검증.xlsx"

    # 집계 실행
    collector = CollectionCollector(config, log)
    result    = collector.collect(source_file)

    # 시트별 소계 출력
    log.info("─" * 50)
    for sheet, subtotal in result.by_sheet().items():
        log.info(f"  {sheet:<35} {subtotal:>15,.2f}")
    log.info("─" * 50)
    log.info(f"  {'영업수금 합계':<35} {result.total:>15,.2f}  ({len(result.rows)}건)")
    log.info("─" * 50)

    # 검증 파일 생성
    collector.export_verification(result, out_path, label=args.label)
    log.info(f"완료: {out_path}")


if __name__ == "__main__":
    main()
