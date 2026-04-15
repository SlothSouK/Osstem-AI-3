"""
공통 유틸리티: 설정 로드, 로깅 설정
"""
import logging
import configparser
from pathlib import Path
from datetime import datetime


def get_config(config_path: str = None) -> configparser.ConfigParser:
    """config.ini 파일을 읽어 반환"""
    if config_path is None:
        base = Path(__file__).parent.parent
        config_path = base / "config" / "config.ini"

    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")
    return config


def setup_logger(name: str, config: configparser.ConfigParser) -> logging.Logger:
    """날짜별 로그 파일 + 콘솔 동시 출력 로거 생성"""
    log_level = config.get("LOGGING", "level", fallback="INFO")
    log_dir = Path(config.get("LOGGING", "log_dir", fallback="vercheck/logs"))
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"{datetime.now().strftime('%Y%m%d')}.log"

    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, log_level))

    if not logger.handlers:
        fmt = logging.Formatter(
            "[%(asctime)s] %(levelname)-8s %(name)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)

        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        logger.addHandler(ch)

    return logger
