"""
Excel 파일 로더: 시트별 헤더 자동 탐지, 키 컬럼 해결, SheetData 반환
"""
from __future__ import annotations

import configparser
import logging
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class SheetData:
    name: str
    header_row_idx: int          # 0-based 헤더 행 인덱스
    df: pd.DataFrame             # 헤더 적용된 DataFrame (dtype=str, 정규화됨)
    key_columns: list[str]       # 행 키로 사용하는 컬럼명 목록
    skipped: bool = False
    skip_reason: str = ""


class ExcelReader:
    def __init__(self, config: configparser.ConfigParser, logger: logging.Logger) -> None:
        self.config = config
        self.logger = logger
        self.header_scan_rows = config.getint("EXCEL", "header_scan_rows", fallback=15)
        self.header_min_fill = config.getfloat("EXCEL", "header_min_fill_ratio", fallback=0.3)
        raw_keys = config.get("EXCEL", "key_columns", fallback="").strip()
        self.configured_keys: list[str] = (
            [k.strip() for k in raw_keys.split(",") if k.strip()] if raw_keys else []
        )
        skip_raw = config.get("EXCEL", "skip_sheets", fallback="").strip()
        self.skip_sheets: set[str] = (
            {s.strip() for s in skip_raw.split(",") if s.strip()} if skip_raw else set()
        )

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def load(self, path: Path) -> dict[str, SheetData]:
        """Excel 파일의 모든 시트를 읽어 {시트명: SheetData} 반환"""
        self.logger.info(f"파일 로드: {path.name}")
        try:
            xl = pd.ExcelFile(path, engine="openpyxl")
        except Exception as exc:
            self.logger.error(f"파일 열기 실패: {path} — {exc}")
            raise

        result: dict[str, SheetData] = {}
        for sheet_name in xl.sheet_names:
            if sheet_name in self.skip_sheets:
                self.logger.debug(f"  시트 건너뜀 (skip_sheets): {sheet_name}")
                continue
            sd = self._load_sheet(xl, sheet_name)
            result[sheet_name] = sd
            status = f"SKIP({sd.skip_reason})" if sd.skipped else f"헤더행={sd.header_row_idx + 1}, 행={len(sd.df)}, 열={len(sd.df.columns)}"
            self.logger.info(f"  [{sheet_name}] {status}")

        return result

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _load_sheet(self, xl: pd.ExcelFile, sheet_name: str) -> SheetData:
        # 1단계: 헤더 없이 원시 읽기 (헤더 탐지용)
        try:
            raw = xl.parse(sheet_name, header=None, dtype=str,
                           nrows=self.header_scan_rows + 50)
        except Exception as exc:
            self.logger.warning(f"  [{sheet_name}] 읽기 실패: {exc}")
            return SheetData(name=sheet_name, header_row_idx=0,
                             df=pd.DataFrame(), key_columns=[],
                             skipped=True, skip_reason=f"읽기 오류: {exc}")

        if raw.empty:
            return SheetData(name=sheet_name, header_row_idx=0,
                             df=pd.DataFrame(), key_columns=[],
                             skipped=True, skip_reason="빈 시트")

        # 2단계: 헤더 행 탐지
        try:
            hdr_idx = self._detect_header_row(raw)
        except ValueError:
            self.logger.warning(f"  [{sheet_name}] 헤더 행 자동 탐지 실패 — 첫 번째 행을 헤더로 사용")
            hdr_idx = 0

        # 3단계: 헤더 적용 후 재읽기
        df = xl.parse(sheet_name, header=hdr_idx, dtype=str)

        # 4단계: 정규화
        df = self._normalize(df)

        if df.empty or len(df.columns) == 0:
            return SheetData(name=sheet_name, header_row_idx=hdr_idx,
                             df=pd.DataFrame(), key_columns=[],
                             skipped=True, skip_reason="데이터 없음")

        # 5단계: 키 컬럼 결정
        key_cols = self._resolve_key_columns(df, sheet_name)

        return SheetData(name=sheet_name, header_row_idx=hdr_idx,
                         df=df, key_columns=key_cols)

    def _detect_header_row(self, raw: pd.DataFrame) -> int:
        """상위 header_scan_rows 행 중 비공백 셀 비율이 가장 높은 행 인덱스 반환"""
        scan = raw.iloc[: self.header_scan_rows]
        best_idx = -1
        best_score = -1.0

        for i, row in scan.iterrows():
            total = len(row)
            if total == 0:
                continue
            non_empty = row.apply(
                lambda v: bool(str(v).strip()) and str(v).strip().lower() != "nan"
            ).sum()
            ratio = non_empty / total
            if ratio > best_score:
                best_score = ratio
                best_idx = int(i)  # type: ignore[arg-type]

        if best_idx < 0 or best_score < self.header_min_fill:
            raise ValueError("헤더 행 탐지 실패")

        return best_idx

    def _normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        """컬럼명·셀값 strip, 'nan' 문자열 → '' 변환"""
        # 컬럼명 정규화 (중복 컬럼은 .1, .2 suffix 유지)
        df.columns = [str(c).strip() for c in df.columns]

        # 셀값 정규화
        df = df.astype(str)
        df = df.applymap(lambda v: "" if v.strip().lower() == "nan" else v.strip())

        # 완전 빈 행 제거
        df = df[~(df == "").all(axis=1)].reset_index(drop=True)

        return df

    def _resolve_key_columns(self, df: pd.DataFrame, sheet_name: str) -> list[str]:
        """config 우선 → 고유도 ≥0.9 컬럼 자동 탐지 → 첫 번째 컬럼 폴백"""
        cols_lower = {c.lower(): c for c in df.columns}

        # 1순위: config에 지정된 키 컬럼
        if self.configured_keys:
            matched = []
            for k in self.configured_keys:
                canonical = cols_lower.get(k.lower())
                if canonical:
                    matched.append(canonical)
                else:
                    self.logger.debug(f"  [{sheet_name}] 키 컬럼 '{k}' 미발견")
            if matched:
                return matched

        # 2순위: 고유도 ≥ 0.9 컬럼 자동 탐지
        n = len(df)
        if n > 0:
            for col in df.columns:
                uniqueness = df[col].nunique() / n
                if uniqueness >= 0.9:
                    self.logger.debug(f"  [{sheet_name}] 키 컬럼 자동 탐지: '{col}' (고유도={uniqueness:.2f})")
                    return [col]

        # 3순위: 첫 번째 컬럼
        first = df.columns[0]
        self.logger.warning(
            f"  [{sheet_name}] 키 컬럼 특정 불가 — 첫 번째 컬럼 사용: '{first}' (중복 키 발생 가능)"
        )
        return [first]
