"""
비교 로직: SheetData 두 개를 받아 diff 결과(SheetDiff, WorkbookDiff)를 반환
셀 위치 기반이 아닌 헤더+행 키 기반 비교
"""
from __future__ import annotations

import configparser
import logging
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

from .excel_reader import SheetData


# ---------------------------------------------------------------------------
# 결과 데이터클래스
# ---------------------------------------------------------------------------

@dataclass
class CellChange:
    sheet: str
    row_key: str        # 복합 키를 ' | '로 연결한 문자열
    col_name: str
    old_val: str
    new_val: str
    delta: Optional[float]      # 숫자 변경이면 new - old, 문자 변경이면 None
    delta_pct: Optional[float]  # delta / abs(old) * 100, 0 나누기이면 None


@dataclass
class RowChange:
    sheet: str
    row_key: str
    kind: str               # "added" | "deleted"
    values: dict[str, str]  # 해당 행의 모든 컬럼값 스냅샷


@dataclass
class SheetDiff:
    sheet_name: str
    added_cols: list[str] = field(default_factory=list)
    removed_cols: list[str] = field(default_factory=list)
    common_cols: list[str] = field(default_factory=list)
    cell_changes: list[CellChange] = field(default_factory=list)
    added_rows: list[RowChange] = field(default_factory=list)
    deleted_rows: list[RowChange] = field(default_factory=list)
    old_only_sheet: bool = False    # 구버전에만 있는 시트
    new_only_sheet: bool = False    # 신버전에만 있는 시트
    skipped: bool = False
    skip_reason: str = ""


@dataclass
class WorkbookDiff:
    sheet_diffs: list[SheetDiff] = field(default_factory=list)

    @property
    def all_cell_changes(self) -> list[CellChange]:
        result = []
        for sd in self.sheet_diffs:
            result.extend(sd.cell_changes)
        return result

    @property
    def summary_rows(self) -> list[dict]:
        rows = []
        for sd in self.sheet_diffs:
            rows.append({
                "sheet": sd.sheet_name,
                "cell_changes": len(sd.cell_changes),
                "added_rows": len(sd.added_rows),
                "deleted_rows": len(sd.deleted_rows),
                "added_cols": len(sd.added_cols),
                "removed_cols": len(sd.removed_cols),
                "note": (
                    "(삭제됨)" if sd.old_only_sheet
                    else "(추가됨)" if sd.new_only_sheet
                    else ("빈 시트" if sd.skipped else "")
                ),
            })
        return rows


# ---------------------------------------------------------------------------
# SheetComparator
# ---------------------------------------------------------------------------

class SheetComparator:
    def __init__(self, config: configparser.ConfigParser, logger: logging.Logger) -> None:
        self.tolerance = config.getfloat("EXCEL", "numeric_tolerance", fallback=0.0001)
        self.logger = logger

    def compare(self, old: SheetData, new: SheetData) -> SheetDiff:
        diff = SheetDiff(sheet_name=old.name)

        if old.skipped and new.skipped:
            diff.skipped = True
            diff.skip_reason = "양쪽 모두 빈/오류 시트"
            return diff

        # 빈 시트 한쪽만
        if old.skipped or new.skipped:
            diff.skipped = True
            diff.skip_reason = f"{'구버전' if old.skipped else '신버전'} 시트 비교 불가 ({(old if old.skipped else new).skip_reason})"
            return diff

        # 컬럼 매핑 (소문자 정규화 기반)
        col_map_old_to_new = self._match_columns(old.df.columns.tolist(), new.df.columns.tolist())
        col_map_new_to_old = {v: k for k, v in col_map_old_to_new.items()}

        diff.added_cols = [c for c in new.df.columns if c not in col_map_new_to_old]
        diff.removed_cols = [c for c in old.df.columns if c not in col_map_old_to_new]
        diff.common_cols = [c for c in old.df.columns if c in col_map_old_to_new]

        if diff.added_cols:
            self.logger.debug(f"  [{old.name}] 추가된 열: {diff.added_cols}")
        if diff.removed_cols:
            self.logger.debug(f"  [{old.name}] 삭제된 열: {diff.removed_cols}")

        # 행 키 맵 생성
        old_key_map = self._build_key_map(old.df, old.key_columns, old.name)
        new_key_map = self._build_key_map(new.df, new.key_columns, new.name)

        old_keys = set(old_key_map)
        new_keys = set(new_key_map)

        # 삭제된 행
        for key in sorted(old_keys - new_keys):
            row_idx = old_key_map[key]
            values = old.df.iloc[row_idx].to_dict()
            diff.deleted_rows.append(RowChange(sheet=old.name, row_key=key,
                                                kind="deleted", values=values))

        # 추가된 행
        for key in sorted(new_keys - old_keys):
            row_idx = new_key_map[key]
            values = new.df.iloc[row_idx].to_dict()
            diff.added_rows.append(RowChange(sheet=old.name, row_key=key,
                                              kind="added", values=values))

        # 공통 행 셀 단위 비교
        for key in sorted(old_keys & new_keys):
            old_row = old.df.iloc[old_key_map[key]]
            new_row = new.df.iloc[new_key_map[key]]

            for old_col in diff.common_cols:
                new_col = col_map_old_to_new[old_col]
                old_val = old_row.get(old_col, "")
                new_val = new_row.get(new_col, "")

                change = self._compare_cell(old.name, key, old_col, old_val, new_val)
                if change is not None:
                    diff.cell_changes.append(change)

        return diff

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _match_columns(self, old_cols: list[str], new_cols: list[str]) -> dict[str, str]:
        """old 컬럼명 → new 컬럼명 매핑 (소문자+공백제거 기준)"""
        new_lower = {c.lower().replace(" ", ""): c for c in new_cols}
        mapping: dict[str, str] = {}
        for oc in old_cols:
            key = oc.lower().replace(" ", "")
            if key in new_lower:
                mapping[oc] = new_lower[key]
        return mapping

    def _build_key_map(self, df: pd.DataFrame, key_cols: list[str],
                       sheet_name: str) -> dict[str, int]:
        """키 컬럼 값들을 ' | '로 연결 → {키문자열: 행인덱스}"""
        result: dict[str, int] = {}
        # key_cols 중 실제 df에 있는 것만 사용
        valid_keys = [c for c in key_cols if c in df.columns]
        if not valid_keys:
            # 첫 번째 컬럼 폴백
            valid_keys = [df.columns[0]]
            self.logger.warning(f"  [{sheet_name}] 키 컬럼이 DataFrame에 없어 첫 번째 컬럼 사용")

        for i, row in df.iterrows():
            key = " | ".join(str(row.get(c, "")) for c in valid_keys)
            if key in result:
                self.logger.warning(f"  [{sheet_name}] 중복 키 발견: '{key}' — 마지막 행 사용")
            result[key] = int(i)  # type: ignore[arg-type]
        return result

    def _compare_cell(self, sheet: str, row_key: str, col: str,
                      old_val: str, new_val: str) -> Optional[CellChange]:
        """두 셀 값을 비교해 변경이 있으면 CellChange 반환, 같으면 None"""
        if old_val == new_val:
            return None

        # 숫자 비교 시도
        try:
            old_f = float(old_val.replace(",", ""))
            new_f = float(new_val.replace(",", ""))
            delta = new_f - old_f
            if abs(delta) <= self.tolerance:
                return None  # 차이가 임계값 이하
            delta_pct = (delta / abs(old_f) * 100) if old_f != 0 else None
            return CellChange(sheet=sheet, row_key=row_key, col_name=col,
                              old_val=old_val, new_val=new_val,
                              delta=delta, delta_pct=delta_pct)
        except (ValueError, AttributeError):
            pass

        # 문자열 변경
        return CellChange(sheet=sheet, row_key=row_key, col_name=col,
                          old_val=old_val, new_val=new_val,
                          delta=None, delta_pct=None)


# ---------------------------------------------------------------------------
# WorkbookComparator
# ---------------------------------------------------------------------------

class WorkbookComparator:
    def __init__(self, config: configparser.ConfigParser, logger: logging.Logger) -> None:
        self.config = config
        self.logger = logger
        self._sheet_comp = SheetComparator(config, logger)
        skip_raw = config.get("EXCEL", "skip_sheets", fallback="").strip()
        self.skip_sheets: set[str] = (
            {s.strip() for s in skip_raw.split(",") if s.strip()} if skip_raw else set()
        )

    def compare(self,
                old_sheets: dict[str, SheetData],
                new_sheets: dict[str, SheetData]) -> WorkbookDiff:
        diff = WorkbookDiff()

        all_names = list(dict.fromkeys(list(old_sheets) + list(new_sheets)))  # 순서 보존

        for name in all_names:
            if name in self.skip_sheets:
                continue

            in_old = name in old_sheets
            in_new = name in new_sheets

            if in_old and not in_new:
                self.logger.info(f"  시트 삭제됨: [{name}]")
                diff.sheet_diffs.append(SheetDiff(sheet_name=name, old_only_sheet=True))

            elif not in_old and in_new:
                self.logger.info(f"  시트 추가됨: [{name}]")
                diff.sheet_diffs.append(SheetDiff(sheet_name=name, new_only_sheet=True))

            else:
                self.logger.info(f"  시트 비교 중: [{name}]")
                sd = self._sheet_comp.compare(old_sheets[name], new_sheets[name])
                diff.sheet_diffs.append(sd)
                self.logger.info(
                    f"    → 변경셀={len(sd.cell_changes)}, "
                    f"추가행={len(sd.added_rows)}, 삭제행={len(sd.deleted_rows)}"
                )

        return diff
