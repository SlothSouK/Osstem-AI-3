"""
엑셀 다운로드 모듈
ERP 상세 화면에서 다운로드 버튼을 클릭하고
파일이 완성될 때까지 대기 후 raw 폴더로 이동
"""
import shutil
import time
import logging
import configparser
from pathlib import Path

from .utils import wait_for_file, retry


class Downloader:
    def __init__(self, config: configparser.ConfigParser, logger: logging.Logger):
        self.config = config
        self.logger = logger

        self.download_dir = Path(config.get("PATHS", "download_dir"))
        self.raw_dir = Path(config.get("PATHS", "raw_dir"))
        self.raw_dir.mkdir(parents=True, exist_ok=True)

        self.download_btn_ctrl = config.get("ERP", "download_btn_control", fallback="btnExcelDownload")

    def download_all(self, erp_controller, item_count: int, month: str) -> list[Path]:
        """
        모든 항목을 순회하며 엑셀 다운로드
        erp_controller: ERPController 인스턴스
        item_count: 항목 수
        month: 기준월 (파일명에 사용)
        반환: 저장된 파일 경로 목록
        """
        collected = []
        failed_items = []

        for i in range(item_count):
            self.logger.info(f"[{i + 1}/{item_count}] 항목 다운로드 중")
            try:
                # 항목 클릭 → 상세 화면 진입
                erp_controller.click_item(i)

                # 다운로드 실행 및 파일 수집
                path = self._download_single(erp_controller, index=i + 1, month=month)
                collected.append(path)
                self.logger.info(f"  저장 완료: {path.name}")

                # 상세 화면 닫기 (뒤로가기 or 목록으로 이동)
                self._back_to_list(erp_controller)

            except Exception as e:
                self.logger.error(f"  항목 {i + 1} 다운로드 실패: {e}")
                failed_items.append(i + 1)
                self._back_to_list(erp_controller)  # 실패해도 목록으로 복귀

        if failed_items:
            self.logger.warning(f"다운로드 실패 항목: {failed_items}")

        self.logger.info(f"다운로드 완료: {len(collected)}건 성공 / {len(failed_items)}건 실패")
        return collected

    @retry(max_attempts=3, delay=2.0)
    def _download_single(self, erp_controller, index: int, month: str) -> Path:
        """단일 항목 다운로드 후 raw 폴더로 이동"""
        # 다운로드 버튼 클릭
        self._click_download_btn(erp_controller)

        # 파일 생성 대기
        downloaded = wait_for_file(self.download_dir, timeout=30.0)

        # raw 폴더로 이동 + 이름 변경 (월_순번.xlsx)
        dest = self.raw_dir / f"{month}_{index:03d}.xlsx"
        shutil.move(str(downloaded), str(dest))
        return dest

    def _click_download_btn(self, erp_controller):
        """ERP 상세 화면의 다운로드 버튼 클릭"""
        mode = erp_controller._mode

        if mode == "pywinauto":
            window = erp_controller.window
            window[self.download_btn_ctrl].click()
            time.sleep(1)
        else:
            import pyautogui
            # TODO: 실제 다운로드 버튼 좌표로 수정
            erp_controller.logger.warning("pyautogui 다운로드 버튼: 좌표를 실제 ERP에 맞게 수정하세요.")
            pyautogui.click(x=800, y=150)
            time.sleep(1)

    def _back_to_list(self, erp_controller):
        """상세 화면에서 항목 목록으로 돌아가기"""
        mode = erp_controller._mode

        if mode == "pywinauto":
            try:
                # 창 닫기 또는 뒤로가기 버튼 시도
                window = erp_controller.window
                # 방법 1: ESC 키
                from pywinauto.keyboard import send_keys
                send_keys("{ESC}")
                time.sleep(0.8)
            except Exception:
                pass
        else:
            import pyautogui
            pyautogui.press("escape")
            time.sleep(0.8)
