"""
ERP 조작 모듈 (설치형 Windows 앱)

우선순위:
  1. pywinauto  — Windows 접근성 API 기반 (안정적, 권장)
  2. pyautogui  — 스크린샷/좌표 기반 fallback

사용 전 반드시 config.ini 의 [ERP] 섹션을 실제 ERP에 맞게 수정하세요.
컨트롤 이름은 Microsoft Accessibility Insights (Inspect.exe) 로 확인할 수 있습니다.
"""
import os
import time
import logging
import configparser
from pathlib import Path
from dotenv import load_dotenv

# pywinauto import (설치 안 된 환경 대비)
try:
    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False

# pyautogui import
try:
    import pyautogui
    pyautogui.FAILSAFE = True   # 마우스를 좌상단 이동 시 즉시 중단
    pyautogui.PAUSE = 0.3       # 각 동작 사이 기본 딜레이
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False


class ERPController:
    """
    ERP 창 자동 조작 클래스.
    pywinauto 사용 가능 시 우선 사용, 불가 시 pyautogui fallback.
    """

    def __init__(self, config: configparser.ConfigParser, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.app = None
        self.window = None

        # .env 에서 로그인 정보 로드
        env_path = Path(__file__).parent.parent / "config" / ".env"
        load_dotenv(dotenv_path=env_path)
        self.user_id = os.getenv("ERP_USER_ID", "")
        self.password = os.getenv("ERP_PASSWORD", "")

        if not self.user_id or not self.password:
            raise ValueError(".env 파일에 ERP_USER_ID / ERP_PASSWORD가 설정되지 않았습니다.")

        self.window_title = config.get("ERP", "window_title", fallback="ERP")
        self.exe_path = config.get("ERP", "exe_path", fallback="")

        self._mode = self._detect_mode()
        self.logger.info(f"ERP 조작 모드: {self._mode}")

    def _detect_mode(self) -> str:
        if PYWINAUTO_AVAILABLE:
            return "pywinauto"
        if PYAUTOGUI_AVAILABLE:
            return "pyautogui"
        raise ImportError("pywinauto / pyautogui 중 하나가 설치되어 있어야 합니다.")

    # ──────────────────────────────────────────────
    # 공개 메서드
    # ──────────────────────────────────────────────

    def connect(self):
        """ERP 창에 연결 (실행 중이면 연결, 아니면 실행 후 연결)"""
        if self._mode == "pywinauto":
            self._connect_pywinauto()
        else:
            self._connect_pyautogui()

    def login(self):
        """ERP 로그인"""
        self.logger.info("ERP 로그인 시도")
        if self._mode == "pywinauto":
            self._login_pywinauto()
        else:
            self._login_pyautogui()
        self.logger.info("ERP 로그인 완료")

    def set_month_and_search(self, month: str):
        """
        기준월 입력 후 조회
        month: 'YYYYMM' 형식 (예: '202503')
        """
        self.logger.info(f"기준월 설정: {month}")
        if self._mode == "pywinauto":
            self._set_month_pywinauto(month)
        else:
            self._set_month_pyautogui(month)
        self.logger.info("조회 완료")

    def get_item_count(self) -> int:
        """항목 목록의 행 수 반환"""
        if self._mode == "pywinauto":
            return self._get_item_count_pywinauto()
        return self._get_item_count_pyautogui()

    def click_item(self, index: int):
        """
        항목 목록에서 index번째 항목 클릭 (0 기준)
        상세 화면이 열릴 때까지 대기
        """
        self.logger.debug(f"항목 클릭: {index}")
        if self._mode == "pywinauto":
            self._click_item_pywinauto(index)
        else:
            self._click_item_pyautogui(index)

    def close(self):
        """ERP 창 닫기 (선택적)"""
        try:
            if self._mode == "pywinauto" and self.window:
                self.window.close()
        except Exception:
            pass

    # ──────────────────────────────────────────────
    # pywinauto 구현
    # ──────────────────────────────────────────────

    def _connect_pywinauto(self):
        try:
            # 이미 실행 중인 창에 연결 시도
            self.app = Application(backend="uia").connect(title_re=f".*{self.window_title}.*")
            self.logger.info("실행 중인 ERP 창에 연결")
        except Exception:
            # 실행 중이 아니면 exe 실행
            if not self.exe_path or not Path(self.exe_path).exists():
                raise FileNotFoundError(
                    f"ERP 실행 파일을 찾을 수 없습니다: {self.exe_path}\n"
                    "config.ini 의 exe_path 를 확인하세요."
                )
            self.logger.info(f"ERP 실행: {self.exe_path}")
            self.app = Application(backend="uia").start(self.exe_path)
            time.sleep(3)  # 실행 대기

        self.window = self.app.window(title_re=f".*{self.window_title}.*")
        self.window.wait("visible", timeout=15)

    def _login_pywinauto(self):
        # ※ 아래 컨트롤 이름은 config.ini 에서 읽음
        id_ctrl = self.config.get("ERP", "login_id_control")
        pw_ctrl = self.config.get("ERP", "login_pw_control")
        btn_ctrl = self.config.get("ERP", "login_btn_control")

        self.window[id_ctrl].set_edit_text(self.user_id)
        self.window[pw_ctrl].set_edit_text(self.password)
        self.window[btn_ctrl].click()
        time.sleep(2)

    def _set_month_pywinauto(self, month: str):
        month_ctrl = self.config.get("ERP", "month_input_control")
        search_ctrl = self.config.get("ERP", "search_btn_control")

        self.window[month_ctrl].set_edit_text(month)
        self.window[search_ctrl].click()
        time.sleep(2)

    def _get_item_count_pywinauto(self) -> int:
        list_ctrl = self.config.get("ERP", "item_list_control")
        grid = self.window[list_ctrl]
        return grid.item_count()

    def _click_item_pywinauto(self, index: int):
        list_ctrl = self.config.get("ERP", "item_list_control")
        grid = self.window[list_ctrl]
        grid.get_item(index).double_click_input()
        time.sleep(1.5)

    # ──────────────────────────────────────────────
    # pyautogui fallback 구현 (좌표 기반)
    # ──────────────────────────────────────────────
    # ※ pyautogui 방식은 화면 해상도와 ERP 창 위치에 의존합니다.
    #   실제 ERP 화면을 보면서 좌표를 직접 측정 후 아래 값을 수정하세요.
    #   pyautogui.locateOnScreen('image.png') 방식을 권장합니다.

    def _connect_pyautogui(self):
        import subprocess
        if self.exe_path and Path(self.exe_path).exists():
            subprocess.Popen(self.exe_path)
            time.sleep(3)
        self.logger.info("pyautogui 모드: ERP 창이 화면에 표시되어 있어야 합니다.")

    def _login_pyautogui(self):
        # TODO: 실제 ERP 화면에서 좌표 확인 후 수정
        self.logger.warning("pyautogui 로그인: 아래 좌표를 실제 ERP 화면에 맞게 수정하세요.")
        pyautogui.click(x=960, y=400)   # ID 입력란
        pyautogui.typewrite(self.user_id, interval=0.05)
        pyautogui.click(x=960, y=450)   # PW 입력란
        pyautogui.typewrite(self.password, interval=0.05)
        pyautogui.click(x=960, y=500)   # 로그인 버튼
        time.sleep(2)

    def _set_month_pyautogui(self, month: str):
        self.logger.warning("pyautogui 기준월 설정: 아래 좌표를 실제 ERP 화면에 맞게 수정하세요.")
        pyautogui.click(x=400, y=200)   # 기준월 입력란
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(month, interval=0.05)
        pyautogui.click(x=600, y=200)   # 조회 버튼
        time.sleep(2)

    def _get_item_count_pyautogui(self) -> int:
        # pyautogui 방식에서는 화면 파싱이 어려우므로 수동 입력 요청
        count = input("항목 목록의 행 수를 직접 입력하세요: ")
        return int(count)

    def _click_item_pyautogui(self, index: int):
        self.logger.warning("pyautogui 항목 클릭: 아래 좌표를 실제 ERP 화면에 맞게 수정하세요.")
        # 예: 첫 번째 항목 y=300, 행 높이=25 가정
        base_y = 300
        row_height = 25
        pyautogui.doubleClick(x=500, y=base_y + index * row_height)
        time.sleep(1.5)
