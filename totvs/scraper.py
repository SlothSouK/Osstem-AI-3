import os
import configparser
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / "config" / ".env")

config = configparser.ConfigParser()
config.read(BASE_DIR / "config" / "config.ini", encoding="utf-8")

TOTVS_URL = os.getenv("TOTVS_URL")
TOTVS_USER_ID = os.getenv("TOTVS_USER_ID")
TOTVS_PASSWORD = os.getenv("TOTVS_PASSWORD")

HEADLESS = config.getboolean("SCRAPER", "headless", fallback=False)
TIMEOUT = config.getint("SCRAPER", "timeout", fallback=30000)


def login(page):
    page.goto(TOTVS_URL)
    # TODO: 로그인 셀렉터 추가
    pass


def fetch_data(page):
    # TODO: 데이터 수집 로직 추가
    pass


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        page.set_default_timeout(TIMEOUT)

        login(page)
        fetch_data(page)

        browser.close()


if __name__ == "__main__":
    main()
