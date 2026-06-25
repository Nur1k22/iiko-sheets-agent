"""
iiko → Google Sheets Agent
Каждый час: нажимает Excel... в iiko Office → читает файл из Temp → пишет в Google Sheets
"""

import os
import time
import glob
import logging
import schedule
import pyautogui
import pygetwindow as gw
import pandas as pd
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime
import pickle

# ─────────────────────────────────────────
#  НАСТРОЙКИ
# ─────────────────────────────────────────

EXPORT_FOLDER           = r"D:\vs\iiko_exports"
SPREADSHEET_ID          = "1U9kZMq1eJ1FtTOzl_eSevgoO_pr2ucNp4MUUxoHhWHs"
SHEET_NAME              = "Почасовой ТО"
GOOGLE_CREDENTIALS_JSON = r"D:\vs\iiko_exports\google_credentials.json"
GOOGLE_TOKEN_PICKLE     = r"D:\vs\iiko_exports\token.pickle"
IIKO_WINDOW_TITLE       = "iikoChain"
IIKO_TEMP_FOLDER        = r"C:\Users\nur1k\AppData\Local\Temp\Resto"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

os.makedirs(EXPORT_FOLDER, exist_ok=True)

# ─────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(EXPORT_FOLDER, "iiko_agent.log"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5


def get_google_credentials():
    creds = None
    if os.path.exists(GOOGLE_TOKEN_PICKLE):
        with open(GOOGLE_TOKEN_PICKLE, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            log.info("Обновляю токен...")
            try:
                creds.refresh(Request())
            except Exception as e:
                log.error(f"Ошибка обновления токена: {e}")
                creds = None

        if not creds:
            log.info("Открываю браузер для авторизации...")
            flow = InstalledAppFlow.from_client_secrets_file(
                GOOGLE_CREDENTIALS_JSON, SCOPES
            )
            creds = flow.run_local_server(port=0)
            with open(GOOGLE_TOKEN_PICKLE, "wb") as f:
                pickle.dump(creds, f)
            log.info("Токен сохранён.")
    return creds


def find_iiko_window():
    import ctypes
    windows = gw.getWindowsWithTitle(IIKO_WINDOW_TITLE)
    if not windows:
        raise RuntimeError(f"Окно '{IIKO_WINDOW_TITLE}' не найдено.")
    win = windows[0]
    hwnd = win._hWnd
    ctypes.windll.user32.ShowWindow(hwnd, 3)
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(1.5)
    log.info(f"Окно активировано: {win.title}")
    return win


def click_excel_button():
    template_path = os.path.join(EXPORT_FOLDER, "excel_button_template.png")
    if os.path.exists(template_path):
        try:
            location = pyautogui.locateOnScreen(template_path, confidence=0.8)
            if location:
                pyautogui.click(pyautogui.center(location))
                log.info(f"Кнопка найдена по шаблону: {location}")
                return
        except Exception as e:
            log.warning(f"Ошибка поиска по шаблону: {e}")

    x, y = 1836, 157
    log.info(f"Кликаю по координатам: ({x}, {y})")
    pyautogui.click(x, y)


def wait_and_close_excel():
    """Ждём открытия Excel, потом закрываем его."""
    import ctypes
    log.info("Жду открытия Excel...")
    time.sleep(4)

    excel_windows = gw.getWindowsWithTitle("Excel")
    if not excel_windows:
        excel_windows = gw.getWindowsWithTitle("Microsoft Excel")

    if excel_windows:
        log.info(f"Excel найден: {excel_windows[0].title}")
        hwnd = excel_windows[0]._hWnd
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.5)
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)
        # Если спросит "Сохранить?" — не сохранять
        pyautogui.press('n')
        time.sleep(1)
    else:
        log.warning("Окно Excel не найдено")


def find_latest_iiko_file():
    """Находит последний xlsx файл в папке Temp iiko."""
    patterns = [
        os.path.join(IIKO_TEMP_FOLDER, "*.xlsx"),
        r"C:\Users\nur1k\AppData\Local\Temp\*.xlsx",
    ]
    all_files = []
    for pattern in patterns:
        all_files.extend(glob.glob(pattern))

    if not all_files:
        log.error("Файлы iiko не найдены в Temp")
        return None

    latest = max(all_files, key=os.path.getmtime)
    log.info(f"Найден файл: {latest}")
    return latest


def read_excel(filepath):
    log.info(f"Читаю Excel: {filepath}")
    df = pd.read_excel(filepath, header=None, engine="openpyxl")
    df = df.dropna(how="all")
    rows = df.fillna("").astype(str).values.tolist()
    log.info(f"Прочитано строк: {len(rows)}")
    return rows


def update_google_sheets(rows):
    log.info("Подключаюсь к Google Sheets...")
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    sheet.clear()
    timestamp_str = f"Обновлено: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}"
    sheet.update(range_name="A1", values=[[timestamp_str]])
    if rows:
        sheet.update(range_name="A2", values=rows)
    log.info(f"Google Sheets обновлён: {len(rows)} строк")


def run_export():
    log.info("=" * 60)
    log.info("ЗАПУСК ЭКСПОРТА")
    try:
        find_iiko_window()
        click_excel_button()
        wait_and_close_excel()
        time.sleep(2)

        filepath = find_latest_iiko_file()
        if not filepath:
            return

        rows = read_excel(filepath)
        update_google_sheets(rows)
        log.info("Экспорт завершён успешно.")
    except Exception as e:
        log.error(f"ОШИБКА: {e}", exc_info=True)


def calibrate_coordinates():
    print("Наведи мышь на кнопку Excel... Ctrl+C для выхода.")
    try:
        while True:
            x, y = pyautogui.position()
            print(f"  X={x}, Y={y}   ", end="\r")
            time.sleep(0.2)
    except KeyboardInterrupt:
        print("\nГотово!")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "calibrate":
        calibrate_coordinates()
    elif len(sys.argv) > 1 and sys.argv[1] == "once":
        run_export()
    else:
        log.info("Агент запущен. Экспорт каждый час.")
        run_export()
        schedule.every().hour.at(":00").do(run_export)
        while True:
            schedule.run_pending()
            time.sleep(5)