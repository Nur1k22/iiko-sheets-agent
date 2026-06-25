"""
iiko → Google Sheets Agent
Каждый час нажимает кнопку Excel в iikoChain, читает экспорт и пишет в Google Sheets.
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
import sys
import ctypes

import config

# ─────────────────────────────────────────
#  SCOPES
# ─────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ─────────────────────────────────────────
#  ЛОГИРОВАНИЕ
# ─────────────────────────────────────────

os.makedirs(config.EXPORT_FOLDER, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(
            os.path.join(config.EXPORT_FOLDER, "iiko_agent.log"),
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5


# ─────────────────────────────────────────
#  GOOGLE АВТОРИЗАЦИЯ
# ─────────────────────────────────────────

def get_google_credentials():
    creds = None
    if os.path.exists(config.GOOGLE_TOKEN_PICKLE):
        with open(config.GOOGLE_TOKEN_PICKLE, "rb") as f:
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
            if not os.path.exists(config.GOOGLE_CREDENTIALS_JSON):
                raise FileNotFoundError(
                    f"Файл credentials не найден: {config.GOOGLE_CREDENTIALS_JSON}\n"
                    "Скачай его из Google Cloud Console и положи в папку iiko_exports."
                )
            log.info("Открываю браузер для авторизации Google...")
            flow = InstalledAppFlow.from_client_secrets_file(
                config.GOOGLE_CREDENTIALS_JSON, SCOPES
            )
            creds = flow.run_local_server(port=0)
            with open(config.GOOGLE_TOKEN_PICKLE, "wb") as f:
                pickle.dump(creds, f)
            log.info("Токен сохранён.")
    return creds


# ─────────────────────────────────────────
#  РАБОТА С ОКНОМ iiko
# ─────────────────────────────────────────

def find_iiko_window():
    windows = gw.getWindowsWithTitle(config.IIKO_WINDOW_TITLE)
    if not windows:
        raise RuntimeError(
            f"Окно '{config.IIKO_WINDOW_TITLE}' не найдено. "
            "Убедись, что iiko открыт и в заголовке окна есть это слово."
        )
    win = windows[0]
    hwnd = win._hWnd
    ctypes.windll.user32.ShowWindow(hwnd, 3)
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(1.5)
    log.info(f"Окно активировано: {win.title}")
    return win


def click_excel_button():
    # Сначала пробуем найти по шаблону (если есть)
    template_path = os.path.join(config.EXPORT_FOLDER, "excel_button_template.png")
    if os.path.exists(template_path):
        try:
            location = pyautogui.locateOnScreen(template_path, confidence=0.8)
            if location:
                pyautogui.click(pyautogui.center(location))
                log.info(f"Кнопка найдена по шаблону: {location}")
                return
        except Exception as e:
            log.warning(f"Поиск по шаблону не сработал: {e}")

    # Иначе — кликаем по координатам из config
    x, y = config.EXCEL_BUTTON_X, config.EXCEL_BUTTON_Y
    log.info(f"Кликаю по координатам: ({x}, {y})")
    pyautogui.click(x, y)


def wait_and_close_excel():
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
        pyautogui.hotkey("alt", "f4")
        time.sleep(1)
        pyautogui.press("n")  # Не сохранять
        time.sleep(1)
    else:
        log.warning("Окно Excel не найдено — возможно, файл уже закрылся")


# ─────────────────────────────────────────
#  ЧТЕНИЕ ФАЙЛА
# ─────────────────────────────────────────

def find_latest_iiko_file():
    patterns = [
        os.path.join(config.IIKO_TEMP_FOLDER, "*.xlsx"),
        r"C:\Users\*\AppData\Local\Temp\Resto\*.xlsx",
        r"C:\Users\*\AppData\Local\Temp\*.xlsx",
    ]
    all_files = []
    for pattern in patterns:
        all_files.extend(glob.glob(pattern))

    if not all_files:
        log.error("Excel-файл iiko не найден в папке Temp. Проверь IIKO_TEMP_FOLDER в config.py")
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


# ─────────────────────────────────────────
#  ЗАПИСЬ В GOOGLE SHEETS
# ─────────────────────────────────────────

def update_google_sheets(rows):
    log.info("Подключаюсь к Google Sheets...")
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(config.SPREADSHEET_ID)

    # Ищем лист по точному имени или по префиксу
    target_sheet = None
    for ws in spreadsheet.worksheets():
        if ws.title.strip() == config.SHEET_NAME_PREFIX.strip():
            target_sheet = ws
            break
        if ws.title.strip().startswith(config.SHEET_NAME_PREFIX.strip()):
            target_sheet = ws

    if target_sheet is None:
        target_sheet = spreadsheet.sheet1
        log.warning(
            f"Лист с префиксом '{config.SHEET_NAME_PREFIX}' не найден. "
            f"Использую первый лист: '{target_sheet.title}'"
        )
    else:
        log.info(f"Записываю в лист: '{target_sheet.title}'")

    target_sheet.clear()
    timestamp_str = f"Обновлено: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}"
    target_sheet.update(range_name="A1", values=[[timestamp_str]])
    if rows:
        target_sheet.update(range_name="A2", values=rows)
    log.info(f"Google Sheets обновлён: {len(rows)} строк")


# ─────────────────────────────────────────
#  ОСНОВНОЙ ЦИКЛ
# ─────────────────────────────────────────

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
    """Вспомогательный режим: показывает координаты мыши для config.py"""
    print("Наведи мышь на кнопку Excel в iiko и запомни координаты.")
    print("Для выхода нажми Ctrl+C.\n")
    try:
        while True:
            x, y = pyautogui.position()
            print(f"  X={x}, Y={y}   ", end="\r")
            time.sleep(0.2)
    except KeyboardInterrupt:
        print(f"\n\nЗапиши в config.py:\n  EXCEL_BUTTON_X = {x}\n  EXCEL_BUTTON_Y = {y}")


def list_sheets():
    """Вспомогательный режим: показывает все листы в Google Таблице"""
    print("Получаю список листов...")
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(config.SPREADSHEET_ID)
    print(f"\nЛисты в таблице '{spreadsheet.title}':")
    for ws in spreadsheet.worksheets():
        print(f"  {repr(ws.title)}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        cmd = sys.argv[1]
        if cmd == "calibrate":
            calibrate_coordinates()
        elif cmd == "once":
            run_export()
        elif cmd == "sheets":
            list_sheets()
        else:
            print(f"Неизвестная команда: {cmd}")
            print("Доступные команды: once | calibrate | sheets")
    else:
        log.info(f"Агент запущен. Экспорт каждый час в {config.SCHEDULE_AT_MINUTE}.")
        run_export()
        schedule.every().hour.at(config.SCHEDULE_AT_MINUTE).do(run_export)
        while True:
            schedule.run_pending()
            time.sleep(5)
