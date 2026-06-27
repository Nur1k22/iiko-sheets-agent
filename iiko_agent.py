"""
iiko → Google Sheets Agent
Каждый час: нажимает Excel в iiko Office → читает файл из Temp → пишет в Google Sheets
Авторизация: Apps Script Web App (без Google Cloud Console)
"""

import os
import time
import glob
import logging
import schedule
import pyautogui
import pygetwindow as gw
import pandas as pd
import requests
from datetime import datetime

# ─────────────────────────────────────────
#  НАСТРОЙКИ
# ─────────────────────────────────────────

EXPORT_FOLDER      = r"D:\vs\iiko_exports"
APPS_SCRIPT_URL    = "https://script.google.com/macros/s/AKfycbxbmPQm-xQCuirQbT47nRbkLbQboldbyqFfUqcDxt421BzJVnE3GgvtJb8ivkrl8jIfnA/exec"
IIKO_WINDOW_TITLE  = "iikoChain"
IIKO_TEMP_FOLDER   = r"C:\Users\nur1k\AppData\Local\Temp\Resto"

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
        pyautogui.press('n')
        time.sleep(1)
    else:
        log.warning("Окно Excel не найдено")


def find_latest_iiko_file():
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
    if APPS_SCRIPT_URL == "ВСТАВЬ_СЮДА_URL_ВЕБ_ПРИЛОЖЕНИЯ":
        log.error("APPS_SCRIPT_URL не задан! Вставь URL веб-приложения в настройки.")
        return

    log.info("Отправляю данные в Google Sheets через Apps Script...")
    try:
        response = requests.post(
            APPS_SCRIPT_URL,
            json={"rows": rows},
            timeout=60
        )
        result = response.json()
        if result.get("status") == "ok":
            log.info(f"Google Sheets обновлён: {result.get('rows')} строк")
        else:
            log.error(f"Ошибка от Apps Script: {result.get('message')}")
    except requests.exceptions.Timeout:
        log.error("Таймаут при отправке в Google Sheets")
    except Exception as e:
        log.error(f"Ошибка отправки: {e}")


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