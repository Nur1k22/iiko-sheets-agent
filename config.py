# ─────────────────────────────────────────
#  НАСТРОЙКИ — заполни перед запуском
# ─────────────────────────────────────────

# Папка для хранения файлов агента (логи, токен, credentials)
EXPORT_FOLDER = r"C:\iiko_exports"

# ID Google Таблицы (из URL: .../spreadsheets/d/ВОТ_ЭТОТ_ID/...)
SPREADSHEET_ID = "ВАШ_SPREADSHEET_ID"

# Название листа в Google Таблице (или префикс — агент найдёт сам)
SHEET_NAME_PREFIX = "Почасовой ТО"

# Путь к файлу credentials от Google (скачать из Google Cloud Console)
GOOGLE_CREDENTIALS_JSON = r"C:\iiko_exports\google_credentials.json"

# Путь к сохранённому токену (создаётся автоматически при первом запуске)
GOOGLE_TOKEN_PICKLE = r"C:\iiko_exports\token.pickle"

# Заголовок окна iiko (обычно содержит "iikoChain")
IIKO_WINDOW_TITLE = "iikoChain"

# Папка, куда iiko сохраняет временные Excel-файлы
# Замени "ИМЯ_ПОЛЬЗОВАТЕЛЯ" на своё имя пользователя Windows
IIKO_TEMP_FOLDER = r"C:\Users\ИМЯ_ПОЛЬЗОВАТЕЛЯ\AppData\Local\Temp\Resto"

# Координаты кнопки Excel в окне iiko
# Запусти: python iiko_agent.py calibrate — и наведи мышь на кнопку
EXCEL_BUTTON_X = 1836
EXCEL_BUTTON_Y = 157

# Расписание: в какую минуту каждого часа делать экспорт (":00" = начало часа)
SCHEDULE_AT_MINUTE = ":00"
