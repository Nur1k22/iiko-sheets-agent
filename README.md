# iiko → Google Sheets Agent

Автоматический агент: каждый час нажимает кнопку Excel в **iikoChain**, читает экспортированный файл и записывает данные в **Google Sheets**.

---

## Что нужно для установки

- Windows 10/11
- Python 3.10 или новее → [скачать](https://www.python.org/downloads/)
- Открытое приложение iikoChain с отчётом на экране
- Доступ к интернету

---

## Установка (пошагово)

### Шаг 1 — Скачать проект

```
git clone https://github.com/ВАШ_ЛОГИН/iiko-sheets-agent.git
cd iiko-sheets-agent
```

Или просто скачай ZIP с GitHub и распакуй.

---

### Шаг 2 — Создать виртуальное окружение и установить библиотеки

```
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

---

### Шаг 3 — Получить credentials от Google

1. Открой [Google Cloud Console](https://console.cloud.google.com/)
2. Создай новый проект (или выбери существующий)
3. Включи **Google Sheets API** и **Google Drive API**:
   - Слева → "APIs & Services" → "Library"
   - Найди и включи оба API
4. Создай credentials:
   - Слева → "APIs & Services" → "Credentials"
   - "+ Create Credentials" → "OAuth client ID"
   - Тип приложения: **Desktop app**
   - Скачай JSON-файл
5. Переименуй скачанный файл в `google_credentials.json`
6. Положи его в папку `C:\iiko_exports\`

---

### Шаг 4 — Настроить config.py

Открой файл `config.py` и заполни:

```python
# ID таблицы из URL: docs.google.com/spreadsheets/d/ЭТОТ_ID/edit
SPREADSHEET_ID = "сюда_вставь_id_таблицы"

# Твоё имя пользователя Windows
IIKO_TEMP_FOLDER = r"C:\Users\ТВОЁ_ИМЯ\AppData\Local\Temp\Resto"
```

Остальные настройки можно оставить по умолчанию.

---

### Шаг 5 — Узнать координаты кнопки Excel (если нужно)

Открой iiko с нужным отчётом и запусти:

```
python iiko_agent.py calibrate
```

Наведи мышь на кнопку Excel — программа покажет координаты. Запиши их в `config.py`:

```python
EXCEL_BUTTON_X = 1836
EXCEL_BUTTON_Y = 157
```

---

### Шаг 6 — Проверить названия листов в Google Таблице

```
python iiko_agent.py sheets
```

Убедись, что в выводе есть лист, начинающийся с `"Почасовой ТО"`.  
Если название другое — поменяй `SHEET_NAME_PREFIX` в `config.py`.

---

### Шаг 7 — Тестовый запуск

```
python iiko_agent.py once
```

При первом запуске откроется браузер — войди в Google-аккаунт и дай разрешения.  
Токен сохранится автоматически, в следующий раз браузер не откроется.

---

### Шаг 8 — Запуск в фоновом режиме (каждый час)

```
python iiko_agent.py
```

Агент будет работать и делать экспорт каждый час в начале часа (`:00`).

---

## Автозапуск вместе с Windows (опционально)

Чтобы агент запускался автоматически при старте Windows:

1. Создай файл `start_agent.bat`:
```bat
@echo off
cd /d C:\путь\к\проекту
call venv\Scripts\activate
python iiko_agent.py
```

2. Нажми `Win + R` → введи `shell:startup`
3. Положи ярлык `start_agent.bat` в открывшуюся папку

---

## Команды

| Команда | Что делает |
|---|---|
| `python iiko_agent.py` | Запуск агента (каждый час) |
| `python iiko_agent.py once` | Один экспорт прямо сейчас |
| `python iiko_agent.py calibrate` | Показывает координаты мыши |
| `python iiko_agent.py sheets` | Показывает листы в таблице |

---

## Структура файлов

```
iiko-sheets-agent/
├── iiko_agent.py          # Основной скрипт
├── config.py              # Все настройки здесь
├── requirements.txt       # Зависимости Python
├── README.md              # Эта инструкция
└── .gitignore             # Исключает секретные файлы из Git

C:\iiko_exports\           # Рабочая папка (создаётся автоматически)
├── google_credentials.json  # Скачать из Google Cloud Console
├── token.pickle             # Создаётся автоматически
├── iiko_agent.log           # Лог работы агента
└── excel_button_template.png  # (опционально) скриншот кнопки
```

---

## Частые проблемы

**`WorksheetNotFound`** — название листа не совпадает.  
→ Запусти `python iiko_agent.py sheets` и проверь точное название.

**`Окно 'iikoChain' не найдено`** — iiko не открыт или свёрнут.  
→ Открой iiko и убедись, что в заголовке есть слово `iikoChain`.

**Excel-файл не найден** — неверный путь к Temp.  
→ Проверь `IIKO_TEMP_FOLDER` в `config.py`. Подставь своё имя пользователя Windows.

**Кнопка кликается не туда** — другое разрешение экрана.  
→ Запусти `python iiko_agent.py calibrate` и обнови координаты в `config.py`.
