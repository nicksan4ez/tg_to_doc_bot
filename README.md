# TG_TO_DOCX_BOT

Телеграм-бот конвертирует:
- входящее сообщение с форматированием -> DOCX
- входящий документ `.docx` -> форматированный пост в Telegram

## Требования
- Python 3.10+

## Установка
```
python3 -m venv .venv
source .venv/bin/activate
pip3 install -r requirements.txt
```

## Запуск
```
python3 bot.py
```
Описание переменных:
- `BOT_TOKEN` — токен бота
- `DOCX_FILENAME` — имя выходного DOCX файла (по умолчанию `message.docx`)
- `DOCX_FILENAME_MAX` — максимальная длина имени, если оно берётся из текста (по умолчанию `60`)
- `ALLOWED_USER_IDS` — список разрешённых ID через запятую (если пусто — доступ открыт)

## Docker
```
docker build -t tg_to_docx_bot .
docker run --env-file .env --restart unless-stopped tg_to_docx_bot
```

Или через docker-compose:
```
docker compose up -d --build
```

## Поведение
- Шрифт DOCX: Times New Roman, 14 pt
- Выравнивание: по ширине
- Отступ первой строки: 1.25 см
- Межстрочный интервал: «точно», 18 пунктов
- Отступы до/после абзаца: 0 pt
- Каждая новая строка в сообщении -> новый абзац

## Примечания
- Ссылки из Telegram переносятся в DOCX как гиперссылки.
- Форматирование из DOCX в Telegram поддерживает: жирный, курсив, подчёркивание, зачёркивание, ссылки.
