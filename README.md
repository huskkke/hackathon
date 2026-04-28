# PresentAI

AI-генератор презентаций: от текстового запроса или документа до готового `.pptx`.

Проект принимает тему, PDF/DOCX-файл, количество слайдов, тон и визуальную тему. После этого LLM собирает структуру презентации, опционально добавляет изображения через RT API, а сервер генерирует PPTX и показывает редактируемое превью.

## Возможности

- генерация структуры слайдов через RT LLM API (`/api/1.0/llama/chat`);
- локальный fallback-режим без API-ключей;
- извлечение текста из PDF и DOCX;
- визуальные темы: Глубокий неон, Тёмная волна, Деловая, Минимализм, Технологичная, Креативная;
- опциональная генерация изображений через Stable Diffusion или Yandex ART API;
- редактирование слайдов в браузере без полной регенерации презентации;
- повторная сборка и скачивание готового `.pptx`;
- отображение источника генерации и отчёта QA-проверки.

## Быстрый старт

Нужен Python 3.10 или новее.

```bash
cd hackathon
python3 -m venv .venv
.venv/bin/python -m ensurepip --upgrade
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python iigenerator/app.py
```

После запуска откройте:

```text
http://localhost:8000
```

Если порт `8000` занят, можно запустить через `uvicorn` на другом порту:

```bash
.venv/bin/python -m uvicorn iigenerator.app:app --host 127.0.0.1 --port 8001
```

## Windows PowerShell

```powershell
cd путь\к\папке\hackathon
py -3 -m venv .venv --upgrade-deps
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe iigenerator\app.py
```

Если окружение создалось без `pip`, выполните:

```powershell
.\.venv\Scripts\python.exe -m ensurepip --upgrade
.\.venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## API-ключи

RT Bearer-токен можно вставить прямо в интерфейсе. Он используется для LLM и, если включена генерация изображений, для `/sd/img` или `/ya/image`.

Дополнительно поддерживается Anthropic как запасной LLM-провайдер:

```bash
export ANTHROPIC_API_KEY=...
export ANTHROPIC_MODEL=claude-sonnet-4-20250514
```

Если ключи не заданы или внешний API недоступен, приложение создаёт демо-структуру локально и показывает предупреждение в результате.

## Переменные окружения

- `PRESENTAI_WORK_DIR` - папка для временных `.json` и `.pptx`; по умолчанию `temp_files/`.
- `ANTHROPIC_API_KEY` - ключ Anthropic для запасной генерации.
- `ANTHROPIC_MODEL` - модель Anthropic; по умолчанию `claude-sonnet-4-20250514`.

## Основные endpoints

- `GET /` - веб-интерфейс;
- `POST /loading` - страница генерации с прогрессом;
- `POST /api/jobs/{job_id}/run` - запуск задачи генерации;
- `GET /result/{session_id}` - результат с предпросмотром и редактором;
- `POST /api/generate` - JSON API для генерации структуры и PPTX;
- `POST /api/rebuild/{session_id}` - пересборка PPTX после редактирования;
- `POST /api/session/{session_id}/slide/{slide_index}/edit` - точечное редактирование слайда;
- `GET /api/download/{session_id}` - скачивание презентации;
- `GET /api/demo` - демо-PPTX без JavaScript и API-ключей.

## Структура

```text
hackathon/
├── iigenerator/
│   ├── __init__.py
│   ├── app.py      # FastAPI-приложение, генерация, редактор и PPTX-сборка
│   └── main.py     # совместимый entrypoint для python -m iigenerator.main
├── requirements.txt
├── README.md
└── temp_files/     # runtime-файлы, игнорируются Git
```

## Примечания

- Токены хранятся только в памяти процесса и не записываются в JSON/PPTX.
- Runtime-файлы и виртуальное окружение закрыты в `.gitignore`.
- Для чистого обновления страницы после правок используйте `Ctrl+F5`.
