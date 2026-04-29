# PresentAI

> **AI-генератор презентаций:** от текстового запроса или документа до готового `.pptx` за несколько секунд.

🌐 **Хостинг:** [https://presentai-y64a.onrender.com/](https://presentai-y64a.onrender.com/)

---

## Описание продукта

PresentAI — это веб-сервис, который автоматически создаёт презентации в формате `.pptx` на основе текстового промпта или загруженного документа (PDF / DOCX). Пользователь задаёт тему, количество слайдов, тон подачи и визуальную тему — LLM формирует структуру, сервер собирает готовый файл и показывает редактируемое превью прямо в браузере.

Проект создан в рамках хакатона Ростелекома по кейсу «AI-генератор презентаций: от промпта до готового PPTX».

---

## Стек технологий

| Слой | Технологии |
|------|-----------|
| **Backend** | Python 3.10+, FastAPI, Uvicorn |
| **Генерация PPTX** | python-pptx |
| **LLM** | Ростелеком LLM API (llama), Anthropic Claude (fallback) |
| **Генерация изображений** | Stable Diffusion (RT API), Yandex ART API |
| **Парсинг документов** | PyPDF2 / pdfplumber, python-docx |
| **Frontend** | Jinja2-шаблоны, Vanilla JS, HTML/CSS |
| **Хостинг** | Render.com |

---

## Архитектура приложения

```
Пользователь (браузер)
        │
        ▼
┌─────────────────────────────────────────┐
│              FastAPI (app.py)           │
│                                         │
│  GET  /               веб-интерфейс     │
│  POST /api/generate   генерация         │
│  GET  /result/{id}    превью + редактор │
│  POST /api/rebuild    пересборка PPTX   │
│  GET  /api/download   скачивание        │
└────────────┬──────────────┬─────────────┘
             │              │
     ┌───────▼──────┐  ┌────▼────────────┐
     │   LLM Layer  │  │  Image Layer    │
     │              │  │                 │
     │ RT LLM API   │  │ Stable Diffusion│
     │ Anthropic    │  │ Yandex ART API  │
     │ local fallback│  │ circuit breaker│
     └───────┬──────┘  └────┬────────────┘
             │              │
             ▼              ▼
     ┌────────────────────────────┐
     │      python-pptx           │
     │  сборка .pptx по структуре │
     └────────────────────────────┘
             │
             ▼
     temp_files/{session_id}.pptx
```

**Ключевые решения:**
- **Circuit breaker** — при 2 ошибках подряд image-бэкенд отключается на весь батч, PPTX собирается без изображений и без зависания
- **Fallback-цепочка** — RT LLM → Anthropic → локальная демо-структура
- **Rate limit handling** — обработка HTTP 429 с паузой и повторной попыткой
- **Voice-mode** — голосовой ввод в поле описание презентации
- **Взможность редактирования** — после генерации презентации пользователя перенаправляет на страницу где он может посмотреть как примерно будет выглядеть презентация и при желании изменить текст, картинку
- **Выбор цветовой темы** - пользователь может выбрать одну из предложенных 6 тем
- **Пересборка и скачавание .pptx** - после того как пользователь отредактировал презентацию он может скачать текущий вариант .pptx

---

## Возможности

- генерация структуры слайдов через RT LLM API (`/api/1.0/llama/chat`);
- локальный fallback-режим без API-ключей;
- извлечение текста из PDF и DOCX;
- визуальные темы: Глубокий неон, Тёмная волна, Деловая, Минимализм, Технологичная, Креативная;
- опциональная генерация изображений через Stable Diffusion или Yandex ART API (первый и последний слайды без изображений);
- автоматический fallback с Yandex ART на Stable Diffusion при недоступности сервиса;
- circuit breaker: при 2 ошибках подряд бэкенд отключается на весь батч, PPTX собирается без зависания;
- обработка rate limit (HTTP 429) с паузой и повторной попыткой;
- редактирование слайдов в браузере без полной регенерации презентации;
- повторная сборка и скачивание готового `.pptx`;
- отображение источника генерации и отчёта QA-проверки.

---

## Инструкция по развёртыванию в локальном контуре

### Требования

- Python 3.10 или новее
- Git

### Linux / macOS

```bash
# 1. Клонировать репозиторий
git clone https://github.com/huskkke/hackathon.git
cd hackathon

# 2. Создать виртуальное окружение
python3 -m venv .venv
.venv/bin/python -m ensurepip --upgrade
.venv/bin/python -m pip install --upgrade pip

# 3. Установить зависимости
.venv/bin/python -m pip install -r requirements.txt

# 4. (Опционально) Задать API-ключи
export ANTHROPIC_API_KEY=your_key_here

# 5. Запустить сервер
.venv/bin/python iigenerator/app.py
```

### Windows CMD

```CMD
# 1. Клонироватьерированных презентаций в количестве не менее 3 штук. Презентации должны быть созданы одним и тем  репозиторий
git clone https://github.com/huskkke/hackathon.git
cd hackathon

# 2. Установить зависимости
python.exe -m pip install -r requirements.txt

# 3. Запустить сервер
cd C:\Users\ИмяПользователя\hackathon\iigenerator
python app.py
```

После запуска откройте в браузере:

```
http://localhost:8000
```

> Если страница не обновилась после правок — нажмите `Ctrl+F5`.

---

## Переменные окружения

| Переменная | Описание | По умолчанию |
|-----------|----------|-------------|
| `ANTHROPIC_API_KEY` | Ключ Anthropic для запасной LLM-генерации | — |
| `ANTHROPIC_MODEL` | Модель Anthropic | `claude-sonnet-4-20250514` |
| `PRESENTAI_WORK_DIR` | Папка для временных `.json` и `.pptx` | `temp_files/` |
| `RT_IMAGE_MAX_IMAGES` | Максимум изображений за одну генерацию | `20` |
| `RT_IMAGE_BATCH_TIMEOUT` | Общий лимит времени на батч изображений (сек) | `270` |
| `RT_IMAGE_SINGLE_TIMEOUT` | Лимит времени на одно изображение (сек) | `70` |
| `RT_IMAGE_ALLOW_BACKEND_FALLBACK` | Fallback на второй image-backend при ошибке | `1` |
| `RT_IMAGE_RATELIMIT_PAUSE` | Пауза при HTTP 429 (сек) | `8` |
| `RT_IMAGE_CIRCUIT_BREAKER_THRESHOLD` | Ошибок подряд до отключения бэкенда в батче | `2` |

---

## Основные endpoints

| Метод | Путь | Описание |
|-------|------|----------|
| `GET` | `/` | Веб-интерфейс |
| `POST` | `/loading` | Страница генерации с прогрессом |
| `POST` | `/api/jobs/{job_id}/run` | Запуск задачи генерации |
| `GET` | `/result/{session_id}` | Результат с превью и редактором |
| `POST` | `/api/generate` | JSON API для генерации структуры и PPTX |
| `POST` | `/api/rebuild/{session_id}` | Пересборка PPTX после редактирования |
| `POST` | `/api/session/{session_id}/slide/{slide_index}/edit` | Точечное редактирование слайда |
| `GET` | `/api/download/{session_id}` | Скачивание презентации |
| `GET` | `/api/demo` | Демо-PPTX без JavaScript и API-ключей |

---

## Структура проекта

```
hackathon/
├── iigenerator/
│   ├── __init__.py
│   ├── app.py      # FastAPI-приложение, генерация, редактор и PPTX-сборка
│   └── main.py     # совместимый entrypoint для python -m iigenerator.main
├── requirements.txt
├── README.md
└── temp_files/     # runtime-файлы, игнорируются Git
```

---

## Материалы

- 📊 **Презентация, Скринкаст:** [https://disk.yandex.ru/d/8J-_rUVg0OaHvw](https://disk.yandex.ru/d/8J-_rUVg0OaHvw)

---

## Примечания

- Токены хранятся только в памяти процесса и не записываются в JSON/PPTX.
- Runtime-файлы и виртуальное окружение закрыты в `.gitignore`.
