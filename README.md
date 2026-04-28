# PresentAI

Прототип для кейса Ростелекома «AI-генератор презентаций: от промпта до готового PPTX».

Сервис принимает текстовый запрос, PDF или DOCX документ, количество слайдов, стиль и тон. LLM формирует структуру презентации, опционально генерируются изображения через RT API, после чего сервер собирает PPTX и показывает редактируемое превью.

## Возможности

- генерация структуры слайдов через Ростелеком LLM API (`/api/1.0/llama/chat`);
- fallback-режим без ключа API, чтобы прототип можно было показать локально;
- загрузка PDF и DOCX как источника контента;
- стили презентации: Ростелеком, Modern Dark, Corporate, Minimal, Tech;
- опциональная генерация изображений через Stable Diffusion или Yandex ART API;
- редактирование заголовков и текста слайдов в веб-интерфейсе;
- большой предпросмотр выбранного слайда перед скачиванием;
- отображение источника генерации: `RT LLM`, `Anthropic` или локальный fallback;
- пересборка и скачивание готового `.pptx`.

## Запуск

Нужен Python 3.10 или новее. Перед запуском перейдите в папку проекта.

### Windows PowerShell

```CMD
cd путь\к\папке\hackathon
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt (Либо устанавливаете вручную через requirements.txt)
python -m uvicorn iigenerator.app:app --reload
```

Если PowerShell не разрешает запускать скрипты виртуального окружения, используйте команды выше именно через `.\.venv\Scripts\python.exe` - активация окружения не требуется.

### Linux

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python iigenerator/app.py
```

### macOS

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python iigenerator/app.py
```

После запуска откройте:

```text
http://localhost:8000
```

Если страница в браузере не обновилась после изменений в коде, откройте `http://localhost:8000/?v=10` или нажмите `Ctrl+F5`.

## API-ключи

В интерфейсе можно вставить Bearer-токен Ростелеком API. Он используется для LLM и, если включена галочка генерации изображений, для `/sd/img` или `/ya/image`.

Дополнительно поддерживается Anthropic как запасной LLM-провайдер:

```bash
export ANTHROPIC_API_KEY=...
```

Если токены не заданы или API недоступен, сервис создаст демо-структуру презентации локально и покажет предупреждение в результате.

## Основные endpoints

- `GET /` - веб-интерфейс;
- `POST /loading` - отдельная страница загрузки с процентами;
- `POST /api/jobs/{job_id}/run` - запуск задачи генерации со страницы загрузки;
- `GET /result/{session_id}` - страница результата с предпросмотром и скачиванием;
- `POST /api/generate` - генерация структуры и PPTX;
- `POST /api/rebuild/{session_id}` - пересборка PPTX после редактирования;
- `GET /api/download/{session_id}` - скачивание презентации.

Временные PPTX и JSON-файлы сохраняются в `/tmp/presentai`.
