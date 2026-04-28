import os, uuid, json, io, time, re, base64, subprocess, traceback
from pathlib import Path
from typing import Optional, List
import requests
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI(title="PresentAI Pro v3")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

WORK_DIR = Path("outputs")
WORK_DIR.mkdir(exist_ok=True)
RT_API_BASE = "https://ai.rt.ru/api/1.0"


# --- МОДЕЛИ ДАННЫХ ---
class Slide(BaseModel):
    title: str
    content: str
    image_prompt: str


class PresentationData(BaseModel):
    presentation_title: str
    slides: List[Slide]
    rt_token: str
    use_images: bool
    style: str


# --- ИНТЕРФЕЙС (HTML/JS) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="utf-8">
    <title>PresentAI Pro | Editor Mode</title>
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&family=Inter:wght@300;400;600&display=swap" rel="stylesheet">
    <style>
        :root { --primary: #6366f1; --accent: #ec4899; --glass: rgba(255, 255, 255, 0.95); }
        body { 
            font-family: 'Inter', sans-serif; background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%);
            color: #1e293b; min-height: 100vh; margin: 0; padding: 40px; display: flex; justify-content: center;
        }
        .container { 
            background: var(--glass); backdrop-filter: blur(20px); padding: 40px; 
            border-radius: 30px; width: 100%; max-width: 900px; box-shadow: 0 25px 50px rgba(0,0,0,0.5);
            animation: fadeIn 0.6s ease-out;
        }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }

        h1 { font-family: 'Montserrat'; font-weight: 900; background: linear-gradient(90deg, #6366f1, #ec4899); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-size: 36px; margin: 0 0 20px 0; }

        .card { background: #f8fafc; padding: 20px; border-radius: 15px; margin-bottom: 20px; border: 1px solid #e2e8f0; transition: 0.3s; }
        .card:hover { border-color: var(--primary); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }

        input, textarea, select { width: 100%; padding: 12px; border-radius: 10px; border: 1px solid #cbd5e1; margin-top: 5px; font-family: inherit; }
        label { font-weight: 700; font-size: 12px; color: #64748b; text-transform: uppercase; }

        .btn { background: #1e293b; color: white; padding: 15px 30px; border: none; border-radius: 12px; font-weight: 700; cursor: pointer; transition: 0.3s; }
        .btn-primary { background: var(--primary); width: 100%; font-size: 18px; margin-top: 20px; }
        .btn:hover { transform: translateY(-2px); opacity: 0.9; }

        .toggle-container { display: flex; align-items: center; gap: 10px; margin: 20px 0; background: #eff6ff; padding: 15px; border-radius: 12px; }

        /* Стили для редактора */
        #editor-section { display: none; margin-top: 30px; border-top: 2px solid #e2e8f0; padding-top: 30px; }
        .slide-num { background: var(--primary); color: white; width: 30px; height: 30px; display: inline-flex; align-items: center; justify-content: center; border-radius: 50%; font-size: 14px; margin-bottom: 10px; }

        #loader { display: none; text-align: center; padding: 40px; }
        .spinner { width: 40px; height: 40px; border: 4px solid #f3f3f3; border-top: 4px solid var(--primary); border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto 15px; }
        @keyframes spin { 100% { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="container">
        <h1>PresentAI Pro <span style="font-size: 16px; opacity: 0.7;">v3.1</span></h1>

        <div id="setup-form">
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                <div class="form-group">
                    <label>Тема презентации</label>
                    <input type="text" id="prompt" placeholder="Например: Архитектура будущего">
                </div>
                <div class="form-group">
                    <label>Стиль</label>
                    <select id="style">
                        <option value="Минимализм">✨ Минимализм</option>
                        <option value="Киберпанк">🌃 Киберпанк</option>
                        <option value="Бизнес">💼 Бизнес</option>
                    </select>
                </div>
            </div>

            <div class="toggle-container">
                <input type="checkbox" id="use_images" checked style="width: 20px; height: 20px; margin: 0;">
                <label style="margin: 0; color: #1e40af;">Генерировать изображения с помощью ИИ (Yandex ART)</label>
            </div>

            <label>Токен доступа (Bearer)</label>
            <input type="text" id="rt_token" value="eyJhbGciOiJIUzM4NCJ9...">

            <button class="btn btn-primary" onclick="initDraft()">Создать черновик</button>
        </div>

        <div id="loader">
            <div class="spinner"></div>
            <h3 id="load-text">ИИ анализирует тему...</h3>
        </div>

        <div id="editor-section">
            <h2>Настройка слайдов</h2>
            <div id="slides-container"></div>
            <button class="btn btn-primary" onclick="finalize()">Собрать финальную презентацию (.pptx)</button>
        </div>
    </div>

    <script>
        let currentDraft = null;

        async function initDraft() {
            const btn = document.querySelector('.btn-primary');
            const loader = document.getElementById('loader');
            const setup = document.getElementById('setup-form');

            setup.style.display = 'none';
            loader.style.display = 'block';

            const payload = {
                prompt: document.getElementById('prompt').value,
                style: document.getElementById('style').value,
                rt_token: document.getElementById('rt_token').value,
                slide_count: 6
            };

            try {
                const res = await fetch('/api/init-draft', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                currentDraft = await res.json();
                renderEditor();
            } catch (e) {
                alert('Ошибка инициализации: ' + e);
                setup.style.display = 'block';
            } finally {
                loader.style.display = 'none';
            }
        }

        function renderEditor() {
            const container = document.getElementById('slides-container');
            const editor = document.getElementById('editor-section');
            container.innerHTML = '';
            editor.style.display = 'block';

            currentDraft.slides.forEach((slide, index) => {
                const card = document.createElement('div');
                card.className = 'card';
                card.innerHTML = `
                    <div class="slide-num">${index + 1}</div>
                    <label>Заголовок слайда</label>
                    <input type="text" value="${slide.title}" onchange="updateSlide(${index}, 'title', this.value)">
                    <label style="margin-top:10px; display:block;">Текст слайда</label>
                    <textarea rows="3" onchange="updateSlide(${index}, 'content', this.value)">${slide.content}</textarea>
                    <label style="margin-top:10px; display:block;">Промт для картинки</label>
                    <input type="text" value="${slide.image_prompt}" onchange="updateSlide(${index}, 'image_prompt', this.value)">
                `;
                container.appendChild(card);
            });
        }

        function updateSlide(idx, field, val) {
            currentDraft.slides[idx][field] = val;
        }

        async function finalize() {
            document.getElementById('editor-section').style.display = 'none';
            document.getElementById('loader').style.display = 'block';
            document.getElementById('load-text').innerText = 'Генерируем изображения и собираем файл...';

            const payload = {
                ...currentDraft,
                rt_token: document.getElementById('rt_token').value,
                use_images: document.getElementById('use_images').checked,
                style: document.getElementById('style').value
            };

            const res = await fetch('/api/finalize', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(payload)
            });
            const result = await res.json();
            window.location.href = result.download_url;
        }
    </script>
</body>
</html>
"""


# --- ЛОГИКА BACKEND ---

@app.post("/api/init-draft")
async def init_draft(data: dict):
    # ПЕРВЫЙ ШАГ: Только генерация структуры текста
    headers = {"Authorization": f"Bearer {data['rt_token']}", "Content-Type": "application/json"}

    payload = {
        "uuid": str(uuid.uuid4()),
        "chat": {
            "model": "Qwen/Qwen2.5-72B-Instruct",
            "user_message": f"Тема: {data['prompt']}. Стиль: {data['style']}. Верни JSON: {{'presentation_title': '', 'slides': [{{'title': '', 'content': '', 'image_prompt': ''}}]}}",
            "max_new_tokens": 2048
        }
    }

    try:
        resp = requests.post(f"{RT_API_BASE}/llama/chat", json=payload, headers=headers)
        resp.raise_for_status()
        content = resp.json()[0]["message"]["content"]
        return json.loads(re.sub(r"```json|```", "", content).strip())
    except Exception as e:
        raise HTTPException(500, f"Llama Error: {str(e)}")


@app.post("/api/finalize")
async def finalize(data: PresentationData):
    headers = {"Authorization": f"Bearer {data.rt_token}", "Content-Type": "application/json"}

    # ВТОРОЙ ШАГ: Генерация картинок (если включено)
    final_slides = []
    for slide in data.slides:
        slide_dict = slide.dict()
        if data.use_images:
            img_payload = {
                "uuid": str(uuid.uuid4()),
                "image": {"request": f"{slide.image_prompt}, style: {data.style}", "model": "yandex-art",
                          "aspect": "16:9"}
            }
            img_resp = requests.post(f"{RT_API_BASE}/ya/image", json=img_payload, headers=headers)
            if img_resp.status_code == 200:
                msg_id = img_resp.json()[0]["message"]["id"]
                time.sleep(5)  # Ожидание
                dl_url = f"{RT_API_BASE}/download?id={msg_id}&serviceType=yaArt&imageType=png"
                img_data = requests.get(dl_url, headers=headers).content
                slide_dict["imageData"] = "image/png;base64," + base64.b64encode(img_data).decode()
        final_slides.append(slide_dict)

    # Собираем файл
    sid = str(uuid.uuid4())
    pptx_path = WORK_DIR / f"{sid}.pptx"
    json_path = WORK_DIR / f"{sid}.json"

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"presentation_title": data.presentation_title, "slides": final_slides}, f, ensure_ascii=False)

    subprocess.run(["node", "generate.js", str(json_path), str(pptx_path)], check=True)
    return {"download_url": f"/api/download/{sid}"}


@app.get("/api/download/{sid}")
async def download(sid: str):
    return FileResponse(WORK_DIR / f"{sid}.pptx", filename="presentation_pro.pptx")


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_TEMPLATE