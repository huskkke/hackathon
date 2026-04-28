#!/usr/bin/env python3
"""
PresentAI — AI Presentation Generator
Hackathon: Амурский Код 2026 | Кейс: Ростелеком
"""

import os, uuid, json, io, time, re, base64, subprocess, tempfile, traceback
from pathlib import Path
from typing import Optional
import asyncio

import uvicorn
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import anthropic

# ── Setup ─────────────────────────────────────────────────────────────────────

WORK_DIR = Path("/tmp/presentai")
WORK_DIR.mkdir(exist_ok=True)
SCRIPT_DIR = Path(__file__).parent

anthropic_client = anthropic.Anthropic()

# RT API (optional — used if token provided)
RT_API_BASE = "https://ai.rt.ru/api/1.0"

app = FastAPI(title="PresentAI")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ── Document extraction ───────────────────────────────────────────────────────

def extract_text_from_pdf(data: bytes) -> str:
    try:
        import PyPDF2
        reader = PyPDF2.PdfReader(io.BytesIO(data))
        texts = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                texts.append(t)
        return "\n\n".join(texts)[:8000]
    except Exception as e:
        return f"[PDF извлечение не удалось: {e}]"

def extract_text_from_docx(data: bytes) -> str:
    try:
        from docx import Document
        doc = Document(io.BytesIO(data))
        paras = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n\n".join(paras)[:8000]
    except Exception as e:
        return f"[DOCX извлечение не удалось: {e}]"

# ── LLM: generate slide structure ────────────────────────────────────────────

SYSTEM_PROMPT = """You are an expert presentation designer and content strategist.
You generate structured JSON for professional presentations. Always respond ONLY with valid JSON.
No markdown, no explanations, just the JSON object."""

def build_generation_prompt(user_prompt: str, doc_text: str, slide_count: int,
                             style: str, tone: str) -> str:
    tone_desc = {
        "professional": "профессиональный, деловой, чёткий",
        "creative": "творческий, вдохновляющий, образный",
        "academic": "академический, точный, аналитический",
        "casual": "дружелюбный, понятный, разговорный",
        "persuasive": "убедительный, мотивирующий, динамичный",
    }.get(tone, "профессиональный")

    doc_section = ""
    if doc_text.strip():
        doc_section = f"\n\nДОКУМЕНТ (источник контента):\n---\n{doc_text[:6000]}\n---"

    return f"""Создай структуру презентации на основе запроса.

ЗАПРОС ПОЛЬЗОВАТЕЛЯ: {user_prompt}
КОЛИЧЕСТВО СЛАЙДОВ: {slide_count}
СТИЛЬ: {style}
ТОН: {tone_desc}{doc_section}

Верни JSON в точно следующем формате (без markdown-блоков):
{{
  "presentation_title": "Название презентации",
  "metadata": {{
    "title": "Название",
    "author": "PresentAI",
    "contact": ""
  }},
  "slides": [
    {{
      "layout": "title",
      "title": "Главный заголовок",
      "subtitle": "Подзаголовок или описание",
      "image_prompt": null
    }},
    {{
      "layout": "content",
      "title": "Заголовок слайда",
      "bullets": ["Пункт 1", "Пункт 2", "Пункт 3"],
      "image_prompt": "Описание изображения для этого слайда на английском для image AI"
    }},
    {{
      "layout": "two_column",
      "title": "Сравнение",
      "leftTitle": "Левый столбец",
      "leftContent": ["Пункт 1", "Пункт 2"],
      "rightTitle": "Правый столбец",
      "rightContent": ["Пункт A", "Пункт B"],
      "image_prompt": null
    }},
    {{
      "layout": "stats",
      "title": "Ключевые показатели",
      "stats": [
        {{"value": "95%", "label": "Метрика 1"}},
        {{"value": "2x", "label": "Метрика 2"}},
        {{"value": "10K+", "label": "Метрика 3"}}
      ],
      "content": "Дополнительное описание",
      "image_prompt": null
    }},
    {{
      "layout": "section_break",
      "title": "Название раздела",
      "subtitle": "Описание раздела",
      "image_prompt": null
    }},
    {{
      "layout": "quote",
      "quote": "Вдохновляющая цитата или ключевая мысль",
      "title": "— Источник или автор",
      "image_prompt": null
    }},
    {{
      "layout": "conclusion",
      "title": "Спасибо!",
      "content": "Итоговая мысль или призыв к действию",
      "image_prompt": null
    }}
  ]
}}

ПРАВИЛА:
1. Ровно {slide_count} слайдов в массиве slides
2. Первый слайд ВСЕГДА layout="title"
3. Последний слайд ВСЕГДА layout="conclusion"
4. Используй разные layouts: content, two_column, stats, section_break, quote
5. image_prompt — описание на английском языке (для AI-генерации), или null если изображение не нужно
6. Контент на русском языке, живой и информативный
7. Пункты в bullets должны быть конкретными и краткими (до 15 слов)
8. Только валидный JSON без лишних символов"""

def generate_slide_structure(user_prompt: str, doc_text: str,
                               slide_count: int, style: str, tone: str) -> dict:
    prompt = build_generation_prompt(user_prompt, doc_text, slide_count, style, tone)

    message = anthropic_client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}]
    )

    text = message.content[0].text.strip()

    # Clean JSON if wrapped in markdown
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    text = text.strip()

    return json.loads(text)

# ── Image generation (RT API) ─────────────────────────────────────────────────

def generate_image_rt(prompt: str, token: str, service: str = "sd") -> Optional[bytes]:
    """Generate image using RT API (Stable Diffusion or Yandex ART)"""
    try:
        import requests as req
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        request_id = str(uuid.uuid4())

        if service == "yaArt":
            payload = {
                "uuid": request_id,
                "image": {
                    "request": prompt,
                    "seed": int(time.time()) % 999999999,
                    "translate": True,
                    "model": "yandex-art",
                    "aspect": "16:9",
                },
            }
            resp = req.post(f"{RT_API_BASE}/ya/image", json=payload, headers=headers, timeout=30)
        else:  # SD
            payload = {
                "uuid": request_id,
                "sdImage": {
                    "request": prompt,
                    "seed": int(time.time()) % 999999999,
                    "translate": True,
                },
            }
            resp = req.post(f"{RT_API_BASE}/sd/img", json=payload, headers=headers, timeout=30)

        resp.raise_for_status()
        data = resp.json()
        if not data or not isinstance(data, list):
            return None

        msg_id = data[0].get("message", {}).get("id")
        if not msg_id:
            return None

        # Download the image
        svc_type = "yaArt" if service == "yaArt" else "sd"
        dl_url = f"{RT_API_BASE}/download?id={msg_id}&serviceType={svc_type}&imageType=png"
        dl_resp = req.get(dl_url, headers=headers, timeout=30)
        dl_resp.raise_for_status()
        return dl_resp.content

    except Exception as e:
        print(f"Image generation failed: {e}")
        return None

def image_bytes_to_base64(data: bytes) -> str:
    return "image/png;base64," + base64.b64encode(data).decode()

# ── PPTX generation ──────────────────────────────────────────────────────────

def build_pptx(slide_data: dict, output_path: str, theme: str) -> bool:
    """Call Node.js pptxgenjs script to build the PPTX"""
    data_path = str(WORK_DIR / f"{uuid.uuid4()}_data.json")
    try:
        slide_data["theme"] = theme
        with open(data_path, "w", encoding="utf-8") as f:
            json.dump(slide_data, f, ensure_ascii=False)

        script = str(SCRIPT_DIR / "generate.js")
        result = subprocess.run(
            ["node", script, data_path, output_path],
            capture_output=True, text=True, timeout=60,
            cwd=str(SCRIPT_DIR),
        )
        if result.returncode != 0:
            print("Node error:", result.stderr)
            return False
        return True
    finally:
        if os.path.exists(data_path):
            os.remove(data_path)

# ── API Routes ────────────────────────────────────────────────────────────────

@app.post("/api/generate")
async def generate_presentation(
    prompt: str = Form(...),
    slide_count: int = Form(8),
    style: str = Form("modern"),
    tone: str = Form("professional"),
    generate_images: bool = Form(False),
    rt_token: str = Form(""),
    rt_service: str = Form("sd"),
    document: Optional[UploadFile] = File(None),
):
    session_id = str(uuid.uuid4())
    output_path = str(WORK_DIR / f"{session_id}.pptx")

    try:
        # 1. Extract document text
        doc_text = ""
        if document and document.filename:
            content = await document.read()
            fname = document.filename.lower()
            if fname.endswith(".pdf"):
                doc_text = extract_text_from_pdf(content)
            elif fname.endswith(".docx"):
                doc_text = extract_text_from_docx(content)

        # 2. Generate slide structure via LLM
        slide_count = max(3, min(20, slide_count))
        structure = generate_slide_structure(prompt, doc_text, slide_count, style, tone)

        # 3. Generate images if requested and token provided
        slides = structure.get("slides", [])
        if generate_images and rt_token.strip():
            for slide in slides:
                if slide.get("image_prompt"):
                    img_bytes = generate_image_rt(
                        slide["image_prompt"], rt_token.strip(), rt_service
                    )
                    if img_bytes:
                        slide["imageData"] = image_bytes_to_base64(img_bytes)
                    time.sleep(0.5)  # Rate limit
        elif generate_images:
            # No token — attach placeholder flag
            for slide in slides:
                slide.pop("imageData", None)

        structure["slides"] = slides

        # 4. Build PPTX
        ok = build_pptx(structure, output_path, style)
        if not ok:
            raise HTTPException(500, "Ошибка генерации PPTX")

        # 5. Return preview data + session_id
        preview = []
        for i, s in enumerate(slides):
            preview.append({
                "index": i + 1,
                "layout": s.get("layout", "content"),
                "title": s.get("title", ""),
                "subtitle": s.get("subtitle", s.get("quote", "")),
                "bullets": s.get("bullets", []),
                "leftTitle": s.get("leftTitle", ""),
                "rightTitle": s.get("rightTitle", ""),
                "stats": s.get("stats", []),
                "hasImage": bool(s.get("imageData")),
            })

        return JSONResponse({
            "session_id": session_id,
            "title": structure.get("presentation_title", "Презентация"),
            "slide_count": len(slides),
            "preview": preview,
        })

    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Ошибка парсинга ответа LLM: {e}")
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.get("/api/download/{session_id}")
async def download_pptx(session_id: str):
    path = WORK_DIR / f"{session_id}.pptx"
    if not path.exists():
        raise HTTPException(404, "Файл не найден")
    return FileResponse(
        str(path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="presentation.pptx",
    )


@app.get("/", response_class=HTMLResponse)
async def frontend():
    return HTML_TEMPLATE

# ── Frontend ──────────────────────────────────────────────────────────────────

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PresentAI — Генератор презентаций</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,600;1,9..40,400&display=swap" rel="stylesheet">
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg: #07040f;
  --surface: #0e0820;
  --panel: #160d2e;
  --border: rgba(139,92,246,0.15);
  --border-bright: rgba(139,92,246,0.4);
  --accent-a: #8b5cf6;
  --accent-b: #ec4e99;
  --accent-c: #22d3ee;
  --text: #f0eaff;
  --muted: #8b7aaa;
  --radius: 16px;
  --font-head: 'Syne', sans-serif;
  --font-body: 'DM Sans', sans-serif;
}

html { scroll-behavior: smooth; }

body {
  background: var(--bg);
  color: var(--text);
  font-family: var(--font-body);
  min-height: 100vh;
  overflow-x: hidden;
}

/* ── Animated BG ─────────────────────────────────────── */
.bg-glow {
  position: fixed; inset: 0; pointer-events: none; z-index: 0;
  overflow: hidden;
}
.glow-orb {
  position: absolute; border-radius: 50%;
  filter: blur(120px); opacity: 0.25;
  animation: orbFloat 12s ease-in-out infinite;
}
.glow-orb.a { width: 600px; height: 600px; background: var(--accent-a); top: -200px; left: -200px; animation-delay: 0s; }
.glow-orb.b { width: 500px; height: 500px; background: var(--accent-b); top: 40%; right: -150px; animation-delay: -4s; }
.glow-orb.c { width: 400px; height: 400px; background: var(--accent-c); bottom: -100px; left: 30%; animation-delay: -8s; }

@keyframes orbFloat {
  0%, 100% { transform: translate(0, 0) scale(1); }
  33% { transform: translate(30px, -40px) scale(1.05); }
  66% { transform: translate(-20px, 30px) scale(0.95); }
}

/* ── Layout ──────────────────────────────────────────── */
.container {
  position: relative; z-index: 1;
  max-width: 960px; margin: 0 auto; padding: 0 24px;
}

/* ── Header ──────────────────────────────────────────── */
header {
  padding: 28px 0 0;
  display: flex; align-items: center; gap: 14px;
  animation: fadeDown 0.6s ease both;
}
.logo-mark {
  width: 48px; height: 48px; border-radius: 14px;
  background: linear-gradient(135deg, var(--accent-a), var(--accent-b));
  display: flex; align-items: center; justify-content: center;
  font-size: 22px;
  box-shadow: 0 0 24px rgba(139,92,246,0.4);
}
.logo-text { font-family: var(--font-head); font-weight: 800; font-size: 1.5rem; }
.logo-text span { color: var(--accent-a); }
.badge {
  margin-left: auto;
  font-size: 11px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase;
  padding: 4px 10px; border-radius: 100px;
  border: 1px solid var(--border-bright);
  color: var(--accent-a);
}

/* ── Hero ────────────────────────────────────────────── */
.hero {
  padding: 60px 0 40px;
  animation: fadeUp 0.7s 0.1s ease both;
}
.hero h1 {
  font-family: var(--font-head); font-size: clamp(2rem, 5vw, 3.4rem); font-weight: 800;
  line-height: 1.1; letter-spacing: -0.02em;
  background: linear-gradient(135deg, #fff 30%, var(--accent-a), var(--accent-b));
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text;
  margin-bottom: 16px;
}
.hero p {
  font-size: 1.1rem; color: var(--muted); max-width: 580px; line-height: 1.65;
}

/* ── Steps indicator ────────────────────────────────── */
.steps {
  display: flex; gap: 0; margin-bottom: 36px;
  animation: fadeUp 0.7s 0.15s ease both;
}
.step-item {
  display: flex; align-items: center; gap: 8px;
  font-size: 13px; font-weight: 600; color: var(--muted);
  transition: color 0.3s;
  cursor: default;
}
.step-item.active { color: var(--accent-a); }
.step-item.done { color: var(--muted); }
.step-num {
  width: 28px; height: 28px; border-radius: 50%;
  border: 2px solid currentColor;
  display: flex; align-items: center; justify-content: center;
  font-size: 12px; font-weight: 700;
  transition: all 0.3s;
}
.step-item.active .step-num {
  background: var(--accent-a); border-color: var(--accent-a); color: white;
  box-shadow: 0 0 14px rgba(139,92,246,0.5);
}
.step-item.done .step-num {
  background: rgba(139,92,246,0.2); border-color: rgba(139,92,246,0.4); color: var(--accent-a);
}
.step-connector {
  flex: 1; height: 2px; background: var(--border); margin: 0 8px;
  min-width: 20px; max-width: 60px;
}
.step-connector.active { background: var(--accent-a); }

/* ── Cards ───────────────────────────────────────────── */
.card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 28px;
  animation: fadeUp 0.7s 0.2s ease both;
  transition: border-color 0.3s;
}
.card:hover { border-color: var(--border-bright); }
.card + .card { margin-top: 20px; }

.card-title {
  font-family: var(--font-head); font-weight: 700; font-size: 1rem;
  color: var(--text); margin-bottom: 18px;
  display: flex; align-items: center; gap: 8px;
}
.card-title .icon {
  width: 30px; height: 30px; border-radius: 8px;
  background: rgba(139,92,246,0.15);
  display: flex; align-items: center; justify-content: center;
  font-size: 14px;
}

/* ── Form elements ──────────────────────────────────── */
label { display: block; font-size: 13px; font-weight: 600; color: var(--muted); margin-bottom: 6px; }

textarea, input[type="text"], input[type="number"], select {
  width: 100%; padding: 12px 14px;
  background: var(--panel); border: 1px solid var(--border);
  border-radius: 10px; color: var(--text);
  font-family: var(--font-body); font-size: 15px;
  transition: border-color 0.2s, box-shadow 0.2s;
  outline: none; resize: vertical;
  -webkit-appearance: none;
}
textarea:focus, input:focus, select:focus {
  border-color: var(--accent-a);
  box-shadow: 0 0 0 3px rgba(139,92,246,0.15);
}
textarea { min-height: 100px; }
select { cursor: pointer; }
select option { background: var(--panel); }

.field { margin-bottom: 16px; }

/* Row of fields */
.fields-row { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
@media (max-width: 600px) { .fields-row { grid-template-columns: 1fr; } }

/* ── Upload zone ─────────────────────────────────────── */
.upload-zone {
  border: 2px dashed var(--border-bright);
  border-radius: 12px; padding: 28px 20px;
  text-align: center; cursor: pointer;
  transition: all 0.2s; position: relative;
}
.upload-zone:hover, .upload-zone.dragover {
  border-color: var(--accent-a);
  background: rgba(139,92,246,0.06);
}
.upload-zone input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
.upload-icon { font-size: 28px; margin-bottom: 8px; }
.upload-text { font-size: 14px; color: var(--muted); }
.upload-text strong { color: var(--accent-a); }
.file-name { margin-top: 10px; font-size: 13px; color: var(--accent-c); font-weight: 600; }

/* ── Toggle options ──────────────────────────────────── */
.option-grid {
  display: grid; gap: 8px;
  grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
}
.option-item input { display: none; }
.option-item label {
  display: flex; flex-direction: column; align-items: center;
  gap: 6px; padding: 12px 8px; border-radius: 10px;
  border: 1px solid var(--border); cursor: pointer;
  transition: all 0.2s; color: var(--muted);
  font-size: 12px; font-weight: 600;
  text-align: center; margin: 0;
}
.option-item label .emoji { font-size: 20px; }
.option-item input:checked + label {
  border-color: var(--accent-a); background: rgba(139,92,246,0.12);
  color: var(--accent-a);
  box-shadow: 0 0 12px rgba(139,92,246,0.2);
}

/* ── Slider ──────────────────────────────────────────── */
.slider-wrap { display: flex; align-items: center; gap: 12px; }
input[type="range"] {
  flex: 1; -webkit-appearance: none; background: transparent;
  height: 6px; outline: none;
}
input[type="range"]::-webkit-slider-runnable-track {
  background: var(--border-bright); height: 4px; border-radius: 4px;
}
input[type="range"]::-webkit-slider-thumb {
  -webkit-appearance: none;
  width: 18px; height: 18px; border-radius: 50%;
  background: var(--accent-a); margin-top: -7px; cursor: pointer;
  box-shadow: 0 0 10px rgba(139,92,246,0.5);
}
.slider-val {
  width: 36px; text-align: center; font-weight: 700;
  font-size: 1.1rem; color: var(--accent-a);
  font-family: var(--font-head);
}

/* ── Toggle switch ──────────────────────────────────── */
.toggle-wrap { display: flex; align-items: center; gap: 10px; cursor: pointer; }
.toggle { position: relative; width: 44px; height: 24px; }
.toggle input { display: none; }
.toggle-track {
  position: absolute; inset: 0; border-radius: 12px;
  background: var(--panel); border: 1px solid var(--border);
  transition: all 0.3s;
}
.toggle-thumb {
  position: absolute; top: 3px; left: 3px;
  width: 16px; height: 16px; border-radius: 50%;
  background: var(--muted); transition: all 0.3s;
}
.toggle input:checked ~ .toggle-track {
  background: rgba(139,92,246,0.25); border-color: var(--accent-a);
}
.toggle input:checked ~ .toggle-thumb {
  transform: translateX(20px); background: var(--accent-a);
  box-shadow: 0 0 10px rgba(139,92,246,0.5);
}
.toggle-label { font-size: 14px; color: var(--text); font-weight: 500; }

/* ── Button ──────────────────────────────────────────── */
.btn-generate {
  width: 100%; margin-top: 24px; padding: 16px;
  background: linear-gradient(135deg, var(--accent-a), var(--accent-b));
  border: none; border-radius: 12px; cursor: pointer;
  font-family: var(--font-head); font-size: 1rem; font-weight: 700;
  color: white; letter-spacing: 0.02em;
  transition: all 0.2s;
  box-shadow: 0 4px 24px rgba(139,92,246,0.35);
  display: flex; align-items: center; justify-content: center; gap: 10px;
  animation: fadeUp 0.7s 0.3s ease both;
}
.btn-generate:hover { transform: translateY(-2px); box-shadow: 0 8px 32px rgba(139,92,246,0.5); }
.btn-generate:active { transform: translateY(0); }
.btn-generate:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }

/* ── Progress ────────────────────────────────────────── */
#progress-section { display: none; animation: fadeUp 0.5s ease both; margin-top: 24px; }
.progress-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius); padding: 28px;
  text-align: center;
}
.progress-emoji { font-size: 40px; margin-bottom: 12px; animation: pulse 1.5s ease-in-out infinite; }
@keyframes pulse { 0%,100%{transform:scale(1)} 50%{transform:scale(1.1)} }
.progress-title { font-family: var(--font-head); font-size: 1.2rem; font-weight: 700; margin-bottom: 6px; }
.progress-sub { font-size: 13px; color: var(--muted); margin-bottom: 20px; }
.progress-bar-wrap {
  background: var(--panel); border-radius: 100px; height: 8px; overflow: hidden;
}
.progress-bar {
  height: 100%; border-radius: 100px;
  background: linear-gradient(90deg, var(--accent-a), var(--accent-b));
  width: 0%; transition: width 0.5s ease;
  box-shadow: 0 0 12px rgba(139,92,246,0.5);
}
.progress-steps { margin-top: 16px; display: flex; flex-direction: column; gap: 6px; }
.progress-step { font-size: 13px; color: var(--muted); transition: color 0.3s; display: flex; align-items: center; gap: 8px; }
.progress-step.done { color: var(--text); }
.progress-step .dot { width: 6px; height: 6px; border-radius: 50%; background: var(--muted); flex-shrink: 0; transition: background 0.3s; }
.progress-step.done .dot { background: var(--accent-a); box-shadow: 0 0 6px var(--accent-a); }

/* ── Results ─────────────────────────────────────────── */
#result-section { display: none; margin-top: 24px; animation: fadeUp 0.5s ease both; }

.result-header {
  display: flex; align-items: center; gap: 14px;
  margin-bottom: 20px;
}
.result-title { font-family: var(--font-head); font-weight: 800; font-size: 1.3rem; }
.result-meta { font-size: 13px; color: var(--muted); }
.btn-download {
  margin-left: auto;
  padding: 12px 22px;
  background: linear-gradient(135deg, var(--accent-a), var(--accent-b));
  border: none; border-radius: 10px; cursor: pointer;
  font-family: var(--font-head); font-size: 0.9rem; font-weight: 700;
  color: white; transition: all 0.2s;
  box-shadow: 0 4px 16px rgba(139,92,246,0.3);
  display: flex; align-items: center; gap: 8px; text-decoration: none;
  white-space: nowrap;
}
.btn-download:hover { transform: translateY(-2px); box-shadow: 0 6px 24px rgba(139,92,246,0.45); }

/* ── Slide preview grid ─────────────────────────────── */
.slides-grid {
  display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 14px;
}

.slide-card {
  background: var(--panel); border: 1px solid var(--border);
  border-radius: 12px; overflow: hidden; cursor: pointer;
  transition: all 0.2s;
}
.slide-card:hover { border-color: var(--border-bright); transform: translateY(-3px); box-shadow: 0 8px 24px rgba(0,0,0,0.3); }

.slide-thumb {
  position: relative;
  background: #0D1117;
  padding: 14px; min-height: 120px;
  display: flex; flex-direction: column; gap: 6px;
}
.slide-thumb.layout-title { background: linear-gradient(135deg, #120238, #2D1065); }
.slide-thumb.layout-conclusion { background: linear-gradient(135deg, #0D1117, #1a0533); }
.slide-thumb.layout-section_break { background: linear-gradient(135deg, #0e1a0a, #1a3510); }
.slide-thumb.layout-quote { background: linear-gradient(135deg, #0a1222, #1a2540); }
.slide-thumb.layout-stats { background: #0D1117; }
.slide-thumb.layout-two_column { background: #0D1117; }

.slide-num {
  position: absolute; top: 8px; right: 10px;
  font-size: 10px; font-weight: 700; color: rgba(255,255,255,0.3);
  font-family: var(--font-head);
}

.slide-layout-badge {
  font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em;
  color: var(--accent-a); opacity: 0.8; margin-bottom: 2px;
}

.slide-thumb-title {
  font-size: 12px; font-weight: 700; color: white; line-height: 1.3;
  font-family: var(--font-head);
  display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden;
}
.slide-thumb-content {
  font-size: 9px; color: rgba(255,255,255,0.45); line-height: 1.4;
  display: -webkit-box; -webkit-line-clamp: 3; -webkit-box-orient: vertical; overflow: hidden;
}

.slide-meta {
  padding: 8px 14px; font-size: 10px; color: var(--muted);
  display: flex; align-items: center; justify-content: space-between;
  border-top: 1px solid var(--border);
}

/* mini bullets in preview */
.mini-bullets { display: flex; flex-direction: column; gap: 3px; margin-top: 4px; }
.mini-bullet { font-size: 8px; color: rgba(255,255,255,0.45); display: flex; gap: 4px; align-items: flex-start; }
.mini-bullet::before { content: '▸'; color: var(--accent-a); flex-shrink: 0; }

/* stats preview */
.mini-stats { display: flex; gap: 6px; margin-top: 4px; flex-wrap: wrap; }
.mini-stat { background: rgba(139,92,246,0.15); border-radius: 4px; padding: 2px 6px; }
.mini-stat-val { font-size: 11px; font-weight: 700; color: var(--accent-a); }
.mini-stat-lab { font-size: 7px; color: var(--muted); }

/* ── Error ───────────────────────────────────────────── */
.error-card {
  background: rgba(239,68,68,0.08); border: 1px solid rgba(239,68,68,0.3);
  border-radius: 12px; padding: 20px; margin-top: 20px;
  display: none; color: #fca5a5; font-size: 14px;
}

/* ── RT API section ──────────────────────────────────── */
.rt-section {
  border-top: 1px solid var(--border); margin-top: 16px; padding-top: 16px;
  transition: all 0.3s;
}
.rt-section.hidden { display: none; }

/* ── Animations ──────────────────────────────────────── */
@keyframes fadeUp {
  from { opacity: 0; transform: translateY(20px); }
  to { opacity: 1; transform: translateY(0); }
}
@keyframes fadeDown {
  from { opacity: 0; transform: translateY(-10px); }
  to { opacity: 1; transform: translateY(0); }
}

/* ── Footer ──────────────────────────────────────────── */
footer {
  text-align: center; padding: 48px 24px 32px;
  font-size: 12px; color: var(--muted);
}
footer strong { color: var(--accent-a); }
</style>
</head>
<body>

<div class="bg-glow">
  <div class="glow-orb a"></div>
  <div class="glow-orb b"></div>
  <div class="glow-orb c"></div>
</div>

<div class="container">
  <header>
    <div class="logo-mark">✦</div>
    <div class="logo-text">Present<span>AI</span></div>
    <div class="badge">Амурский Код 2026</div>
  </header>

  <div class="hero">
    <h1>От промпта до<br>готового PPTX</h1>
    <p>Введите запрос, загрузите документ — и получите профессиональную презентацию за секунды. Используем LLM и AI-генерацию изображений.</p>
  </div>

  <!-- Steps -->
  <div class="steps" id="steps">
    <div class="step-item active" id="step1-indicator">
      <div class="step-num">1</div>
      <span>Контент</span>
    </div>
    <div class="step-connector" id="conn1"></div>
    <div class="step-item" id="step2-indicator">
      <div class="step-num">2</div>
      <span>Настройки</span>
    </div>
    <div class="step-connector" id="conn2"></div>
    <div class="step-item" id="step3-indicator">
      <div class="step-num">3</div>
      <span>Генерация</span>
    </div>
    <div class="step-connector" id="conn3"></div>
    <div class="step-item" id="step4-indicator">
      <div class="step-num">4</div>
      <span>Результат</span>
    </div>
  </div>

  <!-- Step 1: Prompt + Document -->
  <div class="card" id="form-section">
    <div class="card-title"><div class="icon">💬</div> Запрос и документ</div>
    <div class="field">
      <label>Описание презентации *</label>
      <textarea id="prompt" placeholder="Например: Создай презентацию о преимуществах AI в телекоммуникациях для руководителей компании. Включи статистику, кейсы и план внедрения."></textarea>
    </div>
    <div class="field">
      <label>Исходный документ (необязательно)</label>
      <div class="upload-zone" id="upload-zone">
        <input type="file" id="doc-file" accept=".pdf,.docx" onchange="handleFileSelect(this)">
        <div class="upload-icon">📄</div>
        <div class="upload-text">
          <strong>Выберите файл</strong> или перетащите сюда<br>
          PDF или DOCX — контент будет учтён при генерации
        </div>
        <div class="file-name" id="file-name-display" style="display:none"></div>
      </div>
    </div>

    <!-- Step 2: Settings -->
    <div class="card-title" style="margin-top:24px"><div class="icon">⚙️</div> Параметры</div>

    <div class="field">
      <label>Количество слайдов</label>
      <div class="slider-wrap">
        <input type="range" id="slide-count" min="4" max="20" value="8"
               oninput="document.getElementById('slide-count-val').textContent=this.value">
        <div class="slider-val" id="slide-count-val">8</div>
      </div>
    </div>

    <div class="fields-row">
      <div class="field">
        <label>Тон</label>
        <select id="tone">
          <option value="professional">Профессиональный</option>
          <option value="creative">Творческий</option>
          <option value="academic">Академический</option>
          <option value="casual">Дружелюбный</option>
          <option value="persuasive">Убедительный</option>
        </select>
      </div>
      <div class="field">
        <label>Визуальный стиль</label>
        <select id="style">
          <option value="modern">Modern Dark</option>
          <option value="rostelecom" selected>Ростелеком</option>
          <option value="corporate">Corporate</option>
          <option value="minimal">Minimal</option>
          <option value="tech">Tech Green</option>
        </select>
      </div>
    </div>

    <div class="field">
      <label>Тема презентации (шаблон слайдов)</label>
      <div class="option-grid">
        <div class="option-item">
          <input type="radio" name="style-radio" id="s-rostelecom" value="rostelecom" checked onchange="document.getElementById('style').value=this.value">
          <label for="s-rostelecom"><span class="emoji">🚀</span>Ростелеком</label>
        </div>
        <div class="option-item">
          <input type="radio" name="style-radio" id="s-modern" value="modern" onchange="document.getElementById('style').value=this.value">
          <label for="s-modern"><span class="emoji">🌙</span>Modern Dark</label>
        </div>
        <div class="option-item">
          <input type="radio" name="style-radio" id="s-corporate" value="corporate" onchange="document.getElementById('style').value=this.value">
          <label for="s-corporate"><span class="emoji">🏢</span>Corporate</label>
        </div>
        <div class="option-item">
          <input type="radio" name="style-radio" id="s-minimal" value="minimal" onchange="document.getElementById('style').value=this.value">
          <label for="s-minimal"><span class="emoji">⬜</span>Minimal</label>
        </div>
        <div class="option-item">
          <input type="radio" name="style-radio" id="s-tech" value="tech" onchange="document.getElementById('style').value=this.value">
          <label for="s-tech"><span class="emoji">💻</span>Tech</label>
        </div>
      </div>
    </div>

    <!-- Image generation toggle -->
    <div class="rt-section">
      <div class="toggle-wrap" onclick="toggleImages()">
        <div class="toggle">
          <input type="checkbox" id="gen-images">
          <div class="toggle-track"></div>
          <div class="toggle-thumb"></div>
        </div>
        <span class="toggle-label">Генерировать изображения (RT API)</span>
      </div>
      <div id="rt-fields" class="hidden" style="margin-top:14px">
        <div class="fields-row">
          <div class="field">
            <label>API Token (Bearer)</label>
            <input type="text" id="rt-token" placeholder="Вставьте токен Ростелеком API">
          </div>
          <div class="field">
            <label>Модель изображений</label>
            <select id="rt-service">
              <option value="yaArt">Yandex ART</option>
              <option value="sd">Stable Diffusion</option>
            </select>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Generate button -->
  <button class="btn-generate" id="gen-btn" onclick="generate()">
    <span>✦</span> Сгенерировать презентацию
  </button>

  <!-- Progress -->
  <div id="progress-section">
    <div class="progress-card">
      <div class="progress-emoji" id="progress-emoji">🧠</div>
      <div class="progress-title" id="progress-title">Анализируем запрос...</div>
      <div class="progress-sub" id="progress-sub">LLM обрабатывает ваш запрос</div>
      <div class="progress-bar-wrap">
        <div class="progress-bar" id="progress-bar"></div>
      </div>
      <div class="progress-steps">
        <div class="progress-step" id="ps1"><div class="dot"></div>Чтение документа</div>
        <div class="progress-step" id="ps2"><div class="dot"></div>Генерация структуры слайдов</div>
        <div class="progress-step" id="ps3"><div class="dot"></div>Создание контента</div>
        <div class="progress-step" id="ps4"><div class="dot"></div>Сборка PPTX файла</div>
      </div>
    </div>
  </div>

  <!-- Error -->
  <div class="error-card" id="error-card"></div>

  <!-- Results -->
  <div id="result-section">
    <div class="result-header">
      <div>
        <div class="result-title" id="result-title">Готово! 🎉</div>
        <div class="result-meta" id="result-meta"></div>
      </div>
      <a class="btn-download" id="download-btn" href="#" download="presentation.pptx">
        ⬇ Скачать PPTX
      </a>
    </div>

    <div class="slides-grid" id="slides-grid"></div>

    <button class="btn-generate" style="margin-top:24px" onclick="resetForm()">
      ↩ Создать новую
    </button>
  </div>

</div>

<footer>
  Разработано для хакатона <strong>Амурский Код 2026</strong> · Кейс Ростелеком
</footer>

<script>
let sessionId = null;

function handleFileSelect(input) {
  const file = input.files[0];
  const display = document.getElementById('file-name-display');
  if (file) {
    display.textContent = '📎 ' + file.name;
    display.style.display = 'block';
  } else {
    display.style.display = 'none';
  }
}

// Drag-drop
const zone = document.getElementById('upload-zone');
zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
zone.addEventListener('drop', e => {
  e.preventDefault(); zone.classList.remove('dragover');
  const f = e.dataTransfer.files[0];
  if (f) {
    document.getElementById('doc-file').files = e.dataTransfer.files;
    handleFileSelect(document.getElementById('doc-file'));
  }
});

function toggleImages() {
  const cb = document.getElementById('gen-images');
  cb.checked = !cb.checked;
  document.getElementById('rt-fields').classList.toggle('hidden', !cb.checked);
}

function setStep(n) {
  for (let i = 1; i <= 4; i++) {
    const el = document.getElementById(`step${i}-indicator`);
    if (i < n) { el.className = 'step-item done'; }
    else if (i === n) { el.className = 'step-item active'; }
    else { el.className = 'step-item'; }
  }
  for (let i = 1; i <= 3; i++) {
    const c = document.getElementById(`conn${i}`);
    c.className = i < n ? 'step-connector active' : 'step-connector';
  }
}

function setProgress(pct, emoji, title, sub) {
  document.getElementById('progress-bar').style.width = pct + '%';
  if (emoji) document.getElementById('progress-emoji').textContent = emoji;
  if (title) document.getElementById('progress-title').textContent = title;
  if (sub) document.getElementById('progress-sub').textContent = sub;
}

function doneStep(id) {
  document.getElementById(id).classList.add('done');
}

async function generate() {
  const prompt = document.getElementById('prompt').value.trim();
  if (!prompt) { alert('Введите описание презентации'); return; }

  const genImages = document.getElementById('gen-images').checked;
  const rtToken = document.getElementById('rt-token').value.trim();
  const rtService = document.getElementById('rt-service').value;
  const slideCount = document.getElementById('slide-count').value;
  const style = document.getElementById('style').value;
  const tone = document.getElementById('tone').value;

  // UI transitions
  document.getElementById('gen-btn').disabled = true;
  document.getElementById('error-card').style.display = 'none';
  document.getElementById('result-section').style.display = 'none';
  document.getElementById('progress-section').style.display = 'block';
  setStep(3);
  setProgress(5, '📄', 'Подготовка...', 'Обрабатываем ваш запрос');

  // Progress animation
  setTimeout(() => { doneStep('ps1'); setProgress(20, '🧠', 'LLM генерирует слайды...', 'Создаём структуру презентации'); }, 800);
  setTimeout(() => { doneStep('ps2'); setProgress(50, '✍️', 'Формируем контент...', 'Подбираем тексты и заголовки'); }, 2500);
  setTimeout(() => { doneStep('ps3'); setProgress(75, '🎨', 'Применяем дизайн...', 'Компонуем PPTX файл'); }, 5000);

  const formData = new FormData();
  formData.append('prompt', prompt);
  formData.append('slide_count', slideCount);
  formData.append('style', style);
  formData.append('tone', tone);
  formData.append('generate_images', genImages);
  formData.append('rt_token', rtToken);
  formData.append('rt_service', rtService);

  const fileInput = document.getElementById('doc-file');
  if (fileInput.files[0]) {
    formData.append('document', fileInput.files[0]);
  }

  try {
    const resp = await fetch('/api/generate', { method: 'POST', body: formData });
    const data = await resp.json();

    if (!resp.ok) throw new Error(data.detail || 'Ошибка генерации');

    doneStep('ps4');
    setProgress(100, '🎉', 'Готово!', 'Ваша презентация создана');

    setTimeout(() => showResults(data), 600);

  } catch (err) {
    document.getElementById('progress-section').style.display = 'none';
    const errCard = document.getElementById('error-card');
    errCard.style.display = 'block';
    errCard.innerHTML = `<strong>❌ Ошибка:</strong> ${err.message}`;
    document.getElementById('gen-btn').disabled = false;
    setStep(1);
  }
}

function showResults(data) {
  sessionId = data.session_id;
  setStep(4);

  document.getElementById('progress-section').style.display = 'none';
  document.getElementById('result-section').style.display = 'block';
  document.getElementById('result-title').textContent = data.title || 'Презентация готова!';
  document.getElementById('result-meta').textContent = `${data.slide_count} слайдов`;
  document.getElementById('download-btn').href = `/api/download/${data.session_id}`;

  const grid = document.getElementById('slides-grid');
  grid.innerHTML = '';

  (data.preview || []).forEach(s => {
    grid.appendChild(buildSlideCard(s));
  });

  document.getElementById('gen-btn').disabled = false;
}

const LAYOUT_LABELS = {
  title: 'Титульный', content: 'Контент', two_column: 'Два столбца',
  stats: 'Статистика', quote: 'Цитата', section_break: 'Раздел',
  image: 'Изображение', conclusion: 'Заключение', section: 'Раздел',
};

function buildSlideCard(s) {
  const card = document.createElement('div');
  card.className = 'slide-card';

  const layout = (s.layout || 'content').replace('-', '_');
  const thumbCls = `slide-thumb layout-${layout}`;

  let contentHtml = '';

  if (s.stats && s.stats.length > 0) {
    contentHtml = `<div class="mini-stats">${s.stats.slice(0,3).map(st =>
      `<div class="mini-stat"><div class="mini-stat-val">${st.value}</div><div class="mini-stat-lab">${st.label}</div></div>`
    ).join('')}</div>`;
  } else if (s.bullets && s.bullets.length > 0) {
    contentHtml = `<div class="mini-bullets">${s.bullets.slice(0,3).map(b =>
      `<div class="mini-bullet">${b.substring(0,50)}${b.length>50?'…':''}</div>`
    ).join('')}</div>`;
  } else if (s.subtitle) {
    contentHtml = `<div class="slide-thumb-content">${s.subtitle}</div>`;
  } else if (s.leftTitle || s.rightTitle) {
    contentHtml = `<div style="display:flex;gap:6px;margin-top:4px">
      <div style="flex:1;background:rgba(139,92,246,0.15);border-radius:4px;padding:4px 6px;font-size:9px;color:rgba(255,255,255,0.5)">${s.leftTitle||''}</div>
      <div style="flex:1;background:rgba(236,78,153,0.15);border-radius:4px;padding:4px 6px;font-size:9px;color:rgba(255,255,255,0.5)">${s.rightTitle||''}</div>
    </div>`;
  }

  card.innerHTML = `
    <div class="${thumbCls}">
      <div class="slide-num">${s.index}</div>
      <div class="slide-layout-badge">${LAYOUT_LABELS[layout] || layout}</div>
      <div class="slide-thumb-title">${s.title || ''}</div>
      ${contentHtml}
      ${s.hasImage ? '<div style="margin-top:6px;font-size:9px;color:rgba(139,92,246,0.7)">🖼 + изображение</div>' : ''}
    </div>
    <div class="slide-meta">
      <span>${LAYOUT_LABELS[layout] || layout}</span>
      <span style="color:var(--accent-a)">→</span>
    </div>`;
  return card;
}

function resetForm() {
  document.getElementById('result-section').style.display = 'none';
  document.getElementById('progress-section').style.display = 'none';
  document.getElementById('error-card').style.display = 'none';
  document.getElementById('gen-btn').disabled = false;
  setStep(1);
  // Reset progress steps
  ['ps1','ps2','ps3','ps4'].forEach(id => document.getElementById(id).classList.remove('done'));
  document.getElementById('progress-bar').style.width = '0%';
}

// Sync radio buttons with select
document.getElementById('style').addEventListener('change', function() {
  const val = this.value;
  const radio = document.querySelector(`input[name="style-radio"][value="${val}"]`);
  if (radio) radio.checked = true;
});
</script>
</body>
</html>
"""

# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")