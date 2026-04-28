#!/usr/bin/env python3
"""
PresentAI — AI Presentation Generator
Hackathon: Амурский Код 2026 | Кейс: Ростелеком
"""

import os, uuid, json, io, time, re, base64, traceback, zipfile
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape
import asyncio

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

# ── Setup ─────────────────────────────────────────────────────────────────────

WORK_DIR = Path("./temp_files")
WORK_DIR.mkdir(exist_ok=True)
SCRIPT_DIR = Path(__file__).parent

# RT API (optional — used if token provided)
RT_API_BASE = "https://ai.rt.ru/api/1.0"
RT_LLM_MODEL = "Qwen/Qwen2.5-72B-Instruct"

app = FastAPI(title="PresentAI")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

JOBS: dict[str, dict] = {}

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

    return f"""Верни только валидный JSON без markdown.
Создай презентацию на русском языке.
Тема: {user_prompt}
Количество слайдов: {slide_count}
Стиль: {style}
Тон: {tone_desc}{doc_section}

Формат:
{{
  "presentation_title": "...",
  "slides": [
    {{"layout":"title","title":"...","subtitle":"..."}},
    {{"layout":"content","title":"...","bullets":["...","...","..."],"image_prompt":"english image prompt"}},
    {{"layout":"two_column","title":"...","leftTitle":"...","leftContent":["...","..."],"rightTitle":"...","rightContent":["...","..."],"image_prompt":"english image prompt"}},
    {{"layout":"stats","title":"...","stats":[{{"value":"...","label":"..."}}],"content":"...","image_prompt":"english image prompt"}},
    {{"layout":"conclusion","title":"...","content":"..."}}
  ]
}}

Правила: ровно {slide_count} слайдов; первый layout title; последний layout conclusion; текст слайдов строго по теме; bullets короткие; image_prompt на английском."""

def _extract_json_object(text: str) -> dict:
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start:end + 1])
        raise

def _find_text_in_response(data: object) -> str:
    if isinstance(data, str):
        return data
    if isinstance(data, list):
        parts = [_find_text_in_response(item) for item in data]
        return "\n".join(part for part in parts if part)
    if isinstance(data, dict):
        message = data.get("message")
        if isinstance(message, dict):
            content = message.get("content")
            if isinstance(content, str) and content.strip():
                return content
        for key in ("content", "text", "answer", "response", "message", "generated_text"):
            value = data.get(key)
            if isinstance(value, str) and value.strip():
                return value
        for value in data.values():
            found = _find_text_in_response(value)
            if found:
                return found
    return ""

def _call_rt_llm(prompt: str, token: str) -> str:
    import requests as req

    payload = {
        "uuid": str(uuid.uuid4()),
        "chat": {
            "model": RT_LLM_MODEL,
            "contents": [{"type": "text", "text": prompt}],
            "system_prompt": SYSTEM_PROMPT,
            "max_new_tokens": 1800,
            "temperature": 0.25,
            "top_p": 0.9,
            "repetition_penalty": 1.05,
        },
    }
    try:
        resp = req.post(
            f"{RT_API_BASE}/llama/chat",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
            timeout=60,
        )
        resp.raise_for_status()
    except req.HTTPError as e:
        body = e.response.text[:500] if e.response is not None else ""
        raise RuntimeError(f"HTTP {e.response.status_code if e.response is not None else ''}: {body}") from e

    raw = resp.text.strip()
    if not raw:
        raise RuntimeError(f"RT API вернул пустой ответ, HTTP {resp.status_code}")
    try:
        data = resp.json()
    except ValueError:
        return raw

    text = _find_text_in_response(data).strip()
    if not text:
        raise RuntimeError(f"RT API ответил без текста. Фрагмент ответа: {raw[:500]}")
    return text

def _call_anthropic(prompt: str) -> str:
    import anthropic

    client = anthropic.Anthropic()
    message = client.messages.create(
        model=os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-20250514"),
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text

def _fallback_slide_structure(user_prompt: str, doc_text: str, slide_count: int,
                              style: str, tone: str) -> dict:
    source = re.sub(r"\s+", " ", doc_text).strip()
    topic = user_prompt.strip(" .") or "выбранная тема"
    facts = [item.strip(" .") for item in re.split(r"[.\n;]+", source) if len(item.strip()) > 30]
    if not facts:
        facts = [
            f"Тема требует краткого обзора: что происходит, почему это важно и кого затрагивает",
            f"Для раскрытия темы «{topic}» важно показать причины, последствия и текущие вызовы",
            f"Отдельный блок стоит посвятить статистике, географии и динамике изменений",
            f"Практическая часть должна включать меры профилактики, реагирования и снижения рисков",
        ]

    title = user_prompt[:90].strip(" .") or "Презентация"
    slides = [{
        "layout": "title",
        "title": title,
        "subtitle": f"Краткий аналитический обзор по теме",
        "image_prompt": None,
    }]

    middle_count = max(1, slide_count - 2)
    for i in range(middle_count):
        fact = facts[i % len(facts)]
        if i % 4 == 1:
            slides.append({
                "layout": "two_column",
                "title": "Причины и последствия",
                "leftTitle": "Ключевые причины",
                "leftContent": ["Природные и климатические факторы", "Человеческий фактор", "Недостаточная профилактика"],
                "rightTitle": "Последствия",
                "rightContent": ["Риски для людей и инфраструктуры", "Экономический ущерб", "Экологические потери"],
                "image_prompt": None,
            })
        elif i % 4 == 2:
            slides.append({
                "layout": "stats",
                "title": "Что важно показать цифрами",
                "stats": [
                    {"value": "1", "label": "Масштаб проблемы"},
                    {"value": "2", "label": "Затронутые регионы"},
                    {"value": "3", "label": "Динамика и прогноз"},
                ],
                "content": fact,
                "image_prompt": None,
            })
        elif i % 4 == 3:
            slides.append({
                "layout": "quote",
                "title": "Ключевая идея",
                "quote": fact,
                "image_prompt": None,
            })
        else:
            slides.append({
                "layout": "content",
                "title": f"Слайд {i + 2}",
                "bullets": [
                    fact[:110],
                    f"Свяжите тезис с темой «{topic}» через конкретные примеры",
                    "Добавьте вывод: что нужно сделать дальше",
                ],
                "image_prompt": f"realistic editorial illustration about {topic}, 16:9 presentation slide",
            })

    slides.append({
        "layout": "conclusion",
        "title": "Выводы",
        "content": f"Тема «{topic}» требует системного анализа, профилактики и понятного плана действий.",
        "image_prompt": None,
    })
    return {"presentation_title": title, "metadata": {"title": title, "author": "PresentAI", "contact": ""}, "slides": slides[:slide_count]}

def _structure_from_plain_text(text: str, user_prompt: str, slide_count: int, style: str, tone: str) -> dict:
    chunks = [part.strip(" -\n\t") for part in re.split(r"\n+|(?<=[.!?])\s+", text) if len(part.strip()) > 20]
    if not chunks:
        return _fallback_slide_structure(user_prompt, text, slide_count, style, tone)
    structure = _fallback_slide_structure(user_prompt, text, slide_count, style, tone)
    slides = structure["slides"]
    for idx, chunk in enumerate(chunks[:max(1, slide_count - 2)], start=1):
        if idx >= len(slides) - 1:
            break
        slides[idx] = {
            "layout": "content",
            "title": f"Ключевой тезис {idx}",
            "bullets": [chunk[:150]],
            "image_prompt": f"editorial presentation image about {user_prompt}, 16:9",
        }
    structure["slides"] = slides
    return structure

def _normalize_slide_structure(data: dict, slide_count: int) -> dict:
    if not isinstance(data, dict):
        raise ValueError("LLM вернула не JSON-объект")
    slides = data.get("slides")
    if not isinstance(slides, list):
        raise ValueError("В JSON нет массива slides")
    if len(slides) > slide_count:
        conclusion = next((s for s in reversed(slides) if s.get("layout") == "conclusion"), slides[-1])
        slides = slides[:max(1, slide_count - 1)] + [conclusion]
    while len(slides) < slide_count:
        slides.append({"layout": "content", "title": f"Слайд {len(slides) + 1}", "bullets": [], "image_prompt": None})
    if slides:
        slides[0]["layout"] = "title"
        slides[-1]["layout"] = "conclusion"
    data["slides"] = slides
    data.setdefault("presentation_title", slides[0].get("title", "Презентация") if slides else "Презентация")
    data.setdefault("metadata", {"title": data["presentation_title"], "author": "PresentAI", "contact": ""})
    return data

def generate_slide_structure(user_prompt: str, doc_text: str,
                               slide_count: int, style: str, tone: str,
                               rt_token: str = "") -> dict:
    prompt = build_generation_prompt(user_prompt, doc_text, slide_count, style, tone)
    errors = []

    if rt_token.strip():
        try:
            rt_text = _call_rt_llm(prompt, rt_token.strip())
            try:
                structure = _normalize_slide_structure(_extract_json_object(rt_text), slide_count)
            except Exception:
                structure = _normalize_slide_structure(_structure_from_plain_text(rt_text, user_prompt, slide_count, style, tone), slide_count)
                structure["generation_warning"] = "RT LLM вернула текст не в JSON-формате, слайды собраны из текста ответа."
            structure["generation_source"] = "RT LLM"
            return structure
        except Exception as e:
            errors.append(f"RT LLM: {e}")

    if os.getenv("USE_ANTHROPIC_FALLBACK") == "1" and os.getenv("ANTHROPIC_API_KEY"):
        try:
            structure = _normalize_slide_structure(_extract_json_object(_call_anthropic(prompt)), slide_count)
            structure["generation_source"] = "Anthropic"
            return structure
        except Exception as e:
            errors.append(f"Anthropic: {e}")

    structure = _fallback_slide_structure(user_prompt, doc_text, slide_count, style, tone)
    structure["generation_source"] = "Локальный fallback"
    if errors:
        structure["generation_warning"] = "Использован локальный fallback. " + " | ".join(errors)
    else:
        structure["generation_warning"] = "Использован локальный fallback: API-токен LLM не задан."
    return _normalize_slide_structure(structure, slide_count)

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
            resp = req.post(f"{RT_API_BASE}/ya/image", json=payload, headers=headers, timeout=90)
        else:  # SD
            payload = {
                "uuid": request_id,
                "sdImage": {
                    "request": prompt,
                    "seed": int(time.time()) % 999999999,
                    "translate": True,
                },
            }
            resp = req.post(f"{RT_API_BASE}/sd/img", json=payload, headers=headers, timeout=90)

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
        dl_resp = req.get(dl_url, headers=headers, timeout=90)
        dl_resp.raise_for_status()
        return dl_resp.content

    except Exception as e:
        print(f"Image generation failed: {e}")
        return None

def image_bytes_to_base64(data: bytes) -> str:
    return "image/png;base64," + base64.b64encode(data).decode()

# ── PPTX generation ──────────────────────────────────────────────────────────

THEMES = {
    "modern": {"bg": "0F172A", "title": "F8FAFC", "text": "CBD5E1", "accent": "38BDF8"},
    "rostelecom": {"bg": "101828", "title": "FFFFFF", "text": "D0D5DD", "accent": "7700FF"},
    "corporate": {"bg": "FFFFFF", "title": "0F172A", "text": "334155", "accent": "2563EB"},
    "creative": {"bg": "2A0F2F", "title": "FFF7ED", "text": "F5D0FE", "accent": "F97316"},
    "minimal": {"bg": "FAFAFA", "title": "111827", "text": "374151", "accent": "111827"},
    "tech": {"bg": "07130D", "title": "ECFDF5", "text": "BBF7D0", "accent": "22C55E"},
}

def _xml_text(value: object) -> str:
    return escape(str(value or ""))

def _slide_lines(slide: dict) -> list[str]:
    layout = slide.get("layout", "content")
    lines: list[str] = []

    if layout == "title":
        lines.append(slide.get("subtitle", ""))
    elif layout == "two_column":
        lines.append(slide.get("leftTitle", ""))
        lines.extend(f"- {item}" for item in slide.get("leftContent", []))
        lines.append("")
        lines.append(slide.get("rightTitle", ""))
        lines.extend(f"- {item}" for item in slide.get("rightContent", []))
    elif layout == "stats":
        for stat in slide.get("stats", []):
            lines.append(f"{stat.get('value', '')} - {stat.get('label', '')}")
        if slide.get("content"):
            lines.append("")
            lines.append(slide.get("content", ""))
    elif layout == "quote":
        lines.append(slide.get("quote", ""))
        if slide.get("title"):
            lines.append(slide.get("title", ""))
    elif layout in {"section_break", "conclusion"}:
        lines.append(slide.get("subtitle", slide.get("content", "")))
    else:
        lines.extend(f"- {item}" for item in slide.get("bullets", []))
        if slide.get("content"):
            lines.append(slide.get("content", ""))

    return [line for line in lines if line is not None]

def _text_shape(shape_id: int, x: int, y: int, w: int, h: int,
                lines: list[str], font_size: int, color: str, bold: bool = False) -> str:
    paragraphs = []
    for line in lines or [""]:
        paragraphs.append(f"""
        <a:p>
          <a:r>
            <a:rPr lang="ru-RU" sz="{font_size}" b="{1 if bold else 0}">
              <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
            </a:rPr>
            <a:t>{_xml_text(line)}</a:t>
          </a:r>
        </a:p>""")

    return f"""
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="{shape_id}" name="TextBox {shape_id}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="{x}" y="{y}"/>
          <a:ext cx="{w}" cy="{h}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" anchor="top"/>
        <a:lstStyle/>
        {''.join(paragraphs)}
      </p:txBody>
    </p:sp>"""

def _slide_xml(slide: dict, index: int, colors: dict) -> str:
    title = slide.get("title") or slide.get("presentation_title") or f"Слайд {index}"
    body_lines = _slide_lines(slide)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr><a:solidFill><a:srgbClr val="{colors['bg']}"/></a:solidFill></p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {_text_shape(2, 650000, 650000, 10850000, 1300000, [title], 3600, colors["title"], True)}
      {_text_shape(3, 700000, 2150000, 10650000, 3900000, body_lines, 2100, colors["text"])}
      {_text_shape(4, 700000, 6350000, 2500000, 450000, [f"{index}"], 1400, colors["accent"], True)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

def build_pptx(slide_data: dict, output_path: str, theme: str) -> bool:
    """Build PPTX with python-pptx and embed generated images when present."""
    slides = slide_data.get("slides") or []
    if not slides:
        return False

    colors = THEMES.get(theme, THEMES["modern"])

    try:
        from pptx import Presentation
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        from pptx.util import Inches, Pt

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]

        def rgb(hex_color: str) -> RGBColor:
            return RGBColor.from_string(hex_color)

        def add_textbox(slide_obj, x, y, w, h, text, size=24, bold=False, color=None):
            box = slide_obj.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
            tf = box.text_frame
            tf.clear()
            tf.word_wrap = True
            for idx, line in enumerate(str(text or "").split("\n") or [""]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.text = line
                p.font.size = Pt(size)
                p.font.bold = bold
                p.font.name = "Arial"
                p.font.color.rgb = rgb(color or colors["text"])
            return box

        for index, item in enumerate(slides, start=1):
            slide = prs.slides.add_slide(blank)
            bg = slide.background.fill
            bg.solid()
            bg.fore_color.rgb = rgb(colors["bg"])

            title = item.get("title") or f"Слайд {index}"
            add_textbox(slide, 0.55, 0.35, 7.4, 1.0, title, size=30 if index > 1 else 38, bold=True, color=colors["title"])

            image_data = item.get("imageData")
            has_image = False
            if image_data:
                try:
                    raw = image_data.split(",", 1)[1] if "," in image_data else image_data
                    stream = io.BytesIO(base64.b64decode(raw))
                    slide.shapes.add_picture(stream, Inches(8.15), Inches(1.25), width=Inches(4.55), height=Inches(3.25))
                    has_image = True
                except Exception as e:
                    print(f"Image embed failed: {e}")

            body_width = 7.2 if has_image else 11.8
            layout = item.get("layout", "content")
            if layout == "title":
                add_textbox(slide, 0.7, 2.35, 11.5, 1.4, item.get("subtitle", ""), size=24, color=colors["text"])
            elif layout == "two_column":
                add_textbox(slide, 0.75, 1.55, 5.4, 0.45, item.get("leftTitle", "Слева"), size=20, bold=True, color=colors["accent"])
                add_textbox(slide, 6.7, 1.55, 5.4, 0.45, item.get("rightTitle", "Справа"), size=20, bold=True, color=colors["accent"])
                add_textbox(slide, 0.85, 2.2, 5.2, 3.9, "\n".join(f"• {x}" for x in item.get("leftContent", [])), size=18)
                add_textbox(slide, 6.8, 2.2, 5.2, 3.9, "\n".join(f"• {x}" for x in item.get("rightContent", [])), size=18)
            elif layout == "stats":
                stats = item.get("stats", [])[:3]
                for pos, stat in enumerate(stats):
                    x = 0.75 + pos * 3.9
                    add_textbox(slide, x, 1.55, 3.2, 0.7, stat.get("value", ""), size=30, bold=True, color=colors["accent"])
                    add_textbox(slide, x, 2.25, 3.2, 0.7, stat.get("label", ""), size=15)
                add_textbox(slide, 0.75, 3.45, body_width, 2.1, item.get("content", ""), size=20)
            elif layout == "quote":
                add_textbox(slide, 0.95, 2.0, body_width, 2.2, item.get("quote", ""), size=26, color=colors["title"])
            elif layout == "conclusion":
                add_textbox(slide, 0.75, 2.0, body_width, 2.4, item.get("content", item.get("subtitle", "")), size=24)
            else:
                bullets = item.get("bullets") or [item.get("content", "")]
                add_textbox(slide, 0.75, 1.65, body_width, 4.5, "\n".join(f"• {x}" for x in bullets if x), size=19)

            add_textbox(slide, 0.6, 6.85, 1.0, 0.3, str(index), size=12, bold=True, color=colors["accent"])

        prs.save(output_path)
        return True
    except Exception as e:
        print(f"PPTX generation failed: {e}")
        return False

def build_preview(slides: list[dict]) -> list[dict]:
    preview = []
    for i, s in enumerate(slides):
        preview.append({
            "index": i + 1,
            "layout": s.get("layout", "content"),
            "title": s.get("title", ""),
            "subtitle": s.get("subtitle", s.get("quote", "")),
            "content": s.get("content", ""),
            "quote": s.get("quote", ""),
            "bullets": s.get("bullets", []),
            "leftTitle": s.get("leftTitle", ""),
            "leftContent": s.get("leftContent", []),
            "rightTitle": s.get("rightTitle", ""),
            "rightContent": s.get("rightContent", []),
            "stats": s.get("stats", []),
            "image_prompt": s.get("image_prompt"),
            "hasImage": bool(s.get("imageData")),
        })
    return preview

def extract_document_text(filename: str, content: bytes) -> str:
    fname = (filename or "").lower()
    if not content:
        return ""
    if fname.endswith(".pdf"):
        return extract_text_from_pdf(content)
    if fname.endswith(".docx"):
        return extract_text_from_docx(content)
    return ""

def create_presentation_from_data(
    prompt: str,
    slide_count: int,
    style: str,
    tone: str,
    generate_images: bool,
    rt_token: str,
    rt_service: str,
    document_filename: str = "",
    document_content: bytes = b"",
) -> dict:
    session_id = str(uuid.uuid4())
    output_path = str(WORK_DIR / f"{session_id}.pptx")

    doc_text = extract_document_text(document_filename, document_content)

    slide_count = max(3, min(20, slide_count))
    structure = generate_slide_structure(prompt, doc_text, slide_count, style, tone, rt_token)

    slides = structure.get("slides", [])
    if generate_images and rt_token.strip():
        for slide in slides:
            if slide.get("image_prompt"):
                img_bytes = generate_image_rt(slide["image_prompt"], rt_token.strip(), rt_service)
                if img_bytes:
                    slide["imageData"] = image_bytes_to_base64(img_bytes)
                time.sleep(0.5)
    elif generate_images:
        for slide in slides:
            slide.pop("imageData", None)

    structure["slides"] = slides
    if not build_pptx(structure, output_path, style):
        raise HTTPException(500, "Ошибка генерации PPTX")

    with open(WORK_DIR / f"{session_id}.json", "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False, indent=2)

    return {
        "session_id": session_id,
        "title": structure.get("presentation_title", "Презентация"),
        "slide_count": len(slides),
        "preview": build_preview(slides),
        "warning": structure.get("generation_warning", ""),
        "source": structure.get("generation_source", ""),
    }

async def create_presentation(
    prompt: str,
    slide_count: int,
    style: str,
    tone: str,
    generate_images: bool,
    rt_token: str,
    rt_service: str,
    document: Optional[UploadFile],
) -> dict:
    filename = ""
    content = b""
    if document and document.filename:
        filename = document.filename
        content = await document.read()
    return create_presentation_from_data(
        prompt, slide_count, style, tone, generate_images, rt_token, rt_service, filename, content
    )

def render_result_page(data: dict) -> str:
    slides_html = "".join(
        f"<li><strong>{_xml_text(s.get('title'))}</strong><br>{_xml_text(s.get('subtitle') or s.get('content') or ' '.join(s.get('bullets') or []))}</li>"
        for s in data["preview"]
    )
    warning = f"<p class='warning'>{_xml_text(data['warning'])}</p>" if data.get("warning") else ""
    return f"""<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Презентация готова</title>
<style>
body{{margin:0;background:#080313;color:#f8f5ff;font-family:Arial,sans-serif;padding:40px}}
.wrap{{max-width:920px;margin:auto}}a,.btn{{display:inline-block;background:#8b5cf6;color:white;padding:14px 20px;border-radius:10px;text-decoration:none;font-weight:700;margin:8px 8px 18px 0}}
.card{{background:#160d2e;border:1px solid rgba(139,92,246,.35);border-radius:14px;padding:22px}}
li{{margin:0 0 14px;line-height:1.45}}.warning{{color:#fed7aa}}.src{{color:#a78bfa;font-weight:700}}
</style></head><body><div class="wrap">
<h1>{_xml_text(data["title"])}</h1>
<p class="src">Источник генерации: {_xml_text(data.get("source") or "неизвестен")}</p>
{warning}
<a class="btn" href="/api/download/{data["session_id"]}">Скачать PPTX</a>
<a class="btn" href="/">Создать еще</a>
<div class="card"><h2>Предпросмотр</h2><ol>{slides_html}</ol></div>
</div></body></html>"""

def render_loading_page(job_id: str) -> str:
    safe_job_id = _xml_text(job_id)
    return f"""<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Генерация презентации</title>
<style>
*{{box-sizing:border-box}}body{{margin:0;min-height:100vh;background:#07040f;color:#f5f1ff;font-family:Arial,sans-serif;display:grid;place-items:center;padding:28px}}
.card{{width:min(760px,100%);background:#100821;border:1px solid rgba(139,92,246,.35);border-radius:22px;padding:38px;text-align:center;box-shadow:0 24px 90px rgba(139,92,246,.22)}}
.logo{{width:58px;height:58px;border-radius:16px;margin:0 auto 18px;display:grid;place-items:center;background:linear-gradient(135deg,#8b5cf6,#ec4e99);font-size:28px}}
h1{{margin:0 0 8px;font-size:clamp(28px,5vw,44px)}}p{{margin:0 0 24px;color:#a99ac8;line-height:1.5}}
.percent{{font-size:52px;font-weight:800;color:#a78bfa;margin:4px 0 16px}}
.bar{{height:14px;background:#21143e;border-radius:999px;overflow:hidden;border:1px solid rgba(255,255,255,.08)}}
.fill{{height:100%;width:0;background:linear-gradient(90deg,#8b5cf6,#ec4e99);border-radius:999px;transition:width .45s ease}}
.steps{{display:grid;gap:10px;margin-top:24px;text-align:left}}.step{{padding:12px 14px;border-radius:12px;background:rgba(255,255,255,.04);color:#b9aad6}}.step.done{{color:#fff;background:rgba(34,211,238,.12)}}
.error{{display:none;margin-top:20px;padding:14px;border-radius:12px;background:rgba(239,68,68,.14);color:#fecaca;text-align:left;white-space:pre-wrap}}
a{{color:#c4b5fd}}
</style></head><body>
<main class="card">
  <div class="logo">✦</div>
  <h1 id="title">Генерируем презентацию</h1>
  <p id="sub">Не закрывайте страницу. Сервис анализирует запрос и собирает PPTX.</p>
  <div class="percent" id="percent">0%</div>
  <div class="bar"><div class="fill" id="fill"></div></div>
  <div class="steps">
    <div class="step" id="s1">Чтение документа и параметров</div>
    <div class="step" id="s2">Генерация структуры слайдов через LLM</div>
    <div class="step" id="s3">Подготовка изображений и оформления</div>
    <div class="step" id="s4">Сборка PPTX и страницы результата</div>
  </div>
  <div class="error" id="error"></div>
</main>
<script>
const jobId = "{safe_job_id}";
const percent = document.getElementById('percent');
const fill = document.getElementById('fill');
const title = document.getElementById('title');
const sub = document.getElementById('sub');
let value = 0;
function setProgress(next) {{
  value = Math.max(value, Math.min(100, Math.round(next)));
  percent.textContent = value + '%';
  fill.style.width = value + '%';
  if (value >= 18) document.getElementById('s1').classList.add('done');
  if (value >= 42) document.getElementById('s2').classList.add('done');
  if (value >= 70) document.getElementById('s3').classList.add('done');
  if (value >= 92) document.getElementById('s4').classList.add('done');
}}
const timer = setInterval(() => {{
  const step = value < 50 ? 4 : value < 82 ? 2 : 1;
  setProgress(Math.min(96, value + step));
}}, 650);
async function run() {{
  try {{
    setProgress(6);
    const resp = await fetch('/api/jobs/' + encodeURIComponent(jobId) + '/run', {{method:'POST'}});
    const data = await resp.json().catch(() => ({{}}));
    if (!resp.ok) throw new Error(data.detail || 'Ошибка генерации');
    clearInterval(timer);
    title.textContent = 'Презентация готова';
    sub.textContent = 'Открываем страницу скачивания и предпросмотра.';
    setProgress(100);
    setTimeout(() => {{ window.location.href = '/result/' + encodeURIComponent(data.session_id); }}, 700);
  }} catch (err) {{
    clearInterval(timer);
    document.getElementById('error').style.display = 'block';
    document.getElementById('error').innerHTML = '<strong>Ошибка:</strong> ' + String(err.message || err) + '<br><br><a href="/">Вернуться к форме</a>';
  }}
}}
run();
</script>
</body></html>"""

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
    try:
        return JSONResponse(await create_presentation(
            prompt, slide_count, style, tone, generate_images, rt_token, rt_service, document
        ))

    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Ошибка парсинга ответа LLM: {e}")
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.post("/loading", response_class=HTMLResponse)
async def create_loading_job(
    prompt: str = Form(...),
    slide_count: int = Form(8),
    style: str = Form("rostelecom"),
    tone: str = Form("professional"),
    generate_images: bool = Form(False),
    rt_token: str = Form(""),
    rt_service: str = Form("sd"),
    document: Optional[UploadFile] = File(None),
):
    filename = ""
    content = b""
    if document and document.filename:
        filename = document.filename
        content = await document.read()

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "prompt": prompt,
        "slide_count": slide_count,
        "style": style,
        "tone": tone,
        "generate_images": generate_images,
        "rt_token": rt_token,
        "rt_service": rt_service,
        "document_filename": filename,
        "document_content": content,
        "created_at": time.time(),
    }
    return render_loading_page(job_id)


@app.post("/api/jobs/{job_id}/run")
async def run_loading_job(job_id: str):
    job = JOBS.pop(job_id, None)
    if not job:
        raise HTTPException(404, "Задача генерации не найдена. Вернитесь к форме и запустите генерацию заново.")
    try:
        data = create_presentation_from_data(
            job["prompt"],
            job["slide_count"],
            job["style"],
            job["tone"],
            job["generate_images"],
            job["rt_token"],
            job["rt_service"],
            job["document_filename"],
            job["document_content"],
        )
        return JSONResponse(data)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.post("/generate", response_class=HTMLResponse)
async def generate_presentation_page(
    prompt: str = Form(...),
    slide_count: int = Form(8),
    style: str = Form("rostelecom"),
    tone: str = Form("professional"),
    generate_images: bool = Form(False),
    rt_token: str = Form(""),
    rt_service: str = Form("sd"),
    document: Optional[UploadFile] = File(None),
):
    try:
        data = await create_presentation(
            prompt, slide_count, style, tone, generate_images, rt_token, rt_service, document
        )
        return render_result_page(data)
    except Exception as e:
        traceback.print_exc()
        return HTMLResponse(f"<h1>Ошибка генерации</h1><pre>{_xml_text(e)}</pre><p><a href='/'>Назад</a></p>", status_code=500)


@app.post("/api/rebuild/{session_id}")
async def rebuild_pptx(session_id: str, request: Request):
    try:
        payload = await request.json()
        slides = payload.get("slides")
        style = payload.get("style", "modern")
        title = payload.get("title", "Презентация")
        if not isinstance(slides, list) or not slides:
            raise HTTPException(400, "Нет данных слайдов для пересборки")

        structure = {
            "presentation_title": title,
            "metadata": {"title": title, "author": "PresentAI", "contact": ""},
            "slides": slides,
        }
        output_path = str(WORK_DIR / f"{session_id}.pptx")
        if not build_pptx(structure, output_path, style):
            raise HTTPException(500, "Ошибка пересборки PPTX")

        with open(WORK_DIR / f"{session_id}.json", "w", encoding="utf-8") as f:
            json.dump(structure, f, ensure_ascii=False, indent=2)

        return JSONResponse({
            "session_id": session_id,
            "title": title,
            "slide_count": len(slides),
            "preview": build_preview(slides),
        })
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.get("/api/demo")
async def demo_pptx():
    """Generate and download a demo PPTX without JavaScript or API tokens."""
    session_id = str(uuid.uuid4())
    output_path = str(WORK_DIR / f"{session_id}.pptx")
    structure = generate_slide_structure(
        "Демо-презентация: AI-генератор презентаций для Ростелекома",
        "Проверка локальной генерации без API-ключей. Сервис принимает промпт, документ PDF или DOCX, настройки стиля и тона, затем собирает PPTX.",
        6,
        "rostelecom",
        "professional",
        "",
    )
    if not build_pptx(structure, output_path, "rostelecom"):
        raise HTTPException(500, "Ошибка генерации demo PPTX")
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="presentai-demo.pptx",
    )


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


@app.get("/result/{session_id}", response_class=HTMLResponse)
async def result_page(session_id: str):
    json_path = WORK_DIR / f"{session_id}.json"
    pptx_path = WORK_DIR / f"{session_id}.pptx"
    if not json_path.exists() or not pptx_path.exists():
        return HTMLResponse(
            "<h1>Презентация не найдена</h1><p><a href='/'>Вернуться к форме</a></p>",
            status_code=404,
        )
    with open(json_path, "r", encoding="utf-8") as f:
        structure = json.load(f)
    slides = structure.get("slides", [])
    return render_result_page({
        "session_id": session_id,
        "title": structure.get("presentation_title", "Презентация"),
        "slide_count": len(slides),
        "preview": build_preview(slides),
        "warning": structure.get("generation_warning", ""),
        "source": structure.get("generation_source", ""),
    })


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

/* ── Capability status ──────────────────────────────── */
.capability-grid {
  display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px;
  margin-bottom: 30px; animation: fadeUp 0.7s 0.15s ease both;
}
.capability-item {
  min-height: 76px; padding: 12px; border-radius: 12px;
  border: 1px solid var(--border);
  background: rgba(22,13,46,0.78);
  display: grid; align-content: center; gap: 5px;
}
.capability-item strong {
  color: var(--text); font-size: 13px; font-weight: 800;
}
.capability-item span {
  color: var(--muted); font-size: 11px; line-height: 1.35;
}
.capability-item.ready { border-color: rgba(34,211,238,0.35); }
.capability-item.ready strong { color: var(--accent-c); }
@media (max-width: 760px) {
  .capability-grid { grid-template-columns: repeat(2, 1fr); }
}

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
.slider-wrap {
  --slider-progress: 25%;
  display: grid; grid-template-columns: minmax(0, 1fr) 96px; align-items: center; gap: 16px;
  width: 100%; max-width: none;
  padding: 10px 12px; border: 1px solid var(--border); border-radius: 12px;
  background: rgba(139,92,246,0.06);
}
input[type="range"] {
  width: 100%; -webkit-appearance: none; appearance: none;
  background: transparent; height: 28px; outline: none; cursor: pointer;
}
input[type="range"]:focus {
  box-shadow: none;
}
input[type="range"]::-webkit-slider-runnable-track {
  background: linear-gradient(90deg,
    var(--accent-a) 0%,
    var(--accent-b) var(--slider-progress),
    rgba(139,92,246,0.18) var(--slider-progress),
    rgba(139,92,246,0.18) 100%);
  height: 8px; border-radius: 999px;
  box-shadow: inset 0 0 0 1px rgba(255,255,255,0.12);
}
input[type="range"]::-moz-range-track {
  background: rgba(139,92,246,0.18);
  height: 8px; border-radius: 999px;
}
input[type="range"]::-moz-range-progress {
  background: linear-gradient(90deg, var(--accent-a), var(--accent-b));
  height: 8px; border-radius: 999px;
}
input[type="range"]::-webkit-slider-thumb {
  -webkit-appearance: none;
  width: 20px; height: 20px; border-radius: 50%;
  background: #fff; margin-top: -6px; cursor: pointer;
  border: 3px solid var(--accent-a);
  box-shadow: 0 0 12px rgba(139,92,246,0.55);
}
input[type="range"]::-moz-range-thumb {
  width: 20px; height: 20px; border-radius: 50%;
  background: #fff; cursor: pointer;
  border: 3px solid var(--accent-a);
  box-shadow: 0 0 12px rgba(139,92,246,0.55);
}
.slider-val {
  min-width: 112px; min-height: 54px; padding: 6px 8px;
  display: grid; grid-template-columns: 24px 1fr 24px; grid-template-rows: 1fr auto;
  align-items: center; justify-items: center; text-align: center; column-gap: 4px;
  border-radius: 12px; color: #fff;
  background: linear-gradient(135deg, rgba(139,92,246,0.95), rgba(236,78,153,0.82));
  border: 1px solid rgba(255,255,255,0.18);
  box-shadow: 0 8px 22px rgba(139,92,246,0.22);
}
.slider-val .num {
  grid-column: 2; font-feature-settings: "tnum";
  font-family: Arial, system-ui, sans-serif; font-weight: 800;
  font-size: 1.45rem; line-height: 1; color: #fff;
  font-stretch: normal; letter-spacing: 0;
  min-width: 2ch;
}
.slider-step {
  width: 24px; height: 24px; border-radius: 8px;
  border: 1px solid rgba(255,255,255,0.2);
  background: rgba(255,255,255,0.1); color: #fff;
  font-size: 16px; line-height: 1; cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  padding: 0;
}
.slider-step:hover { background: rgba(255,255,255,0.18); }
.slider-step.minus { grid-column: 1; }
.slider-step.plus { grid-column: 3; }
.slider-hidden-input { display: none; }
.slider-val:focus-within {
  box-shadow: 0 0 0 3px rgba(139,92,246,0.18), 0 8px 22px rgba(139,92,246,0.22);
}
.slider-val .unit {
  grid-column: 1 / -1; display: block; margin-top: 2px; font-size: 10px;
  color: rgba(255,255,255,0.74); font-weight: 700;
}
@media (max-width: 520px) {
  .slider-wrap { width: 100%; grid-template-columns: 1fr 104px; gap: 10px; }
  input[type="range"] { width: 100%; }
}

/* ── Toggle switch ──────────────────────────────────── */
.toggle-wrap {
  display: flex; align-items: center; gap: 12px; cursor: pointer;
  padding: 12px; border: 1px solid var(--border-bright); border-radius: 12px;
  background: rgba(139,92,246,0.08);
}
.toggle { position: relative; width: 52px; height: 30px; flex: 0 0 auto; }
.toggle input {
  position: absolute; inset: 0; width: 100%; height: 100%;
  opacity: 0; cursor: pointer; z-index: 3;
}
.toggle-track {
  position: absolute; inset: 0; border-radius: 12px;
  background: var(--panel); border: 1px solid var(--border);
  transition: all 0.3s;
}
.toggle-thumb {
  position: absolute; top: 4px; left: 4px;
  width: 22px; height: 22px; border-radius: 50%;
  background: var(--muted); transition: all 0.3s;
}
.toggle input:checked ~ .toggle-track {
  background: rgba(139,92,246,0.25); border-color: var(--accent-a);
}
.toggle input:checked ~ .toggle-thumb {
  transform: translateX(22px); background: #fff;
  box-shadow: 0 0 10px rgba(139,92,246,0.5);
}
.toggle-label { font-size: 14px; color: var(--text); font-weight: 700; }

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
.loading-mode .hero,
.loading-mode .capability-grid,
.loading-mode #generator-form,
.loading-mode footer {
  display: none;
}
.loading-mode #progress-section {
  display: block;
  min-height: calc(100vh - 120px);
  margin-top: 34px;
}
.loading-mode .progress-card {
  min-height: 560px;
  display: flex; flex-direction: column; justify-content: center;
  border-color: var(--border-bright);
  box-shadow: 0 18px 80px rgba(139,92,246,0.18);
}
.progress-emoji { font-size: 40px; margin-bottom: 12px; animation: pulse 1.5s ease-in-out infinite; }
@keyframes pulse { 0%,100%{transform:scale(1)} 50%{transform:scale(1.1)} }
.progress-title { font-family: var(--font-head); font-size: 1.2rem; font-weight: 700; margin-bottom: 6px; }
.loading-mode .progress-title { font-size: clamp(1.6rem, 4vw, 2.45rem); }
.progress-sub { font-size: 13px; color: var(--muted); margin-bottom: 20px; }
.progress-percent {
  font-family: var(--font-head); font-size: 2.2rem; font-weight: 800;
  color: var(--accent-a); margin-bottom: 14px;
}
.progress-bar-wrap {
  background: var(--panel); border-radius: 100px; height: 8px; overflow: hidden;
}
.loading-mode .progress-bar-wrap { height: 12px; max-width: 620px; width: 100%; margin: 0 auto; }
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
  border-radius: 12px; overflow: hidden;
  transition: all 0.2s;
}
.slide-card:hover { border-color: var(--border-bright); transform: translateY(-3px); box-shadow: 0 8px 24px rgba(0,0,0,0.3); }

.slide-editor { padding: 12px; display: grid; gap: 8px; }
.slide-editor input, .slide-editor textarea {
  width: 100%; min-height: 0; padding: 9px 10px; border-radius: 8px;
  font-size: 12px; color: var(--text); background: var(--surface);
  border: 1px solid var(--border);
}
.slide-editor textarea { min-height: 78px; resize: vertical; }
.btn-secondary {
  display: inline-flex; align-items: center; justify-content: center; gap: 8px;
  padding: 13px 18px; border: 1px solid var(--border-bright); border-radius: 12px;
  background: var(--surface); color: var(--text); font-weight: 700; cursor: pointer;
  transition: all 0.25s;
}
.btn-secondary:hover { border-color: var(--accent-a); transform: translateY(-2px); }
.warning-card {
  display: none; margin: 0 0 16px; padding: 12px 14px; border-radius: 12px;
  background: rgba(249,115,22,0.12); border: 1px solid rgba(249,115,22,0.35);
  color: #fed7aa; font-size: 13px; line-height: 1.45;
}
.source-badge {
  display: inline-flex; align-items: center; gap: 8px; margin: 0 0 16px;
  padding: 8px 12px; border-radius: 999px; border: 1px solid var(--border-bright);
  background: rgba(139,92,246,0.12); color: var(--text); font-size: 13px; font-weight: 700;
}
.preview-panel {
  background: var(--surface); border: 1px solid var(--border); border-radius: 14px;
  padding: 16px; margin-bottom: 18px;
}
.preview-toolbar {
  display: flex; align-items: center; justify-content: space-between; gap: 12px;
  margin-bottom: 12px; color: var(--muted); font-size: 13px;
}
.preview-slide {
  aspect-ratio: 16 / 9; border-radius: 10px; overflow: hidden;
  background: #101828; border: 1px solid rgba(255,255,255,0.08);
  padding: 5%; display: flex; flex-direction: column; justify-content: center;
  box-shadow: inset 0 0 80px rgba(119,0,255,0.16);
}
.preview-slide h2 {
  font-family: var(--font-head); font-size: clamp(1.35rem, 4vw, 2.6rem);
  line-height: 1.08; margin-bottom: 18px; color: #fff;
}
.preview-slide p, .preview-slide li { color: rgba(255,255,255,0.82); font-size: clamp(0.9rem, 2vw, 1.2rem); line-height: 1.45; }
.preview-slide ul { padding-left: 22px; display: grid; gap: 8px; }
.preview-columns { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
.preview-column { background: rgba(255,255,255,0.06); border-radius: 10px; padding: 14px; }
.preview-column h3 { color: #fff; margin-bottom: 10px; font-size: 1rem; }
.preview-stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 16px; }
.preview-stat { background: rgba(255,255,255,0.08); border-radius: 10px; padding: 14px; text-align: center; }
.preview-stat strong { display: block; color: var(--accent-a); font-size: 1.35rem; margin-bottom: 4px; }
.slide-card.active { border-color: var(--accent-a); box-shadow: 0 0 0 2px rgba(139,92,246,0.22); }

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

  <div class="capability-grid" aria-label="Возможности сервиса">
    <div class="capability-item ready">
      <strong>RT LLM</strong>
      <span>структура и текст слайдов</span>
    </div>
    <div class="capability-item ready">
      <strong>PDF / DOCX</strong>
      <span>анализ загруженного документа</span>
    </div>
    <div class="capability-item ready">
      <strong>SD / Yandex ART</strong>
      <span>изображения по промптам</span>
    </div>
    <div class="capability-item ready">
      <strong>PPTX</strong>
      <span>сборка и скачивание файла</span>
    </div>
  </div>

  <form id="generator-form" action="/loading" method="post" enctype="multipart/form-data">
  <!-- Step 1: Prompt + Document -->
  <div class="card" id="form-section">
    <div class="card-title"><div class="icon">💬</div> Запрос и документ</div>
    <div class="field">
      <label>Описание презентации *</label>
      <textarea id="prompt" name="prompt" required placeholder="Например: Создай презентацию о преимуществах AI в телекоммуникациях для руководителей компании. Включи статистику, кейсы и план внедрения."></textarea>
    </div>
    <div class="field">
      <label>Исходный документ (необязательно)</label>
      <div class="upload-zone" id="upload-zone">
        <input type="file" id="doc-file" name="document" accept=".pdf,.docx" onchange="handleFileSelect(this)">
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
        <input type="range" id="slide-count-range" min="4" max="20" value="8"
               data-slide-count-range>
        <div class="slider-val" id="slide-count-val">
          <button class="slider-step minus" type="button" data-slide-delta="-1" aria-label="Уменьшить количество слайдов">−</button>
          <span class="num" id="slide-count-display">8</span>
          <button class="slider-step plus" type="button" data-slide-delta="1" aria-label="Увеличить количество слайдов">+</button>
          <input class="slider-hidden-input" type="hidden" id="slide-count" name="slide_count" value="8">
          <span class="unit">слайдов</span>
        </div>
      </div>
      <script>
      (function () {
        var range = document.getElementById('slide-count-range');
        var hidden = document.getElementById('slide-count');
        var display = document.getElementById('slide-count-display');
        var unit = document.querySelector('#slide-count-val .unit');
        var minus = document.querySelector('[data-slide-delta="-1"]');
        var plus = document.querySelector('[data-slide-delta="1"]');

        function word(n) {
          var last = n % 10;
          var lastTwo = n % 100;
          if (last === 1 && lastTwo !== 11) return 'слайд';
          if ([2, 3, 4].indexOf(last) !== -1 && [12, 13, 14].indexOf(lastTwo) === -1) return 'слайда';
          return 'слайдов';
        }

        function clamp(value) {
          value = parseInt(value, 10);
          if (isNaN(value)) value = 8;
          return Math.max(4, Math.min(20, value));
        }

        function render(value) {
          value = clamp(value);
          var min = parseInt(range.min, 10);
          var max = parseInt(range.max, 10);
          var pct = ((value - min) / (max - min)) * 100;
          range.value = String(value);
          hidden.value = String(value);
          display.textContent = String(value);
          unit.textContent = word(value);
          range.parentElement.style.setProperty('--slider-progress', pct + '%');
        }

        range.addEventListener('input', function () { render(range.value); });
        range.addEventListener('change', function () { render(range.value); });
        minus.addEventListener('click', function () { render(clamp(hidden.value) - 1); });
        plus.addEventListener('click', function () { render(clamp(hidden.value) + 1); });
        render(hidden.value);
      })();
      </script>
    </div>

    <div class="fields-row">
      <div class="field">
        <label>Тон</label>
        <select id="tone" name="tone">
          <option value="professional">Профессиональный</option>
          <option value="creative">Творческий</option>
          <option value="academic">Академический</option>
          <option value="casual">Дружелюбный</option>
          <option value="persuasive">Убедительный</option>
        </select>
      </div>
      <div class="field">
        <label>Визуальный стиль</label>
        <select id="style" name="style">
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
      <div class="field">
        <label>API Token Ростелеком для LLM</label>
        <input type="text" id="rt-token" name="rt_token" placeholder="Bearer-токен для ai.rt.ru">
      </div>
      <label class="toggle-wrap" for="gen-images">
        <div class="toggle">
          <input type="checkbox" id="gen-images" name="generate_images" value="true">
          <div class="toggle-track"></div>
          <div class="toggle-thumb"></div>
        </div>
        <span class="toggle-label">Генерировать изображения (RT API)</span>
      </label>
      <div id="rt-fields" style="margin-top:14px">
        <div class="fields-row">
          <div class="field">
            <label>Модель изображений</label>
            <select id="rt-service" name="rt_service">
              <option value="yaArt">Yandex ART</option>
              <option value="sd">Stable Diffusion</option>
            </select>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Generate button -->
  <button class="btn-generate" id="gen-btn" type="submit">
    <span>✦</span> Сгенерировать презентацию
  </button>
  <a class="btn-secondary" style="width:100%; margin:12px 0 0; text-decoration:none" href="/api/demo">
    Скачать демо PPTX без API
  </a>
  </form>

  <!-- Progress -->
  <div id="progress-section">
    <div class="progress-card">
      <div class="progress-emoji" id="progress-emoji">🧠</div>
      <div class="progress-title" id="progress-title">Анализируем запрос...</div>
      <div class="progress-sub" id="progress-sub">LLM обрабатывает ваш запрос</div>
      <div class="progress-percent" id="progress-percent">0%</div>
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
      <button class="btn-secondary" id="save-edits-btn" onclick="saveEdits()">
        ↻ Обновить PPTX
      </button>
    </div>

    <div class="warning-card" id="warning-card"></div>
    <div class="source-badge" id="source-badge">Источник: ожидает генерации</div>
    <div class="preview-panel">
      <div class="preview-toolbar">
        <span id="preview-counter">Слайд 1</span>
        <span>Кликните карточку ниже, чтобы открыть слайд</span>
      </div>
      <div class="preview-slide" id="preview-slide"></div>
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
let currentSlides = [];
let currentStyle = 'rostelecom';

window.addEventListener('error', event => {
  const message = event.message || 'Неизвестная ошибка JavaScript';
  const card = document.getElementById('error-card');
  if (card) {
    const safeMessage = String(message).replace(/[&<>"']/g, ch => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[ch]));
    card.style.display = 'block';
    card.innerHTML = `<strong>Ошибка JavaScript:</strong> ${safeMessage}`;
  }
});

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
  // The old visual stepper was replaced by functional capability cards.
}

function setProgress(pct, emoji, title, sub) {
  document.getElementById('progress-bar').style.width = pct + '%';
  document.getElementById('progress-percent').textContent = Math.round(pct) + '%';
  if (emoji) document.getElementById('progress-emoji').textContent = emoji;
  if (title) document.getElementById('progress-title').textContent = title;
  if (sub) document.getElementById('progress-sub').textContent = sub;
}

function doneStep(id) {
  document.getElementById(id).classList.add('done');
}

function clampSlideCount(value) {
  return Math.max(4, Math.min(20, Number(value) || 8));
}

function updateSlideCountUI(value) {
  const slider = document.getElementById('slide-count-range');
  const hidden = document.getElementById('slide-count');
  const display = document.getElementById('slide-count-display');
  const unit = document.querySelector('#slide-count-val .unit');
  value = clampSlideCount(value);
  slider.value = String(value);
  hidden.value = String(value);
  display.textContent = String(value);
  unit.textContent = getSlideWord(value);
  const min = Number(slider.min);
  const max = Number(slider.max);
  const pct = ((value - min) / (max - min)) * 100;
  slider.closest('.slider-wrap').style.setProperty('--slider-progress', `${pct}%`);
}

function syncSlideCountFromRange(slider) {
  updateSlideCountUI(slider.value);
}

function syncSlideCountFromNumber(input) {
  updateSlideCountUI(input.value);
}

function changeSlideCount(delta) {
  const current = Number(document.getElementById('slide-count').value);
  updateSlideCountUI(current + delta);
}

function bindSlideCountControls() {
  const slider = document.getElementById('slide-count-range');
  const hidden = document.getElementById('slide-count');
  slider.addEventListener('input', () => updateSlideCountUI(slider.value));
  slider.addEventListener('change', () => updateSlideCountUI(slider.value));
  document.addEventListener('click', event => {
    const button = event.target.closest('[data-slide-delta]');
    if (!button) return;
    event.preventDefault();
    updateSlideCountUI(Number(hidden.value) + Number(button.dataset.slideDelta));
  });
  updateSlideCountUI(hidden.value);
}

function getSlideWord(value) {
  const last = value % 10;
  const lastTwo = value % 100;
  if (last === 1 && lastTwo !== 11) return 'слайд';
  if ([2, 3, 4].includes(last) && ![12, 13, 14].includes(lastTwo)) return 'слайда';
  return 'слайдов';
}

function showError(message) {
  document.body.classList.remove('loading-mode');
  document.getElementById('progress-section').style.display = 'none';
  const errCard = document.getElementById('error-card');
  errCard.style.display = 'block';
  errCard.innerHTML = `<strong>Ошибка:</strong> ${escapeHtml(message)}`;
  const btn = document.getElementById('gen-btn');
  btn.disabled = false;
  btn.innerHTML = '<span>✦</span> Сгенерировать презентацию';
  setStep(1);
}

async function generate() {
  let progressTimer = null;
  try {
    const prompt = document.getElementById('prompt').value.trim();
    if (!prompt) {
      showError('Введите описание презентации');
      return;
    }

    const genImages = document.getElementById('gen-images').checked;
    const rtToken = document.getElementById('rt-token').value.trim();
    const rtService = document.getElementById('rt-service').value;
    const slideCount = document.getElementById('slide-count').value;
    const style = document.getElementById('style').value;
    const tone = document.getElementById('tone').value;
    currentStyle = style;

    const btn = document.getElementById('gen-btn');
    btn.disabled = true;
    btn.innerHTML = '<span>✦</span> Генерируем...';
    document.getElementById('error-card').style.display = 'none';
    document.getElementById('result-section').style.display = 'none';
    ['ps1','ps2','ps3','ps4'].forEach(id => document.getElementById(id).classList.remove('done'));
    document.body.classList.add('loading-mode');
    document.getElementById('progress-section').style.display = 'block';
    window.scrollTo({ top: 0, behavior: 'smooth' });
    setStep(3);
    setProgress(3, '📄', 'Генерация презентации', 'Подготавливаем запрос и документ');

    const loadingSteps = [
      { at: 12, emoji: '📄', title: 'Чтение входных данных', sub: 'Учитываем промпт, PDF/DOCX и параметры', done: 'ps1' },
      { at: 34, emoji: '🧠', title: 'RT LLM формирует структуру', sub: 'Создаем заголовки, тезисы и логику слайдов', done: 'ps2' },
      { at: 62, emoji: '🎨', title: 'Готовим визуальную часть', sub: 'Обрабатываем стиль и изображения, если они включены', done: 'ps3' },
      { at: 86, emoji: '📦', title: 'Собираем PPTX', sub: 'Формируем файл презентации для скачивания', done: 'ps4' },
    ];
    let visualProgress = 3;
    progressTimer = setInterval(() => {
      visualProgress = Math.min(96, visualProgress + (visualProgress < 60 ? 3 : 1));
      const active = [...loadingSteps].reverse().find(step => visualProgress >= step.at);
      if (active) {
        doneStep(active.done);
        setProgress(visualProgress, active.emoji, active.title, active.sub);
      } else {
        setProgress(visualProgress);
      }
    }, 700);

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

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 120000);
    const resp = await fetch('/api/generate', {
      method: 'POST',
      body: formData,
      signal: controller.signal
    });
    clearTimeout(timeoutId);

    const data = await resp.json().catch(() => ({}));

    if (!resp.ok) throw new Error(data.detail || 'Ошибка генерации');

    clearInterval(progressTimer);
    doneStep('ps4');
    setProgress(100, '🎉', 'Готово!', 'Презентация создана, открываем результат');

    setTimeout(() => {
      document.body.classList.remove('loading-mode');
      showResults(data);
    }, 800);

  } catch (err) {
    if (progressTimer) clearInterval(progressTimer);
    const message = err.name === 'AbortError'
      ? 'Сервер не ответил за 120 секунд. Попробуйте временно выключить генерацию изображений или повторить запрос.'
      : (err.message || String(err));
    showError(message);
  }
}

function showResults(data) {
  sessionId = data.session_id;
  currentSlides = data.preview || [];
  setStep(4);

  document.getElementById('progress-section').style.display = 'none';
  document.getElementById('result-section').style.display = 'block';
  document.getElementById('result-title').textContent = data.title || 'Презентация готова!';
  document.getElementById('result-meta').textContent = `${data.slide_count} слайдов`;
  document.getElementById('download-btn').href = `/api/download/${data.session_id}`;
  document.getElementById('warning-card').style.display = data.warning ? 'block' : 'none';
  document.getElementById('warning-card').textContent = data.warning || '';
  document.getElementById('source-badge').textContent = data.source ? `Источник: ${data.source}` : 'Источник: неизвестен';

  const grid = document.getElementById('slides-grid');
  grid.innerHTML = '';

  currentSlides.forEach(s => {
    grid.appendChild(buildSlideCard(s));
  });
  renderPreview(0);

  document.getElementById('gen-btn').disabled = false;
  document.getElementById('gen-btn').innerHTML = '<span>✦</span> Сгенерировать презентацию';
}

const LAYOUT_LABELS = {
  title: 'Титульный', content: 'Контент', two_column: 'Два столбца',
  stats: 'Статистика', quote: 'Цитата', section_break: 'Раздел',
  image: 'Изображение', conclusion: 'Заключение', section: 'Раздел',
};

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[ch]));
}

function buildSlideCard(s) {
  const card = document.createElement('div');
  card.className = 'slide-card';
  card.dataset.index = String((s.index || 1) - 1);
  card.addEventListener('click', event => {
    if (event.target.closest('input, textarea')) return;
    renderPreview(Number(card.dataset.index));
  });

  const layout = (s.layout || 'content').replace('-', '_');
  const thumbCls = `slide-thumb layout-${layout}`;

  let contentHtml = '';

  if (s.stats && s.stats.length > 0) {
    contentHtml = `<div class="mini-stats">${s.stats.slice(0,3).map(st =>
      `<div class="mini-stat"><div class="mini-stat-val">${escapeHtml(st.value)}</div><div class="mini-stat-lab">${escapeHtml(st.label)}</div></div>`
    ).join('')}</div>`;
  } else if (s.bullets && s.bullets.length > 0) {
    contentHtml = `<div class="mini-bullets">${s.bullets.slice(0,3).map(b =>
      `<div class="mini-bullet">${escapeHtml(b.substring(0,50))}${b.length>50?'...':''}</div>`
    ).join('')}</div>`;
  } else if (s.subtitle) {
    contentHtml = `<div class="slide-thumb-content">${escapeHtml(s.subtitle)}</div>`;
  } else if (s.leftTitle || s.rightTitle) {
    contentHtml = `<div style="display:flex;gap:6px;margin-top:4px">
      <div style="flex:1;background:rgba(139,92,246,0.15);border-radius:4px;padding:4px 6px;font-size:9px;color:rgba(255,255,255,0.5)">${escapeHtml(s.leftTitle||'')}</div>
      <div style="flex:1;background:rgba(236,78,153,0.15);border-radius:4px;padding:4px 6px;font-size:9px;color:rgba(255,255,255,0.5)">${escapeHtml(s.rightTitle||'')}</div>
    </div>`;
  }

  const editorBody = (s.bullets && s.bullets.length)
    ? s.bullets.join('\n')
    : (s.content || s.subtitle || s.quote || '');

  card.innerHTML = `
    <div class="${thumbCls}">
      <div class="slide-num">${s.index}</div>
      <div class="slide-layout-badge">${LAYOUT_LABELS[layout] || layout}</div>
      <div class="slide-thumb-title">${escapeHtml(s.title || '')}</div>
      ${contentHtml}
      ${s.hasImage ? '<div style="margin-top:6px;font-size:9px;color:rgba(139,92,246,0.7)">🖼 + изображение</div>' : ''}
    </div>
    <div class="slide-editor">
      <input class="edit-title" value="${escapeHtml(s.title || '')}" placeholder="Заголовок">
      <textarea class="edit-body" placeholder="Текст, пункты или цитата">${escapeHtml(editorBody)}</textarea>
    </div>
    <div class="slide-meta">
      <span>${LAYOUT_LABELS[layout] || layout}</span>
      <span style="color:var(--accent-a)">→</span>
    </div>`;
  return card;
}

function slideBodyHtml(s) {
  if (s.stats && s.stats.length) {
    return `<div class="preview-stats">${s.stats.slice(0, 3).map(st =>
      `<div class="preview-stat"><strong>${escapeHtml(st.value)}</strong><span>${escapeHtml(st.label)}</span></div>`
    ).join('')}</div><p>${escapeHtml(s.content || '')}</p>`;
  }
  if (s.layout === 'two_column') {
    return `<div class="preview-columns">
      <div class="preview-column"><h3>${escapeHtml(s.leftTitle || 'Слева')}</h3><ul>${(s.leftContent || []).map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul></div>
      <div class="preview-column"><h3>${escapeHtml(s.rightTitle || 'Справа')}</h3><ul>${(s.rightContent || []).map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul></div>
    </div>`;
  }
  if (s.bullets && s.bullets.length) {
    return `<ul>${s.bullets.map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul>`;
  }
  const text = s.quote || s.content || s.subtitle || '';
  return `<p>${escapeHtml(text)}</p>`;
}

function renderPreview(index) {
  const slide = currentSlides[index];
  if (!slide) return;
  document.querySelectorAll('.slide-card').forEach((card, i) => {
    card.classList.toggle('active', i === index);
  });
  document.getElementById('preview-counter').textContent = `Слайд ${index + 1} из ${currentSlides.length}`;
  document.getElementById('preview-slide').innerHTML = `
    <h2>${escapeHtml(slide.title || `Слайд ${index + 1}`)}</h2>
    ${slideBodyHtml(slide)}
  `;
}

function collectEditedSlides() {
  return [...document.querySelectorAll('.slide-card')].map(card => {
    const idx = Number(card.dataset.index);
    const original = currentSlides[idx] || {};
    const body = card.querySelector('.edit-body').value.trim();
    const title = card.querySelector('.edit-title').value.trim();
    const slide = {...original, title};

    if (slide.layout === 'title') {
      slide.subtitle = body;
    } else if (slide.layout === 'quote') {
      slide.quote = body;
    } else if (slide.layout === 'conclusion' || slide.layout === 'section_break') {
      slide.content = body;
      slide.subtitle = body;
    } else if (slide.layout === 'stats') {
      slide.content = body;
    } else if (slide.layout === 'two_column') {
      const lines = body.split('\n').map(x => x.trim()).filter(Boolean);
      const half = Math.ceil(lines.length / 2);
      slide.leftContent = lines.slice(0, half);
      slide.rightContent = lines.slice(half);
    } else {
      slide.bullets = body.split('\n').map(x => x.trim()).filter(Boolean);
    }
    return slide;
  });
}

async function saveEdits() {
  if (!sessionId) return;
  const btn = document.getElementById('save-edits-btn');
  const slides = collectEditedSlides();
  btn.disabled = true;
  btn.textContent = 'Собираем PPTX...';
  try {
    const resp = await fetch(`/api/rebuild/${sessionId}`, {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        title: document.getElementById('result-title').textContent,
        style: currentStyle,
        slides
      })
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.detail || 'Ошибка пересборки');
    currentSlides = data.preview || slides;
    document.getElementById('download-btn').href = `/api/download/${sessionId}?t=${Date.now()}`;
    btn.textContent = 'PPTX обновлен';
    setTimeout(() => { btn.textContent = '↻ Обновить PPTX'; }, 1400);
  } catch (err) {
    alert(err.message);
    btn.textContent = '↻ Обновить PPTX';
  } finally {
    btn.disabled = false;
  }
}

function resetForm() {
  document.body.classList.remove('loading-mode');
  document.getElementById('result-section').style.display = 'none';
  document.getElementById('progress-section').style.display = 'none';
  document.getElementById('error-card').style.display = 'none';
  document.getElementById('gen-btn').disabled = false;
  document.getElementById('gen-btn').innerHTML = '<span>✦</span> Сгенерировать презентацию';
  setStep(1);
  // Reset progress steps
  ['ps1','ps2','ps3','ps4'].forEach(id => document.getElementById(id).classList.remove('done'));
  document.getElementById('progress-bar').style.width = '0%';
  document.getElementById('progress-percent').textContent = '0%';
}

// Sync radio buttons with select
document.getElementById('style').addEventListener('change', function() {
  const val = this.value;
  currentStyle = val;
  const radio = document.querySelector(`input[name="style-radio"][value="${val}"]`);
  if (radio) radio.checked = true;
});

document.getElementById('generator-form').addEventListener('submit', function(event) {
  const prompt = document.getElementById('prompt').value.trim();
  if (!prompt) {
    event.preventDefault();
    showError('Введите описание презентации');
  }
});

bindSlideCountControls();

</script>
</body>
</html>
"""

# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
