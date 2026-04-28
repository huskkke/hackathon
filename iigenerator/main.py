"""Compatibility entrypoint for PresentAI.

Run `python iigenerator/app.py` or `python -m iigenerator.main`.
"""

from .app import app


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
