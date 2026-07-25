# -*- coding: utf-8 -*-
"""
doc2md — веб-интерфейс конвертера документов в Markdown.

Запуск:
    streamlit run app.py

Возможности:
  - выбор нескольких файлов (.pptx, .docx, .pdf) или перетаскивание их
    в рабочую область (drag & drop встроен в загрузчик Streamlit)
  - конвертация каждого файла через единый реестр formats.CONVERTERS
  - предпросмотр результата, предупреждения конвертации
  - скачивание .md по одному или всех сразу ZIP-архивом
"""

import io
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

# Гарантируем видимость пакета formats независимо от того,
# из какой папки запущен «streamlit run app.py»
sys.path.insert(0, str(Path(__file__).resolve().parent))

from formats import SUPPORTED_EXTENSIONS, convert_file
from formats.utils.ocr import is_tesseract_available

st.set_page_config(
    page_title="doc2md — документы в Markdown",
    page_icon="📄",
    layout="wide",
)

# ─── Сайдбар: справка и статус окружения ─────────────────────────────────────

with st.sidebar:
    st.header("О конвертере")
    st.markdown(
        "Извлекает текст из документов и сохраняет его в **Markdown**:\n"
        + "\n".join(f"- `{ext}`" for ext in sorted(SUPPORTED_EXTENSIONS))
    )
    st.divider()
    if is_tesseract_available():
        st.success("OCR (Tesseract): доступен")
        st.caption("Сканы PDF и изображения будут распознаны.")
    else:
        st.warning("OCR (Tesseract): недоступен")
        st.caption(
            "Сканы и изображения будут пропущены. "
            "Установка: github.com/UB-Mannheim/tesseract"
        )

# ─── Основная область ────────────────────────────────────────────────────────

st.title("📄 doc2md — конвертер документов в Markdown")
st.caption("Перетащите файлы в область ниже или выберите их кнопкой Browse files.")

uploaded_files = st.file_uploader(
    "Файлы для конвертации",
    type=[ext.lstrip(".") for ext in SUPPORTED_EXTENSIONS],
    accept_multiple_files=True,
    help="Можно выбрать сразу несколько файлов или перетащить их сюда мышью.",
)

# Результаты храним в session_state: нажатие кнопки скачивания перезапускает
# скрипт, и без этого конвертация выполнялась бы заново на каждый клик.
if "results" not in st.session_state:
    st.session_state.results = []


def _convert_uploaded(files) -> list[dict]:
    """Конвертирует загруженные файлы, возвращает список результатов."""
    results: list[dict] = []
    progress = st.progress(0.0, text="Конвертация...")

    # Загруженный файл существует только в памяти, а конвертеры работают
    # с путями на диске — поэтому пишем во временную папку
    with tempfile.TemporaryDirectory(prefix="md_convert_") as tmp_dir:
        for i, uf in enumerate(files, start=1):
            progress.progress(i / len(files), text=f"[{i}/{len(files)}] {uf.name}")

            src = Path(tmp_dir) / uf.name
            src.write_bytes(uf.getbuffer())

            item = {
                "name": uf.name,
                "md_name": Path(uf.name).stem + ".md",
                "content": "",
                "warnings": [],
                "error": None,
            }
            try:
                conv = convert_file(src)
                item["content"] = conv.content
                item["warnings"] = conv.warnings
            except Exception as e:
                item["error"] = str(e)
            results.append(item)

    progress.empty()
    return results


if uploaded_files:
    st.write(f"Выбрано файлов: **{len(uploaded_files)}**")
    if st.button("🚀 Конвертировать", type="primary"):
        st.session_state.results = _convert_uploaded(uploaded_files)
else:
    # Файлы убрали из загрузчика — сбрасываем старые результаты
    st.session_state.results = []

results = st.session_state.results

# ─── Вывод результатов ───────────────────────────────────────────────────────

if results:
    ok = [r for r in results if not r["error"] and not r["warnings"]]
    warn = [r for r in results if not r["error"] and r["warnings"]]
    err = [r for r in results if r["error"]]

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("✅ Успешно", len(ok))
    c2.metric("⚠️ С предупреждениями", len(warn))
    c3.metric("❌ С ошибками", len(err))

    # ZIP со всеми успешными результатами
    converted = [r for r in results if not r["error"]]
    if converted:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            used_names: set[str] = set()
            for r in converted:
                # При совпадении имён (одинаковые stem у разных файлов)
                # добавляем числовой суффикс, чтобы не потерять файлы в архиве
                name = r["md_name"]
                n = 1
                while name in used_names:
                    n += 1
                    name = f"{Path(r['md_name']).stem}_{n}.md"
                used_names.add(name)
                zf.writestr(name, r["content"])

        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            f"📦 Скачать всё ({len(converted)} .md) — ZIP",
            data=zip_buf.getvalue(),
            file_name=f"markdown_{stamp}.zip",
            mime="application/zip",
        )

    # Карточка результата на каждый файл
    for i, r in enumerate(results):
        if r["error"]:
            icon = "❌"
        elif r["warnings"]:
            icon = "⚠️"
        else:
            icon = "✅"

        with st.expander(f"{icon} {r['name']}", expanded=bool(r["error"])):
            if r["error"]:
                st.error(f"Ошибка конвертации: {r['error']}")
                continue

            for w in r["warnings"]:
                st.warning(w)

            st.download_button(
                f"⬇️ Скачать {r['md_name']}",
                data=r["content"],
                file_name=r["md_name"],
                mime="text/markdown",
                key=f"dl_{i}",
            )

            tab_view, tab_raw = st.tabs(["Просмотр", "Markdown-код"])
            with tab_view:
                st.markdown(r["content"])
            with tab_raw:
                st.code(r["content"], language="markdown")
