"""
OCR (Optical Character Recognition) через Tesseract.

Используется в двух случаях:
  1. PDF-страница является сканом (мало или нет текстового слоя)
  2. Изображение внутри PPTX/DOCX содержит текст

Требования:
  - Установленный Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
  - Русский языковой пакет (rus.traineddata) в папке tessdata
  - pytesseract + Pillow в requirements.txt
"""

import io

import pytesseract
from PIL import Image

from formats.config import OCR_LANGUAGE, OCR_FAILED_PLACEHOLDER, TESSERACT_CMD


# Указываем путь к Tesseract, если задан в конфиге
if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD


def ocr_image(image: Image.Image) -> str:
    """
    Распознаёт текст из PIL-изображения через Tesseract.

    Args:
        image: объект PIL.Image.

    Returns:
        Распознанный текст или плейсхолдер если OCR не дал результата.
    """
    try:
        # --oem 3: использовать LSTM-движок (лучшее качество)
        # --psm 6: считать изображение единым блоком текста
        config = "--oem 3 --psm 6"
        text = pytesseract.image_to_string(image, lang=OCR_LANGUAGE, config=config)
        text = text.strip()
        return text if text else OCR_FAILED_PLACEHOLDER
    except pytesseract.TesseractNotFoundError:
        raise RuntimeError(
            "Tesseract не найден. Установите Tesseract и укажите путь в config.py (TESSERACT_CMD)."
        )
    except Exception as e:
        # Не прерываем всю конвертацию из-за одного изображения
        return f"{OCR_FAILED_PLACEHOLDER} [{e}]"


def ocr_image_bytes(image_bytes: bytes) -> str:
    """
    Распознаёт текст из изображения переданного как bytes.
    Удобно для изображений, извлечённых из PPTX/DOCX.

    Args:
        image_bytes: содержимое изображения в байтах.

    Returns:
        Распознанный текст или плейсхолдер.
    """
    image = Image.open(io.BytesIO(image_bytes))
    # Конвертируем в RGB — Tesseract не работает с RGBA и палитровыми изображениями
    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")
    return ocr_image(image)


def ocr_pdf_page_image(page_image: Image.Image) -> str:
    """
    OCR для страницы PDF, рендеренной в изображение.
    Используется когда страница является сканом.

    Args:
        page_image: страница PDF как PIL.Image.

    Returns:
        Распознанный текст.
    """
    # Для сканов лучше работает psm 3 (полная автосегментация страницы)
    try:
        config = "--oem 3 --psm 3"
        text = pytesseract.image_to_string(page_image, lang=OCR_LANGUAGE, config=config)
        return text.strip()
    except Exception as e:
        return f"{OCR_FAILED_PLACEHOLDER} [{e}]"


def is_tesseract_available() -> bool:
    """
    Проверяет доступность Tesseract без выброса исключения.
    Используется при старте для информирования пользователя.
    """
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False
