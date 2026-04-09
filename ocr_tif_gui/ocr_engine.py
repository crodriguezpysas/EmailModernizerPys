#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Motor OCR: preprocesamiento de imagen, extracción con Tesseract y
expresiones regulares para los campos requeridos.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

from PIL import Image, ImageOps

try:
    import cv2
    import numpy as np
    _CV2_AVAILABLE = True
except ImportError:
    cv2 = None
    np = None
    _CV2_AVAILABLE = False

import pytesseract

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Expresiones regulares optimizadas
# ---------------------------------------------------------------------------

TURNO_REGEXES = [
    re.compile(
        r"\bturno\b[^0-9]{0,20}([0-9]{4}-\d{1,5}-\d{1,5}-\d{1,10})",
        re.IGNORECASE,
    ),
    re.compile(
        r"\bturno\b.*?([0-9]{4}(?:-\d+){3,5})",
        re.IGNORECASE | re.DOTALL,
    ),
]

# Dos patrones consolidados (eliminan la redundancia del script original)
MATRICULA_REGEXES = [
    re.compile(
        r"\b(?:Nro\.?|No\.?|N°)\s*(?:de\s*)?Mat\s*ricula\b\s*[:\-]?\s*([0-9]{1,6}-[0-9]{1,10})",
        re.IGNORECASE,
    ),
    re.compile(
        r"\bMatricula\b\s*[:\-]?\s*([0-9]{1,6}-[0-9]{1,10})",
        re.IGNORECASE,
    ),
]

# Límite de longitud 2–40 chars para evitar capturas excesivas
MUNICIPIO_REGEXES = [
    re.compile(
        r"\bMUNICIP[I1]?[O0]?[A@]?\b\s*[:\-\.]?\s*([A-Za-zÁÉÍÓÚáéíóúñÑ][A-Za-zÁÉÍÓÚáéíóúñÑ\s\-]{1,39}?)(?=\s*(?:DEPARTAMENTO|DPTO\b|$))",
        re.IGNORECASE,
    ),
    re.compile(
        r"\bMPIO\b\s*[:\-\.]?\s*([A-Za-zÁÉÍÓÚáéíóúñÑ][A-Za-zÁÉÍÓÚáéíóúñÑ\s\-]{1,39}?)(?=\s*(?:DEPARTAMENTO|DPTO\b|$))",
        re.IGNORECASE,
    ),
]

FECHA_REGEXES = [
    re.compile(
        r"\bF\s*E\s*C\s*H\s*A\b\s*[:\-\.]?\s*(\d{1,2}[\s/\-\.]\d{1,2}[\s/\-\.]\d{2,4})",
        re.IGNORECASE,
    ),
    re.compile(
        r"\bF\s*C\b\s*[:\-\.]?\s*(\d{1,2}[\s/\-\.]\d{1,2}[\s/\-\.]\d{2,4})",
        re.IGNORECASE,
    ),
]

RADICACION_REGEXES = [
    re.compile(
        r"\bRadicac[i1][oó0]n\b\s*[:\-\.]?\s*([0-9]{4}(?:-\d+){3,5})",
        re.IGNORECASE,
    ),
    re.compile(
        r"\bR\s*a\s*d\s*[i1]\s*c\s*[aá@]?\s*c\s*[i1]\s*[oó0]\s*n\b\s*[:\-\.]?\s*([0-9]{4}(?:-\d+){3,5})",
        re.IGNORECASE | re.DOTALL,
    ),
    re.compile(
        r"\bRad\.?\b\s*[:\-\.]?\s*([0-9]{4}(?:-\d+){3,5})",
        re.IGNORECASE,
    ),
]

# ---------------------------------------------------------------------------
# Funciones auxiliares
# ---------------------------------------------------------------------------

def normalize_text(s: str) -> str:
    """Elimina guiones suaves, colapsa espacios y saltos de línea."""
    s = s.replace("\u00ad", "")
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def extract_with_regexes(text: str, regexes) -> str:
    """Devuelve el primer grupo capturado que coincida con alguna regex."""
    for rx in regexes:
        m = rx.search(text)
        if m:
            return m.group(1).strip()
    return ""


# ---------------------------------------------------------------------------
# Preprocesamiento de imagen
# ---------------------------------------------------------------------------

def preprocess_image_pil(img: Image.Image) -> Image.Image:
    """Preprocesamiento básico con Pillow (fallback sin OpenCV)."""
    img = img.convert("L")
    img = ImageOps.autocontrast(img)
    w, h = img.size
    if max(w, h) < 2000:
        img = img.resize((w * 2, h * 2), Image.Resampling.LANCZOS)
    return img


def preprocess_image_cv(img: Image.Image) -> Image.Image:
    """
    Preprocesamiento avanzado con OpenCV (binarización adaptativa).
    Si OpenCV no está disponible recae en Pillow.
    """
    if not _CV2_AVAILABLE:
        return preprocess_image_pil(img)

    gray = np.array(img.convert("L"))
    gray = cv2.bilateralFilter(gray, 9, 75, 75)
    bin_img = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        35, 15,
    )
    kernel = np.ones((2, 2), np.uint8)
    bin_img = cv2.morphologyEx(bin_img, cv2.MORPH_CLOSE, kernel, iterations=1)

    pil_out = Image.fromarray(bin_img)
    w, h = pil_out.size
    if max(w, h) < 2000:
        pil_out = pil_out.resize((w * 2, h * 2), Image.Resampling.LANCZOS)
    return pil_out


# ---------------------------------------------------------------------------
# OCR con fallback de PSM
# ---------------------------------------------------------------------------

_MIN_TEXT_LEN = 20  # umbral para reintentar con --psm 3


def ocr_image(img: Image.Image, lang: str = "spa") -> str:
    """
    Ejecuta Tesseract sobre *img*.  Intenta primero con ``--psm 6``; si el
    resultado es muy corto (< 20 chars), reintenta con ``--psm 3``.
    """
    text = pytesseract.image_to_string(img, lang=lang, config="--oem 3 --psm 6")
    if len(text.strip()) < _MIN_TEXT_LEN:
        logger.debug("PSM 6 produjo texto muy corto, reintentando con PSM 3")
        text_retry = pytesseract.image_to_string(img, lang=lang, config="--oem 3 --psm 3")
        if len(text_retry.strip()) > len(text.strip()):
            text = text_retry
    return text


# ---------------------------------------------------------------------------
# Validación de Tesseract
# ---------------------------------------------------------------------------

def validate_tesseract(tesseract_cmd: str = "") -> None:
    """
    Verifica que Tesseract esté disponible antes de iniciar el procesamiento.

    Raises
    ------
    RuntimeError
        Si Tesseract no se encuentra o no responde correctamente.
    """
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    try:
        version = pytesseract.get_tesseract_version()
        logger.info("Tesseract disponible, versión: %s", version)
    except Exception as exc:
        raise RuntimeError(
            f"Tesseract no está disponible. Verifique la instalación o indique la ruta "
            f"correcta al ejecutable.\nDetalle: {exc}"
        ) from exc


# ---------------------------------------------------------------------------
# Filtrado de archivos
# ---------------------------------------------------------------------------

def iter_tifs(input_dir: Path):
    """
    Generador que recorre *input_dir* de forma recursiva y devuelve
    únicamente los archivos cuyo nombre termine en ``0001.tif`` o
    ``0001.tiff`` (insensible a mayúsculas).
    """
    for p in sorted(input_dir.rglob("*")):
        if p.is_file():
            lower = p.name.lower()
            if lower.endswith("0001.tif") or lower.endswith("0001.tiff"):
                yield p


# ---------------------------------------------------------------------------
# Procesamiento de un archivo .tif
# ---------------------------------------------------------------------------

def process_tif(path: Path, lang: str = "spa") -> dict:
    """
    Abre *path*, ejecuta OCR por página/frame y extrae los campos definidos.

    Returns
    -------
    dict con las claves:
        Turno, RUTA, Ruta Archivo,
        TURNO_OCR, MATRICULA_OCR, MUNICIPIO_OCR, FECHA_OCR, RADICACION_OCR
    """
    nombre_carpeta_padre = path.parent.name
    ruta_sin_archivo = str(path.parent)
    ruta_completa = str(path)

    full_text_parts: list[str] = []

    with Image.open(path) as im:
        n_frames = getattr(im, "n_frames", 1)
        for i in range(n_frames):
            try:
                im.seek(i)
            except Exception:
                break
            frame = im.copy()
            processed = preprocess_image_cv(frame)
            text = ocr_image(processed, lang=lang)
            if text:
                full_text_parts.append(text)
            # Liberar memoria explícitamente
            del frame
            del processed

    raw_text = "\n".join(full_text_parts)
    text = normalize_text(raw_text)

    return {
        "Turno": nombre_carpeta_padre,
        "RUTA": ruta_sin_archivo,
        "Ruta Archivo": ruta_completa,
        "TURNO_OCR": extract_with_regexes(text, TURNO_REGEXES),
        "MATRICULA_OCR": extract_with_regexes(text, MATRICULA_REGEXES),
        "MUNICIPIO_OCR": extract_with_regexes(text, MUNICIPIO_REGEXES),
        "FECHA_OCR": extract_with_regexes(text, FECHA_REGEXES),
        "RADICACION_OCR": extract_with_regexes(text, RADICACION_REGEXES),
    }
