#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Punto de entrada de la aplicación OCR TIF GUI.

Uso
---
    python -m ocr_tif_gui
    # o directamente:
    python ocr_tif_gui/main.py
"""

import logging

from .gui_app import OcrApp


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(name)s  %(message)s",
    )
    app = OcrApp()
    app.mainloop()


if __name__ == "__main__":
    main()
