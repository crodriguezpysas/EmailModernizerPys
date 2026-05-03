# ocr_tif_gui – Extractor OCR de campos desde archivos .tif/.tiff

Aplicación de escritorio con interfaz gráfica (tkinter) para procesar imágenes
escaneadas en formato TIFF y extraer campos estructurados mediante OCR.

## Requisitos del sistema

| Dependencia | Versión mínima | Notas |
|---|---|---|
| Python | 3.8 | Incluye tkinter |
| Tesseract OCR | 4.x | Debe estar en el PATH o indicar ruta |
| pytesseract | 0.3.10 | Wrapper Python para Tesseract |
| Pillow | 10.0 | Procesamiento de imágenes |
| opencv-python-headless | 4.8 | Opcional – mejora el preprocesamiento |
| numpy | 1.24 | Requerido por OpenCV |

### Instalar Tesseract

- **Windows**: https://github.com/UB-Mannheim/tesseract/wiki  
  Añadir el ejecutable al PATH o indicar la ruta en la GUI.
- **Linux**: `sudo apt-get install tesseract-ocr tesseract-ocr-spa`
- **macOS**: `brew install tesseract`

### Instalar dependencias Python

```bash
pip install -r ocr_tif_gui/requirements.txt
```

## Estructura de archivos

```
ocr_tif_gui/
├── __init__.py       # Metadatos del paquete
├── main.py           # Punto de entrada
├── ocr_engine.py     # Motor OCR (preprocesamiento, Tesseract, regex)
├── gui_app.py        # Interfaz gráfica tkinter
├── requirements.txt  # Dependencias
└── README.md         # Este archivo
```

## Ejecución

```bash
# Desde la raíz del repositorio:
python -m ocr_tif_gui

# O directamente:
python ocr_tif_gui/main.py
```

## Uso de la interfaz

1. **Seleccionar carpeta raíz** – Abre un diálogo para elegir la carpeta que
   contiene las subcarpetas con los archivos TIFF.
2. **Iniciar procesamiento** – Recorre recursivamente la carpeta buscando
   archivos cuyo nombre termine en `0001.tif` o `0001.tiff` (sin distinción
   de mayúsculas/minúsculas) y ejecuta el OCR sobre cada uno.
3. **Tabla de resultados** – Las filas se agregan en tiempo real con los
   campos extraídos: Turno, RUTA, Ruta Archivo, TURNO_OCR, MATRICULA_OCR,
   MUNICIPIO_OCR, FECHA_OCR, RADICACION_OCR.
4. **Detener** – Cancela el procesamiento en cualquier momento.
5. **Exportar CSV** – Guarda todos los resultados en un archivo CSV codificado
   en UTF-8 con BOM (compatible con Excel).

## Campos extraídos

| Campo | Descripción |
|---|---|
| Turno | Nombre de la carpeta contenedora del archivo |
| RUTA | Ruta del directorio padre |
| Ruta Archivo | Ruta completa al archivo .tif |
| TURNO_OCR | Número de turno extraído del texto OCR |
| MATRICULA_OCR | Número de matrícula extraído del texto OCR |
| MUNICIPIO_OCR | Municipio extraído del texto OCR |
| FECHA_OCR | Fecha extraída del texto OCR |
| RADICACION_OCR | Número de radicación extraído del texto OCR |

## Optimizaciones respecto al script original

- **OCR con fallback PSM**: intenta primero `--psm 6`; si el texto es muy
  corto (< 20 caracteres), reintenta con `--psm 3`.
- **Regex de MATRÍCULA consolidados**: 3 patrones redundantes reducidos a 2
  más precisos.
- **Regex de MUNICIPIO mejorado**: límite de longitud 2–40 caracteres para
  evitar capturas excesivas.
- **Validación temprana de Tesseract**: se verifica la disponibilidad del
  motor antes de iniciar el procesamiento.
- **Logging estructurado**: se usa el módulo estándar `logging`.
- **Gestión de memoria mejorada**: los frames procesados se liberan
  explícitamente con `del`.
- **GUI no bloqueante**: el OCR corre en un hilo separado via `threading`;
  los resultados se comunican a través de `queue.Queue` y `after()`.
