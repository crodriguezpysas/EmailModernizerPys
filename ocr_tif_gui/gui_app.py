#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Interfaz gráfica principal de la aplicación OCR TIF.

Arquitectura de hilos
---------------------
- Hilo principal (tkinter): actualiza la GUI usando `after()` para consumir
  mensajes de la cola de resultados.
- Hilo de trabajo (threading.Thread): recorre archivos, llama al motor OCR y
  envía resultados a la cola.
"""

from __future__ import annotations

import csv
import logging
import queue
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .ocr_engine import iter_tifs, process_tif, validate_tesseract

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

COLUMNS = [
    "Turno",
    "RUTA",
    "Ruta Archivo",
    "TURNO_OCR",
    "MATRICULA_OCR",
    "MUNICIPIO_OCR",
    "FECHA_OCR",
    "RADICACION_OCR",
]

# Intervalo en ms para comprobar la cola de resultados
_POLL_INTERVAL_MS = 100


# ---------------------------------------------------------------------------
# Clase principal de la aplicación
# ---------------------------------------------------------------------------

class OcrApp(tk.Tk):
    """Ventana principal de la aplicación OCR TIF."""

    def __init__(self):
        super().__init__()
        self.title("OCR TIF – Extractor de campos")
        self.geometry("1200x750")
        self.minsize(900, 600)
        self.configure(bg="#f0f0f0")

        # Estado interno
        self._result_queue: queue.Queue = queue.Queue()
        self._stop_event = threading.Event()
        self._worker_thread: threading.Thread | None = None
        self._rows: list[dict] = []
        self._total_files = 0
        self._processed_files = 0

        self._build_ui()
        self._configure_logging()

    # ------------------------------------------------------------------
    # Construcción de la interfaz
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        """Construye todos los widgets de la ventana."""
        # ── Barra superior (controles) ──────────────────────────────────
        top_frame = ttk.LabelFrame(self, text="Configuración", padding=8)
        top_frame.pack(fill=tk.X, padx=10, pady=(10, 4))

        ttk.Label(top_frame, text="Carpeta raíz:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 4)
        )
        self._folder_var = tk.StringVar()
        self._folder_entry = ttk.Entry(
            top_frame, textvariable=self._folder_var, width=70
        )
        self._folder_entry.grid(row=0, column=1, sticky=tk.EW, padx=4)
        ttk.Button(
            top_frame, text="Seleccionar carpeta raíz", command=self._select_folder
        ).grid(row=0, column=2, padx=4)

        top_frame.columnconfigure(1, weight=1)

        btn_frame = ttk.Frame(top_frame)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=(8, 0), sticky=tk.W)

        self._btn_start = ttk.Button(
            btn_frame, text="▶ Iniciar procesamiento", command=self._start_processing
        )
        self._btn_start.pack(side=tk.LEFT, padx=(0, 6))

        self._btn_stop = ttk.Button(
            btn_frame,
            text="⏹ Detener",
            command=self._stop_processing,
            state=tk.DISABLED,
        )
        self._btn_stop.pack(side=tk.LEFT, padx=6)

        self._btn_export = ttk.Button(
            btn_frame,
            text="💾 Exportar CSV",
            command=self._export_csv,
            state=tk.DISABLED,
        )
        self._btn_export.pack(side=tk.LEFT, padx=6)

        # ── Barra de progreso ───────────────────────────────────────────
        prog_frame = ttk.Frame(self)
        prog_frame.pack(fill=tk.X, padx=10, pady=4)

        self._progress_label = ttk.Label(
            prog_frame, text="Listo para procesar."
        )
        self._progress_label.pack(anchor=tk.W)

        self._progress_bar = ttk.Progressbar(
            prog_frame, orient=tk.HORIZONTAL, mode="determinate", length=400
        )
        self._progress_bar.pack(fill=tk.X, pady=(2, 0))

        # ── Tabla de resultados ─────────────────────────────────────────
        table_frame = ttk.LabelFrame(self, text="Resultados", padding=4)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        self._tree = ttk.Treeview(
            table_frame, columns=COLUMNS, show="headings", selectmode="browse"
        )
        for col in COLUMNS:
            self._tree.heading(col, text=col)
            self._tree.column(col, width=140, minwidth=80, stretch=True)

        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self._tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self._tree.pack(fill=tk.BOTH, expand=True)

        # ── Panel de log ────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self, text="Log en tiempo real", padding=4)
        log_frame.pack(fill=tk.X, padx=10, pady=(4, 10))

        self._log_text = tk.Text(
            log_frame,
            height=8,
            state=tk.DISABLED,
            wrap=tk.WORD,
            bg="#1e1e1e",
            fg="#d4d4d4",
            font=("Consolas", 9),
            relief=tk.FLAT,
        )
        log_vsb = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=log_vsb.set)
        log_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._log_text.pack(fill=tk.X, expand=True)

    # ------------------------------------------------------------------
    # Logging hacia el panel de texto
    # ------------------------------------------------------------------

    def _configure_logging(self) -> None:
        """Redirige los mensajes de logging al panel de log de la GUI."""
        handler = _TextHandler(self._log_text)
        handler.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s"))
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)
        root_logger.setLevel(logging.DEBUG)

    def _append_log(self, msg: str) -> None:
        """Agrega una línea de texto al panel de log (hilo principal)."""
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self._log_text.configure(state=tk.NORMAL)
        self._log_text.insert(tk.END, line)
        self._log_text.see(tk.END)
        self._log_text.configure(state=tk.DISABLED)

    # ------------------------------------------------------------------
    # Acciones de los botones
    # ------------------------------------------------------------------

    def _select_folder(self) -> None:
        folder = filedialog.askdirectory(title="Seleccionar carpeta raíz")
        if folder:
            self._folder_var.set(folder)
            self._append_log(f"Carpeta seleccionada: {folder}")

    def _start_processing(self) -> None:
        folder = self._folder_var.get().strip()
        if not folder:
            messagebox.showwarning("Sin carpeta", "Seleccione una carpeta raíz primero.")
            return

        input_path = Path(folder)
        if not input_path.is_dir():
            messagebox.showerror("Carpeta inválida", f"La ruta no existe o no es una carpeta:\n{folder}")
            return

        # Limpiar estado anterior
        self._rows.clear()
        for item in self._tree.get_children():
            self._tree.delete(item)
        self._log_text.configure(state=tk.NORMAL)
        self._log_text.delete("1.0", tk.END)
        self._log_text.configure(state=tk.DISABLED)

        self._stop_event.clear()
        self._processed_files = 0
        self._total_files = 0
        self._progress_bar["value"] = 0
        self._progress_label.configure(text="Preparando…")

        self._btn_start.configure(state=tk.DISABLED)
        self._btn_stop.configure(state=tk.NORMAL)
        self._btn_export.configure(state=tk.DISABLED)

        self._worker_thread = threading.Thread(
            target=self._worker,
            args=(input_path,),
            daemon=True,
        )
        self._worker_thread.start()
        self.after(_POLL_INTERVAL_MS, self._poll_queue)

    def _stop_processing(self) -> None:
        self._stop_event.set()
        self._append_log("⏹ Deteniendo procesamiento…")

    def _export_csv(self) -> None:
        if not self._rows:
            messagebox.showinfo("Sin datos", "No hay resultados para exportar.")
            return
        path = filedialog.asksaveasfilename(
            title="Guardar CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=COLUMNS)
                writer.writeheader()
                writer.writerows(self._rows)
            messagebox.showinfo("Exportado", f"CSV guardado en:\n{path}")
            self._append_log(f"CSV exportado: {path}")
        except Exception as exc:
            messagebox.showerror("Error al exportar", str(exc))

    # ------------------------------------------------------------------
    # Hilo de trabajo
    # ------------------------------------------------------------------

    def _worker(self, input_dir: Path) -> None:
        """
        Se ejecuta en un hilo separado.  Envía mensajes a la cola para
        que el hilo principal actualice la GUI.
        """
        start_time = time.time()

        # Validar Tesseract antes de comenzar
        try:
            validate_tesseract()
        except RuntimeError as exc:
            self._result_queue.put(("error", str(exc)))
            self._result_queue.put(("done", None))
            return

        # Recopilar archivos para poder calcular el total
        self._result_queue.put(("log", "Buscando archivos…"))
        tif_files = list(iter_tifs(input_dir))
        total = len(tif_files)
        self._result_queue.put(("total", total))

        if total == 0:
            self._result_queue.put(("log", "No se encontraron archivos terminados en 0001.tif/.tiff"))
            self._result_queue.put(("done", None))
            return

        self._result_queue.put(("log", f"Total de archivos a procesar: {total}"))

        processed_count = 0
        for idx, tif_path in enumerate(tif_files, start=1):
            if self._stop_event.is_set():
                self._result_queue.put(("log", "Procesamiento detenido por el usuario."))
                break

            self._result_queue.put(("log", f"Procesando carpeta: {tif_path.parent}"))
            self._result_queue.put(("log", f"Archivo encontrado: {tif_path.name}"))

            try:
                row = process_tif(tif_path)
                self._result_queue.put(("row", row))
                self._result_queue.put(("progress", idx))
            except Exception as exc:
                logger.error("Error en archivo %s: %s", tif_path, exc)
                self._result_queue.put(("log", f"Error en archivo {tif_path.name}: {exc}"))
                # Añadir fila con campos vacíos para registrar el intento
                row = {
                    "Turno": tif_path.parent.name,
                    "RUTA": str(tif_path.parent),
                    "Ruta Archivo": str(tif_path),
                    "TURNO_OCR": "",
                    "MATRICULA_OCR": "",
                    "MUNICIPIO_OCR": "",
                    "FECHA_OCR": "",
                    "RADICACION_OCR": "",
                }
                self._result_queue.put(("row", row))
                self._result_queue.put(("progress", idx))
            processed_count = idx

        elapsed = time.time() - start_time
        count = total if not self._stop_event.is_set() else processed_count
        self._result_queue.put(
            ("log", f"Procesamiento completado: {count} archivos en {elapsed:.1f} segundos")
        )
        self._result_queue.put(("done", None))

    # ------------------------------------------------------------------
    # Consumidor de la cola (hilo principal)
    # ------------------------------------------------------------------

    def _poll_queue(self) -> None:
        """Lee todos los mensajes disponibles en la cola y actualiza la GUI."""
        try:
            while True:
                kind, payload = self._result_queue.get_nowait()
                self._handle_message(kind, payload)
        except queue.Empty:
            pass

        # Seguir sondeando si el hilo de trabajo sigue vivo
        if self._worker_thread and self._worker_thread.is_alive():
            self.after(_POLL_INTERVAL_MS, self._poll_queue)

    def _handle_message(self, kind: str, payload) -> None:
        """Despacha un mensaje de la cola al widget correspondiente."""
        if kind == "log":
            self._append_log(payload)
        elif kind == "total":
            self._total_files = payload
            self._progress_bar["maximum"] = max(payload, 1)
            self._progress_label.configure(text=f"Procesando 0 de {payload} archivos")
        elif kind == "progress":
            self._processed_files = payload
            self._progress_bar["value"] = payload
            pct = int(payload / max(self._total_files, 1) * 100)
            self._progress_label.configure(
                text=f"Procesando {payload} de {self._total_files} archivos  ({pct} %)"
            )
        elif kind == "row":
            self._rows.append(payload)
            values = [payload.get(col, "") for col in COLUMNS]
            self._tree.insert("", tk.END, values=values)
            self._tree.yview_moveto(1)
        elif kind == "error":
            messagebox.showerror("Error crítico", payload)
        elif kind == "done":
            self._progress_bar["value"] = self._total_files
            self._progress_label.configure(
                text=f"Completado: {self._processed_files} de {self._total_files} archivos procesados"
            )
            self._btn_start.configure(state=tk.NORMAL)
            self._btn_stop.configure(state=tk.DISABLED)
            if self._rows:
                self._btn_export.configure(state=tk.NORMAL)


# ---------------------------------------------------------------------------
# Handler de logging para el widget Text
# ---------------------------------------------------------------------------

class _TextHandler(logging.Handler):
    """Envía registros de logging al widget Text de la GUI."""

    def __init__(self, text_widget: tk.Text):
        super().__init__()
        self._widget = text_widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        try:
            self._widget.configure(state=tk.NORMAL)
            self._widget.insert(tk.END, msg + "\n")
            self._widget.see(tk.END)
            self._widget.configure(state=tk.DISABLED)
        except tk.TclError:
            # Widget ya destruido (ventana cerrada)
            pass
