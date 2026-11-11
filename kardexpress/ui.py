"""Interfaz gráfica de Kardexpress v2 (customtkinter)."""
from __future__ import annotations

import logging
import os
import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Callable, Optional

import customtkinter as ctk
from PIL import Image

from . import APP_NAME, __version__
from .core import (Cancelled, KardexError, KardexResult, guess_file_role, process_kardex, scan_kardex,
                   scan_requisitions)
from .report import STATES, build_report

STATUS_COLORS = {k: s.strong for k, s in STATES.items()}
from .settings import Settings, log_file, resource_path

log = logging.getLogger(__name__)

try:  # arrastrar y soltar es opcional
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND = True
except Exception:  # noqa: BLE001
    TkinterDnD = None
    _DND = False

# ----------------------------------------------------------------------------- tema
FONT = "Segoe UI"
ACCENT = ("#0A6FA8", "#2AA9E0")          # azules Hidroobras
ACCENT_HOVER = ("#085A88", "#1E93C8")
ACCENT_SOFT = ("#E1F1FA", "#12324A")     # fondo de las insignias de paso
BG = ("#F3F4F6", "#0B1220")
CARD = ("#FFFFFF", "#161F30")
CARD_BORDER = ("#E5E7EB", "#243049")
TEXT = ("#111827", "#F3F4F6")
MUTED = ("#6B7280", "#9CA3AF")
SIDEBAR = "#0B1F33"
OK = ("#15803D", "#4ADE80")
WARN = ("#B45309", "#FBBF24")
ERR = ("#B91C1C", "#F87171")

STATE_ICON = {"empty": "○", "loading": "◌", "ok": "✓", "warn": "!", "error": "✕"}
STATE_COLOR = {"empty": MUTED, "loading": MUTED, "ok": OK, "warn": WARN, "error": ERR}


def _fmt(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def _run_bg(fn: Callable, on_done: Callable, widget: tk.Misc) -> None:
    """Ejecuta fn() en un hilo y entrega (resultado, error) a on_done en el hilo de la UI."""
    def worker():
        try:
            res, err = fn(), None
        except Exception as exc:  # noqa: BLE001
            res, err = None, exc
        try:
            widget.after(0, lambda: on_done(res, err))
        except RuntimeError:  # la ventana ya se cerró
            pass
    threading.Thread(target=worker, daemon=True).start()


# ----------------------------------------------------------------------------- tarjeta de archivo
class FileCard(ctk.CTkFrame):
    def __init__(self, master, step: int, title: str, hint: str, on_pick: Callable[[], None],
                 on_clear: Callable[[], None]):
        super().__init__(master, fg_color=CARD, corner_radius=14, border_width=1, border_color=CARD_BORDER)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(0, minsize=82)

        self.badge = ctk.CTkLabel(self, text=str(step), width=34, height=34, corner_radius=17,
                                  fg_color=ACCENT_SOFT, text_color=ACCENT,
                                  font=(FONT, 15, "bold"))
        self.badge.grid(row=0, column=0, rowspan=2, padx=(18, 14), pady=18, sticky="n")

        ctk.CTkLabel(self, text=title, font=(FONT, 15, "bold"), text_color=TEXT, anchor="w") \
            .grid(row=0, column=1, sticky="w", pady=(16, 0))
        self.detail = ctk.CTkLabel(self, text=hint, font=(FONT, 12), text_color=MUTED, anchor="w",
                                   justify="left", wraplength=420)
        self.detail.grid(row=1, column=1, sticky="w", pady=(0, 16))

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.grid(row=0, column=2, rowspan=2, padx=16)
        self.pick_btn = ctk.CTkButton(btns, text="Seleccionar…", width=118, height=34, corner_radius=8,
                                      font=(FONT, 13), command=on_pick, fg_color=("#E5E7EB", "#243049"),
                                      hover_color=("#D1D5DB", "#2F3C59"), text_color=TEXT)
        self.pick_btn.pack(side="left")
        self.clear_btn = ctk.CTkButton(btns, text="✕", width=34, height=34, corner_radius=8,
                                       command=on_clear, fg_color="transparent",
                                       hover_color=("#E5E7EB", "#243049"), text_color=MUTED)
        self.action_btn = ctk.CTkButton(self, text="", height=28, corner_radius=8, font=(FONT, 12),
                                        fg_color="transparent", border_width=1, border_color=WARN,
                                        text_color=WARN, hover_color=("#FEF3C7", "#3B2F12"))
        self._hint = hint
        self.path: Optional[Path] = None
        self.state = "empty"

        for w in (self, self.detail):
            w.bind("<Button-1>", lambda _e: on_pick())
            w.configure(cursor="hand2")

    def set_state(self, state: str, title: Optional[str] = None, detail: Optional[str] = None,
                  action: Optional[tuple[str, Callable[[], None]]] = None) -> None:
        self.state = state
        if state == "empty":
            self.badge.configure(fg_color=ACCENT_SOFT, text_color=ACCENT)
            self.detail.configure(text=self._hint, text_color=MUTED)
            self.configure(border_color=CARD_BORDER)
            self.clear_btn.pack_forget()
        else:
            color = STATE_COLOR[state]
            self.badge.configure(text=STATE_ICON[state], fg_color=CARD, text_color=color)
            text = title or ""
            if detail:
                text = f"{text}\n{detail}" if text else detail
            self.detail.configure(text=text, text_color=TEXT if state in ("ok", "loading") else color)
            self.configure(border_color=color if state in ("ok", "warn", "error") else CARD_BORDER)
            self.clear_btn.pack(side="left", padx=(6, 0))
        if action:
            self.action_btn.configure(text=action[0], command=action[1])
            self.action_btn.grid(row=2, column=1, sticky="w", pady=(0, 14))
        else:
            self.action_btn.grid_forget()

    def reset_badge_number(self, step: int) -> None:
        self.badge.configure(text=str(step))


# ----------------------------------------------------------------------------- tarjeta de estadística
class StatTile(ctk.CTkFrame):
    def __init__(self, master, label: str, swatch: Optional[str] = None):
        super().__init__(master, fg_color=CARD, corner_radius=12, border_width=1, border_color=CARD_BORDER)
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(12, 0))
        if swatch:
            ctk.CTkLabel(top, text="", width=10, height=10, corner_radius=3, fg_color=f"#{swatch}") \
                .pack(side="left", padx=(0, 6))
        ctk.CTkLabel(top, text=label, font=(FONT, 11), text_color=MUTED, anchor="w").pack(side="left")
        self.value = ctk.CTkLabel(self, text="—", font=(FONT, 24, "bold"), text_color=TEXT, anchor="w")
        self.value.pack(fill="x", padx=14, pady=(0, 10))

    def set(self, text: str) -> None:
        self.value.configure(text=text)


# ----------------------------------------------------------------------------- ventana principal
_Base = (type("DnDCTk", (ctk.CTk, TkinterDnD.DnDWrapper), {}) if _DND else ctk.CTk)


class App(_Base):
    def __init__(self):
        super().__init__()
        self.dnd = False
        if _DND:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
                self.dnd = True
            except Exception:  # noqa: BLE001
                log.warning("tkdnd no disponible; se desactiva arrastrar y soltar", exc_info=True)

        self.settings = Settings.load()
        ctk.set_appearance_mode(self.settings.appearance)

        self.title(f"{APP_NAME} {__version__}")
        self.geometry("1120x760")
        self.minsize(980, 680)
        self.configure(fg_color=BG)
        try:
            self.iconbitmap(str(resource_path("assets", "kardexpress.ico")))
        except Exception:  # noqa: BLE001
            pass

        self.req_path: Optional[Path] = None
        self.kx_path: Optional[Path] = None
        self.req_ok = self.kx_ok = False
        self.worker_q: "queue.Queue[tuple]" = queue.Queue()
        self.cancel_event: Optional[threading.Event] = None
        self.running = False
        self.last_result: Optional[KardexResult] = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()

        self.bind("<Return>", lambda _e: self.start())
        self.bind("<Escape>", lambda _e: self.cancel())
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Recuperar últimos archivos usados (solo si siguen existiendo)
        if self.settings.last_requisitions and Path(self.settings.last_requisitions).exists():
            self.set_file("req", Path(self.settings.last_requisitions))
        if self.settings.last_kardex and Path(self.settings.last_kardex).exists():
            self.set_file("kx", Path(self.settings.last_kardex))
        self._refresh_run_button()

    # ------------------------------------------------------------------ layout
    def _build_sidebar(self):
        side = ctk.CTkFrame(self, width=300, corner_radius=0, fg_color=SIDEBAR)
        side.grid(row=0, column=0, sticky="nsw")
        side.grid_propagate(False)
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(5, weight=1)

        head = ctk.CTkFrame(side, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=24, pady=(28, 6))
        try:
            logo = ctk.CTkImage(Image.open(resource_path("assets", "logo.png")), size=(56, 56))
            ctk.CTkLabel(head, image=logo, text="").pack(side="left")
        except Exception:  # noqa: BLE001
            pass
        title = ctk.CTkFrame(head, fg_color="transparent")
        title.pack(side="left", padx=12)
        ctk.CTkLabel(title, text=APP_NAME, font=(FONT, 22, "bold"), text_color="#F9FAFB").pack(anchor="w")
        ctk.CTkLabel(title, text=f"Hidroobras  ·  v{__version__}", font=(FONT, 11),
                     text_color="#94A3B8").pack(anchor="w")

        ctk.CTkLabel(side, text="Cruza el reporte de requisiciones con el kardex de compras: "
                                "asigna la fecha de cada requisición y marca el estado de cada línea.",
                     font=(FONT, 12), text_color="#CBD5E1", wraplength=250, justify="left") \
            .grid(row=1, column=0, sticky="w", padx=24, pady=(10, 18))

        ctk.CTkLabel(side, text="SEMÁFORO", font=(FONT, 11, "bold"), text_color="#64748B") \
            .grid(row=2, column=0, sticky="w", padx=24)
        legend = ctk.CTkFrame(side, fg_color="transparent")
        legend.grid(row=3, column=0, sticky="ew", padx=24, pady=(6, 0))
        descr = {"sin_oc": "Cantidad comprada = 0",
                 "sin_entrada": "Comprado, pero entrada = 0",
                 "pendiente": "Con entrada y saldo pendiente"}
        titles = {"sin_oc": "Sin orden de compra", "sin_entrada": "Sin entrada/remisión",
                  "pendiente": "Pendiente proveedor"}
        for key in ("sin_oc", "sin_entrada", "pendiente"):
            row = ctk.CTkFrame(legend, fg_color="transparent")
            row.pack(fill="x", pady=4)
            ctk.CTkLabel(row, text="", width=14, height=14, corner_radius=4,
                         fg_color=f"#{STATUS_COLORS[key]}").pack(side="left", padx=(0, 10), anchor="n", pady=3)
            txt = ctk.CTkFrame(row, fg_color="transparent")
            txt.pack(side="left", fill="x")
            ctk.CTkLabel(txt, text=titles[key], font=(FONT, 12, "bold"),
                         text_color="#E2E8F0", anchor="w").pack(anchor="w")
            ctk.CTkLabel(txt, text=descr[key], font=(FONT, 11), text_color="#94A3B8", anchor="w").pack(anchor="w")


        foot = ctk.CTkFrame(side, fg_color="transparent")
        foot.grid(row=6, column=0, sticky="ew", padx=24, pady=20)
        self.theme = ctk.CTkSegmentedButton(foot, values=["Claro", "Oscuro", "Sistema"], height=28,
                                            font=(FONT, 11), command=self._on_theme,
                                            selected_color=("#1E4A70", "#1E4A70"),
                                            unselected_color="#132E4A", fg_color="#132E4A",
                                            text_color="#E2E8F0")
        self.theme.set({"light": "Claro", "dark": "Oscuro"}.get(self.settings.appearance, "Sistema"))
        self.theme.pack(fill="x")
        ctk.CTkButton(foot, text="Ver registro de actividad", height=26, font=(FONT, 11),
                      fg_color="transparent", hover_color="#132E4A", text_color="#94A3B8",
                      command=lambda: self._open(log_file())).pack(fill="x", pady=(8, 0))

    def _build_main(self):
        main = ctk.CTkScrollableFrame(self, fg_color=BG, corner_radius=0)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        self.main = main
        pad = {"padx": 36}

        ctk.CTkLabel(main, text="Generar kardex", font=(FONT, 26, "bold"), text_color=TEXT, anchor="w") \
            .grid(row=0, column=0, sticky="w", pady=(30, 2), **pad)
        drop = " o arrástralos a esta ventana" if self.dnd else ""
        ctk.CTkLabel(main, text=f"Selecciona los dos reportes de Excel (.xlsx){drop}.",
                     font=(FONT, 13), text_color=MUTED, anchor="w") \
            .grid(row=1, column=0, sticky="w", pady=(0, 18), **pad)

        self.req_card = FileCard(main, 1, "Reporte de requisiciones",
                                 "Contiene los códigos REQUI-… y su fecha de creación.",
                                 on_pick=lambda: self.pick("req"), on_clear=lambda: self.clear("req"))
        self.req_card.grid(row=2, column=0, sticky="ew", pady=(0, 12), **pad)
        self.kx_card = FileCard(main, 2, "Kardex de compras",
                                "El reporte por órdenes («Consecutivo General») al que se le añadirán fechas y colores.",
                                on_pick=lambda: self.pick("kx"), on_clear=lambda: self.clear("kx"))
        self.kx_card.grid(row=3, column=0, sticky="ew", pady=(0, 12), **pad)

        # Opciones
        opts = ctk.CTkFrame(main, fg_color=CARD, corner_radius=14, border_width=1, border_color=CARD_BORDER)
        opts.grid(row=4, column=0, sticky="ew", pady=(0, 18), **pad)
        opts.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(opts, text="3", width=34, height=34, corner_radius=17, fg_color=ACCENT_SOFT,
                     text_color=ACCENT, font=(FONT, 15, "bold")).grid(row=0, column=0, rowspan=4, padx=(18, 14),
                                                                      pady=18, sticky="n")
        ctk.CTkLabel(opts, text="¿Qué quieres obtener?", font=(FONT, 15, "bold"), text_color=TEXT) \
            .grid(row=0, column=1, sticky="w", pady=(16, 6))
        mode = self.settings.output_mode if self.settings.output_mode in ("report", "copy", "overwrite") else "report"
        self.save_var = tk.StringVar(value=mode)
        radios = ctk.CTkFrame(opts, fg_color="transparent")
        radios.grid(row=1, column=1, sticky="w")
        for value, text in (("report", "Reporte claro: resumen, pestañas por estado y filtros (recomendado)"),
                            ("copy", "Kardex original pintado con colores, en una copia (como antes)"),
                            ("overwrite", "Kardex original pintado con colores, sobrescribiendo el archivo")):
            ctk.CTkRadioButton(radios, text=text, value=value, variable=self.save_var, font=(FONT, 12),
                               fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color=TEXT,
                               command=self._on_save_mode).pack(anchor="w", pady=3)
        self.save_hint = ctk.CTkLabel(opts, text="", font=(FONT, 11), text_color=MUTED, anchor="w")
        self.save_hint.grid(row=2, column=1, sticky="w", pady=(4, 6))

        chk = ctk.CTkFrame(opts, fg_color="transparent")
        chk.grid(row=3, column=1, sticky="w", pady=(2, 16))
        self.scoped_var = tk.BooleanVar(value=self.settings.group_scoped_filldown)
        ctk.CTkCheckBox(chk, text="Arrastrar la fecha solo dentro de cada orden", variable=self.scoped_var,
                        font=(FONT, 12), fg_color=ACCENT, hover_color=ACCENT_HOVER,
                        command=self._persist).pack(anchor="w", pady=2)
        self.open_var = tk.BooleanVar(value=self.settings.open_when_done)
        ctk.CTkCheckBox(chk, text="Abrir el archivo al terminar", variable=self.open_var, font=(FONT, 12),
                        fg_color=ACCENT, hover_color=ACCENT_HOVER, command=self._persist).pack(anchor="w", pady=2)
        self._on_save_mode(persist=False)

        # Acción
        act = ctk.CTkFrame(main, fg_color="transparent")
        act.grid(row=5, column=0, sticky="ew", **pad)
        act.grid_columnconfigure(1, weight=1)
        self.run_btn = ctk.CTkButton(act, text="Generar kardex", height=46, width=200, corner_radius=10,
                                     font=(FONT, 15, "bold"), fg_color=ACCENT, hover_color=ACCENT_HOVER,
                                     command=self.start)
        self.run_btn.grid(row=0, column=0, rowspan=2, sticky="w")
        self.cancel_btn = ctk.CTkButton(act, text="Cancelar", height=46, width=110, corner_radius=10,
                                        font=(FONT, 13), fg_color="transparent", border_width=1,
                                        border_color=CARD_BORDER, text_color=TEXT,
                                        hover_color=("#E5E7EB", "#243049"), command=self.cancel)
        self.status = ctk.CTkLabel(act, text="", font=(FONT, 12), text_color=MUTED, anchor="w")
        self.status.grid(row=0, column=1, sticky="ew", padx=(18, 0))
        self.progress = ctk.CTkProgressBar(act, height=8, corner_radius=4, progress_color=ACCENT)
        self.progress.set(0)

        # Resultados
        self.results = ctk.CTkFrame(main, fg_color="transparent")
        self.results.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="t")
        self.res_title = ctk.CTkLabel(self.results, text="", font=(FONT, 17, "bold"), text_color=TEXT, anchor="w")
        self.res_title.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 2))
        self.res_sub = ctk.CTkLabel(self.results, text="", font=(FONT, 12), text_color=MUTED, anchor="w",
                                    justify="left", wraplength=640)
        self.res_sub.grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 12))
        self.tiles = {
            "dates": StatTile(self.results, "Líneas con fecha"),
            "sin_oc": StatTile(self.results, "Sin orden de compra", STATUS_COLORS["sin_oc"]),
            "sin_entrada": StatTile(self.results, "Sin entrada/remisión", STATUS_COLORS["sin_entrada"]),
            "pendiente": StatTile(self.results, "Pendiente proveedor", STATUS_COLORS["pendiente"]),
        }
        for i, t in enumerate(self.tiles.values()):
            t.grid(row=2, column=i, sticky="ew", padx=(0 if i == 0 else 6, 0 if i == 3 else 6))
        self.missing_box = ctk.CTkFrame(self.results, fg_color=CARD, corner_radius=12, border_width=1,
                                        border_color=CARD_BORDER)
        self.missing_box.grid_columnconfigure(0, weight=1)
        self.missing_title = ctk.CTkLabel(self.missing_box, text="", font=(FONT, 13, "bold"), text_color=WARN,
                                          anchor="w")
        self.missing_title.grid(row=0, column=0, sticky="w", padx=14, pady=(12, 0))
        ctk.CTkButton(self.missing_box, text="Copiar lista", width=100, height=26, font=(FONT, 11),
                      fg_color="transparent", border_width=1, border_color=CARD_BORDER, text_color=TEXT,
                      hover_color=("#E5E7EB", "#243049"), command=self._copy_missing) \
            .grid(row=0, column=1, padx=14, pady=(12, 0))
        self.missing_text = ctk.CTkTextbox(self.missing_box, height=74, font=("Consolas", 12),
                                           fg_color="transparent", wrap="word")
        self.missing_text.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 10))

        btns = ctk.CTkFrame(self.results, fg_color="transparent")
        btns.grid(row=4, column=0, columnspan=4, sticky="w", pady=(14, 30))
        ctk.CTkButton(btns, text="Abrir archivo", height=38, corner_radius=8, font=(FONT, 13, "bold"),
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=lambda: self.last_result and self._open(self.last_result.output_path)) \
            .pack(side="left")
        ctk.CTkButton(btns, text="Mostrar en carpeta", height=38, corner_radius=8, font=(FONT, 13),
                      fg_color="transparent", border_width=1, border_color=CARD_BORDER, text_color=TEXT,
                      hover_color=("#E5E7EB", "#243049"),
                      command=lambda: self.last_result and self._reveal(self.last_result.output_path)) \
            .pack(side="left", padx=10)

        if self.dnd:
            for w in (self, self.req_card, self.kx_card):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", lambda e, target=w: self._on_drop(e, target))
                except Exception:  # noqa: BLE001
                    log.debug("No se pudo registrar el destino de arrastre", exc_info=True)

    # ------------------------------------------------------------------ archivos
    def pick(self, slot: str):
        if self.running:
            return
        title = "Selecciona el reporte de requisiciones" if slot == "req" else "Selecciona el kardex"
        path = filedialog.askopenfilename(
            title=title, initialdir=self.settings.last_dir or None,
            filetypes=[("Libros de Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*")])
        if path:
            self.set_file(slot, Path(path))

    def clear(self, slot: str):
        if self.running:
            return
        card = self.req_card if slot == "req" else self.kx_card
        if slot == "req":
            self.req_path, self.req_ok = None, False
        else:
            self.kx_path, self.kx_ok = None, False
        card.set_state("empty")
        card.reset_badge_number(1 if slot == "req" else 2)
        self._refresh_run_button()

    def _on_drop(self, event, target):
        files = [Path(p) for p in self.tk.splitlist(event.data)]
        files = [f for f in files if f.suffix.lower() in (".xlsx", ".xlsm", ".xls")]
        if not files:
            return
        if target is self.req_card:
            self.set_file("req", files[0])
        elif target is self.kx_card:
            self.set_file("kx", files[0])
        else:  # soltado en la ventana: se asigna según el contenido
            for f in files[:2]:
                self._auto_assign(f)

    def _auto_assign(self, path: Path):
        def done(role, _err):
            if role == "kardex" or (role is None and self.kx_path is None and self.req_path is not None):
                self.set_file("kx", path)
            else:
                self.set_file("req", path)
        _run_bg(lambda: guess_file_role(path), done, self)

    def set_file(self, slot: str, path: Path):
        card = self.req_card if slot == "req" else self.kx_card
        if slot == "req":
            self.req_path, self.req_ok = path, False
        else:
            self.kx_path, self.kx_ok = path, False
        self.settings.last_dir = str(path.parent)
        card.set_state("loading", path.name, "Analizando el archivo…")
        self._refresh_run_button()

        if slot == "req":
            def job():
                scan = scan_requisitions(path)
                role = guess_file_role(path) if not scan.dates else None
                return scan, role
        else:
            def job():
                scan = scan_kardex(path)
                role = guess_file_role(path) if scan.groups == 0 else None
                return scan, role

        def done(res, err):
            current = self.req_path if slot == "req" else self.kx_path
            if current != path:  # el usuario ya cambió el archivo
                return
            if err is not None:
                msg = err.message if isinstance(err, KardexError) else f"No se pudo leer: {err}"
                hint = err.hint if isinstance(err, KardexError) else ""
                card.set_state("error", path.name, f"{msg} {hint}".strip())
                log.warning("Validación fallida %s: %s", path, err)
                return
            scan, role = res
            if slot == "req":
                if scan.dates:
                    extra = []
                    if scan.codes_without_date:
                        extra.append(f"{_fmt(len(scan.codes_without_date))} sin fecha")
                    if scan.unparsed_dates:
                        extra.append(f"{_fmt(scan.unparsed_dates)} con fecha en texto")
                    tail = f"  ·  {', '.join(extra)}" if extra else ""
                    card.set_state("ok", path.name, f"{_fmt(len(scan.dates))} requisiciones con fecha{tail}")
                    self.req_ok = True
                else:
                    self._wrong_file(card, path, role, "No se encontraron códigos REQUI-… con fecha.", "req")
            else:
                if scan.groups:
                    card.set_state("ok", path.name, f"{_fmt(scan.groups)} órdenes · {_fmt(scan.rows)} filas · "
                                                    f"{_fmt(scan.requisitions)} referencias a requisiciones")
                    self.kx_ok = True
                else:
                    self._wrong_file(card, path, role,
                                     "No se encontraron filas «Consecutivo General» en la columna B.", "kx")
            self._refresh_run_button()

        _run_bg(job, done, self)

    def _wrong_file(self, card: FileCard, path: Path, role: Optional[str], msg: str, slot: str):
        other = "kx" if slot == "req" else "req"
        expected_other = "kardex" if slot == "req" else "requisiciones"
        if role == expected_other:
            card.set_state("warn", path.name,
                           f"Este archivo parece ser el {'kardex' if slot == 'req' else 'reporte de requisiciones'}.",
                           action=("Mover al paso " + ("2" if slot == "req" else "1"),
                                   lambda: self._move(slot, other, path)))
        else:
            card.set_state("error", path.name, msg + " ¿Es el archivo correcto?")

    def _move(self, from_slot: str, to_slot: str, path: Path):
        other_path = self.kx_path if to_slot == "kx" else self.req_path
        self.clear(from_slot)
        self.set_file(to_slot, path)
        if other_path and other_path != path:
            self.set_file(from_slot, other_path)

    # ------------------------------------------------------------------ opciones
    def _on_save_mode(self, persist: bool = True):
        name = self.kx_path.stem if self.kx_path else "kardex"
        mode = self.save_var.get()
        if mode == "report":
            hint = (f"Se crea «Reporte Kardex - {name} - AAAA-MM-DD HHMM.xlsx» junto al kardex. "
                    "Funciona con archivos de cualquier tamaño.")
        elif mode == "overwrite":
            hint = ("Se guarda sobre el kardex original (antes se crea un respaldo). "
                    "Con archivos muy grandes es más lento y usa mucha memoria.")
        else:
            hint = (f"Se crea «{name} - procesado AAAA-MM-DD HHMM.xlsx» junto al original. "
                    "Con archivos muy grandes es más lento y usa mucha memoria.")
        self.save_hint.configure(text=hint)
        if persist:
            self._persist()

    def _on_theme(self, value: str):
        mode = {"Claro": "light", "Oscuro": "dark"}.get(value, "system")
        ctk.set_appearance_mode(mode)
        self.settings.appearance = mode
        self.settings.save()

    def _persist(self):
        s = self.settings
        s.output_mode = self.save_var.get()
        s.group_scoped_filldown = bool(self.scoped_var.get())
        s.open_when_done = bool(self.open_var.get())
        s.save()

    def _refresh_run_button(self):
        ready = self.req_ok and self.kx_ok and not self.running
        self.run_btn.configure(state="normal" if ready else "disabled",
                               fg_color=ACCENT if ready or self.running else ("#D1D5DB", "#2A3448"),
                               text_color_disabled=("#F9FAFB", "#6B7280") if not self.running else "#FFFFFF")
        self._on_save_mode(persist=False)
        if not self.running:
            if ready:
                self.status.configure(text="Todo listo. Pulsa «Generar kardex» o Enter.", text_color=MUTED)
            elif self.req_path is None or self.kx_path is None:
                self.status.configure(text="Faltan archivos por seleccionar.", text_color=MUTED)
            else:
                self.status.configure(text="Revisa los archivos marcados antes de continuar.", text_color=MUTED)

    # ------------------------------------------------------------------ ejecución
    def start(self):
        if self.running or not (self.req_ok and self.kx_ok):
            return
        self._persist()
        self.settings.last_requisitions = str(self.req_path)
        self.settings.last_kardex = str(self.kx_path)
        self.settings.save()

        self.running = True
        self.cancel_event = threading.Event()
        self.results.grid_forget()
        self.run_btn.configure(state="disabled", text="Procesando…")
        self.cancel_btn.grid(row=0, column=2, rowspan=2, padx=(12, 0))
        self.progress.grid(row=1, column=1, sticky="ew", padx=(18, 0), pady=(4, 0))
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        for card in (self.req_card, self.kx_card):
            card.pick_btn.configure(state="disabled")

        mode = self.settings.output_mode
        args = dict(requisitions_path=self.req_path, kardex_path=self.kx_path,
                    group_scoped_filldown=self.settings.group_scoped_filldown,
                    progress=lambda m, f: self.worker_q.put(("progress", m, f)),
                    cancel=self.cancel_event)
        if mode != "report":
            args["overwrite"] = mode == "overwrite"
        log.info("Inicio (%s): %s", mode, {k: v for k, v in args.items() if k not in ("progress", "cancel")})
        job = build_report if mode == "report" else process_kardex

        def worker():
            try:
                self.worker_q.put(("done", job(**args)))
            except Cancelled:
                self.worker_q.put(("cancelled",))
            except MemoryError:
                log.exception("Sin memoria")
                self.worker_q.put(("error", "El equipo se quedó sin memoria procesando un archivo tan grande.",
                                   "Elige la opción «Reporte claro», que procesa el archivo por partes y "
                                   "funciona con archivos de cualquier tamaño."))
            except KardexError as exc:
                log.warning("Error controlado: %s | %s", exc.message, exc.hint)
                self.worker_q.put(("error", exc.message, exc.hint))
            except Exception as exc:  # noqa: BLE001
                log.exception("Error inesperado")
                self.worker_q.put(("error", "Ocurrió un error inesperado al procesar los archivos.",
                                   f"{type(exc).__name__}: {exc}\nDetalles en el registro: {log_file()}"))

        threading.Thread(target=worker, daemon=True).start()
        self.after(50, self._poll)

    def _poll(self):
        last_progress = None
        try:
            while True:
                msg = self.worker_q.get_nowait()
                if msg[0] == "progress":
                    last_progress = msg
                else:
                    self._finish(msg)
                    return
        except queue.Empty:
            pass
        if last_progress:
            _, text, frac = last_progress
            self.status.configure(text=text, text_color=MUTED)
            if frac is None:
                if self.progress.cget("mode") != "indeterminate":
                    self.progress.configure(mode="indeterminate")
                    self.progress.start()
            else:
                if self.progress.cget("mode") != "determinate":
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
                self.progress.set(frac)
        self.after(50, self._poll)

    def _finish(self, msg: tuple):
        self.running = False
        self.progress.stop()
        self.progress.grid_forget()
        self.cancel_btn.grid_forget()
        self.run_btn.configure(text="Generar kardex")
        for card in (self.req_card, self.kx_card):
            card.pick_btn.configure(state="normal")
        self._refresh_run_button()

        kind = msg[0]
        if kind == "done":
            self._show_result(msg[1])
        elif kind == "cancelled":
            self.status.configure(text="Proceso cancelado. No se modificó ningún archivo.", text_color=WARN)
        else:
            _, text, hint = msg
            self.status.configure(text=text, text_color=ERR)
            messagebox.showerror(APP_NAME, f"{text}\n\n{hint}".strip(), parent=self)

    def _show_result(self, r: KardexResult):
        self.last_result = r
        self.status.configure(text=f"Listo en {r.elapsed_s:.1f} s.", text_color=OK)
        is_report = r.output_path.name.startswith("Reporte Kardex")
        self.res_title.configure(text="✓  Reporte generado" if is_report else "✓  Kardex generado")
        sub = f"Archivo: {r.output_path.name}\nCarpeta: {r.output_path.parent}"
        if r.backup_path:
            sub += f"\nRespaldo del original: {r.backup_path.name}"
        self.res_sub.configure(text=sub)
        self.tiles["dates"].set(_fmt(r.dates_written))
        for k in ("sin_oc", "sin_entrada", "pendiente"):
            self.tiles[k].set(_fmt(r.status_counts.get(k, 0)))
        if r.missing_requisitions:
            n = len(r.missing_requisitions)
            self.missing_title.configure(
                text=f"{_fmt(n)} requisición{'es' if n != 1 else ''} del kardex no aparece{'n' if n != 1 else ''} "
                     "en el reporte de requisiciones (quedaron sin fecha)")
            self.missing_text.configure(state="normal")
            self.missing_text.delete("1.0", "end")
            self.missing_text.insert("1.0", ", ".join(r.missing_requisitions))
            self.missing_text.configure(state="disabled")
            self.missing_box.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(12, 0))
        else:
            self.missing_box.grid_forget()
        self.results.grid(row=6, column=0, sticky="ew", padx=36, pady=(26, 0))
        self.after(100, lambda: self.main._parent_canvas.yview_moveto(1.0))
        if self.settings.open_when_done:
            self._open(r.output_path)

    def cancel(self):
        if self.running and self.cancel_event:
            self.cancel_event.set()
            self.status.configure(text="Cancelando…", text_color=WARN)

    def on_close(self):
        if self.running:
            if not messagebox.askyesno(APP_NAME, "Hay un proceso en curso. ¿Cancelarlo y salir?", parent=self):
                return
            self.cancel()
        self._persist()
        self.destroy()

    # ------------------------------------------------------------------ utilidades
    def _copy_missing(self):
        if self.last_result:
            self.clipboard_clear()
            self.clipboard_append("\n".join(self.last_result.missing_requisitions))
            self.status.configure(text="Lista copiada al portapapeles.", text_color=MUTED)

    def _open(self, path: Path):
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"No se pudo abrir:\n{path}\n\n{exc}", parent=self)

    def _reveal(self, path: Path):
        try:
            subprocess.Popen(["explorer", "/select,", str(path)])
        except Exception:  # noqa: BLE001
            self._open(path.parent)


def main() -> None:
    from .settings import setup_logging
    setup_logging()
    log.info("%s %s iniciado", APP_NAME, __version__)
    app = App()  # customtkinter ya activa el escalado DPI de Windows
    app.mainloop()
