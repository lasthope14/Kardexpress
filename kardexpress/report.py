"""
Reporte profesional de Kardexpress (v2.1).

Mantiene la lógica de la v1 (fecha por requisición + semáforo amarillo/verde/rojo), pero en
lugar de pintar el kardex original genera un libro nuevo, pensado para alguien sin
conocimientos técnicos:

  Resumen               -> qué significa cada color, cuántas líneas hay y qué hacer.
  Pendiente proveedor   -> lista de trabajo (una pestaña por estado).
  Sin entrada
  Sin orden de compra
  Todas las líneas      -> tabla plana con filtros (incluido «Filtrar por color»).

Todo se procesa en streaming (lectura `read_only` + escritura `write_only`): la memoria usada
no crece con el tamaño del kardex, así que los archivos grandes ya no congelan el equipo.
"""
from __future__ import annotations

import datetime as dt
import logging
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.worksheet.properties import PageSetupProperties

from . import __version__
from .core import (DEFAULT_LAYOUT, GROUP_LABEL, HEADER_LABEL, Cancelled, KardexError, KardexResult, Layout,
                   ProgressFn, _atomic_save, _norm_text, _open_workbook, classify_row, is_file_locked,
                   normalize_req, scan_requisitions, to_number)

log = logging.getLogger(__name__)


# --------------------------------------------------------------------------- estados
@dataclass(frozen=True)
class StateInfo:
    key: str
    title: str
    sheet: str
    strong: str          # color de la celda «Estado» (y del filtro por color)
    tint: Optional[str]  # color suave de toda la fila
    meaning: str
    action: str


STATES: dict[str, StateInfo] = {s.key: s for s in (
    StateInfo("pendiente", "Pendiente proveedor", "Pendiente proveedor", "F4A6A0", "FDE7E5",
              "Ya entró una parte al almacén, pero el proveedor todavía debe el resto.",
              "Reclamar al proveedor la cantidad pendiente."),
    StateInfo("sin_entrada", "Sin entrada / remisión", "Sin entrada", "8FD694", "E6F6E7",
              "Ya se hizo la compra, pero al almacén no ha entrado nada.",
              "Preguntar al proveedor por la entrega o registrar la entrada si ya llegó."),
    StateInfo("sin_oc", "Sin orden de compra", "Sin orden de compra", "FFD966", "FFF6D6",
              "Se pidió el material, pero todavía no se ha comprado (cantidad comprada = 0).",
              "Avisar a Compras para que genere la orden de compra."),
    StateInfo("completo", "Completo", "", "CFE0F3", None,
              "Todo lo comprado ya entró al almacén.",
              "Nada: está al día."),
    StateInfo("sin_datos", "Sin cantidades", "", "E5E7EB", None,
              "La línea no tiene cantidades (por ejemplo, notas o subtotales).",
              "Revisar solo si parece un error."),
)}
ACTION_STATES = ("pendiente", "sin_entrada", "sin_oc")
ALL_SHEET = "Todas las líneas"

NAVY = "0B4A75"       # azul Hidroobras
INK = "1F2937"
MUTED = "6B7280"
LINE = "D1D5DB"
DATE_FMT = "dd-mmm-yyyy"     # «15-mar-2023»: no se confunde día con mes


def row_state(bought, received, pending) -> str:
    """Igual que el semáforo de la v1, más dos estados informativos para las líneas sin color."""
    status = classify_row(bought, received, pending)
    if status:
        return status
    if any(to_number(v) is not None for v in (bought, received, pending)):
        return "completo"
    return "sin_datos"


# --------------------------------------------------------------------------- análisis previo
@dataclass
class KardexShape:
    rows: int = 0
    headers: dict[int, str] = field(default_factory=dict)   # col (1-based) -> título
    data_cols: set[int] = field(default_factory=set)
    decimal_cols: set[int] = field(default_factory=set)    # columnas con algún valor con decimales
    has_group_info: bool = False
    groups: int = 0


def _analyze(ws, layout: Layout, cancel, report) -> KardexShape:
    """Primera pasada (solo lectura, streaming): títulos de columnas y columnas con datos."""
    shape = KardexShape()
    seen_group = False
    skip = {layout.kx_label_col, layout.kx_req_col}
    for n, row in enumerate(ws.iter_rows(values_only=True), start=1):
        shape.rows = n
        if n % 5000 == 0:
            if cancel is not None and cancel.is_set():
                raise Cancelled()
            report(f"Analizando kardex… {n:,} filas".replace(",", "."), None)
        label = _norm_text(row[layout.kx_label_col - 1]) if len(row) >= layout.kx_label_col else ""
        if label.startswith(GROUP_LABEL):
            seen_group = True
            shape.groups += 1
            if not shape.has_group_info:
                shape.has_group_info = any(v not in (None, "") for c, v in enumerate(row, 1) if c not in skip)
            continue
        if label == HEADER_LABEL:
            if not shape.headers:
                shape.headers = {c: str(v).strip() for c, v in enumerate(row, 1) if v not in (None, "")}
            continue
        if seen_group:
            for c, v in enumerate(row, 1):
                if v not in (None, ""):
                    shape.data_cols.add(c)
                    if isinstance(v, float) and not v.is_integer():
                        shape.decimal_cols.add(c)
    return shape


def _order_label(value) -> str:
    s = " ".join(str(value or "").split())
    rest = s[len("Consecutivo General"):] if _norm_text(s).startswith(GROUP_LABEL) else s
    return rest.strip(" :#-.") or s


# --------------------------------------------------------------------------- estilos
def _styles(wb) -> dict[str, NamedStyle]:
    thin = Side(style="thin", color=LINE)
    styles = {
        "hdr": NamedStyle("kx_hdr", font=Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
                          fill=PatternFill("solid", start_color=NAVY, end_color=NAVY),
                          alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
                          border=Border(bottom=thin)),
        "date": NamedStyle("kx_date", number_format=DATE_FMT, alignment=Alignment(horizontal="center")),
        "int": NamedStyle("kx_int", number_format="#,##0"),
        "dec": NamedStyle("kx_dec", number_format="#,##0.00"),
        "days": NamedStyle("kx_days", number_format="#,##0", alignment=Alignment(horizontal="center")),
        "center": NamedStyle("kx_center", alignment=Alignment(horizontal="center")),
        "title": NamedStyle("kx_title", font=Font(name="Calibri", size=20, bold=True, color=NAVY)),
        "subtitle": NamedStyle("kx_subtitle", font=Font(name="Calibri", size=10, color=MUTED)),
        "h2": NamedStyle("kx_h2", font=Font(name="Calibri", size=13, bold=True, color=NAVY),
                         border=Border(bottom=Side(style="medium", color=NAVY))),
        "text": NamedStyle("kx_text", font=Font(name="Calibri", size=11, color=INK),
                           alignment=Alignment(vertical="top", wrap_text=True)),
        "bold": NamedStyle("kx_bold", font=Font(name="Calibri", size=11, bold=True, color=INK),
                           alignment=Alignment(vertical="top", wrap_text=True)),
        "big": NamedStyle("kx_big", font=Font(name="Calibri", size=14, bold=True, color=INK),
                          number_format="#,##0", alignment=Alignment(horizontal="center", vertical="top")),
        "pct": NamedStyle("kx_pct", font=Font(name="Calibri", size=11, color=MUTED), number_format="0.0%",
                          alignment=Alignment(horizontal="center", vertical="top")),
        "link": NamedStyle("kx_link", font=Font(name="Calibri", size=11, color="2563EB", underline="single"),
                           alignment=Alignment(vertical="top", wrap_text=True)),
        "code": NamedStyle("kx_code", font=Font(name="Consolas", size=10, color=INK),
                           alignment=Alignment(vertical="top", wrap_text=True)),
    }
    for key, st in STATES.items():
        styles[f"st_{key}"] = NamedStyle(
            f"kx_st_{key}", font=Font(name="Calibri", size=11, bold=True, color=INK),
            fill=PatternFill("solid", start_color=st.strong, end_color=st.strong),
            alignment=Alignment(vertical="top", wrap_text=True))
    for st in styles.values():
        wb.add_named_style(st)
    return styles


def _width_for(title: str) -> float:
    t = _norm_text(title)
    if any(k in t for k in ("descrip", "nombre", "proveedor", "articulo", "observ", "detalle")):
        return 42
    return max(11, min(28, len(title) * 1.15 + 3))


# --------------------------------------------------------------------------- hoja de datos
class _DataSheet:
    """Hoja con encabezado fijo, filtros y color por estado. Escribe en streaming."""

    def __init__(self, wb, title: str, columns: list[tuple[str, float]], styles, tab_color: Optional[str],
                 num_styles: list[Optional[str]]):
        self.ws = wb.create_sheet(title)
        self.styles = styles
        self.num_styles = [styles[s] if s else None for s in num_styles]   # alineado con `values`
        self.ncols = len(columns)
        self.rows = 1
        ws = self.ws
        if tab_color:
            ws.sheet_properties.tabColor = tab_color
        ws.freeze_panes = "B2"                        # encabezado y columna «Estado» siempre visibles
        ws.sheet_view.zoomScale = 90
        for i, (_, width) in enumerate(columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_title_rows = "1:1"
        ws.row_dimensions[1].height = 32
        ws.append([self._cell(name, styles["hdr"]) for name, _ in columns])

    def _cell(self, value, style=None):
        c = WriteOnlyCell(self.ws, value)
        if style is not None:
            c.style = style.name
        return c

    def append(self, state: str, values: list) -> None:
        out = [self._cell(STATES[state].title, self.styles[f"st_{state}"])]
        for v, num_style in zip(values, self.num_styles):
            if isinstance(v, (dt.datetime, dt.date)):
                out.append(self._cell(v, self.styles["date"]))
            elif num_style is not None and isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append(self._cell(v, num_style))
            else:
                out.append(v)
        self.ws.append(out)
        self.rows += 1

    def finish(self) -> None:
        last_col = get_column_letter(self.ncols)
        ws = self.ws
        ws.auto_filter.ref = f"A1:{last_col}{max(self.rows, 2)}"
        if self.rows > 1:
            # Tono suave en toda la fila según el estado (se actualiza solo si alguien cambia el estado)
            for st in STATES.values():
                if st.tint:
                    ws.conditional_formatting.add(
                        f"B2:{last_col}{self.rows}",
                        FormulaRule(formula=[f'$A2="{st.title}"'],
                                    fill=PatternFill("solid", start_color=st.tint, end_color=st.tint)))


# --------------------------------------------------------------------------- proceso principal
def build_report(
    requisitions_path: Path,
    kardex_path: Path,
    output_path: Optional[Path] = None,
    *,
    group_scoped_filldown: bool = True,
    layout: Layout = DEFAULT_LAYOUT,
    progress: Optional[ProgressFn] = None,
    cancel: Optional[threading.Event] = None,
    today: Optional[dt.date] = None,
) -> KardexResult:
    t0 = time.perf_counter()
    report = progress or (lambda _m, _f: None)
    today = today or dt.date.today()
    requisitions_path, kardex_path = Path(requisitions_path), Path(kardex_path)

    def check_cancel():
        if cancel is not None and cancel.is_set():
            raise Cancelled()

    if requisitions_path.resolve() == kardex_path.resolve():
        raise KardexError("Seleccionaste el mismo archivo para requisiciones y para kardex.")
    target = Path(output_path or default_report_path(kardex_path))
    if target.exists() and is_file_locked(target):
        raise KardexError(f"«{target.name}» está abierto en otro programa.",
                          "Cierra el archivo en Excel e inténtalo de nuevo.")

    # 1) Requisiciones ----------------------------------------------------------
    report("Leyendo requisiciones…", None)
    req = scan_requisitions(requisitions_path, layout, cancel)
    if not req.dates:
        raise KardexError(
            f"No se encontró ninguna requisición con fecha en «{requisitions_path.name}».",
            f"Se esperan códigos «{layout.req_prefix}…» en la columna {get_column_letter(layout.req_code_col)} "
            f"y la fecha en la columna {get_column_letter(layout.req_date_col)} de la fila siguiente. "
            "¿Seleccionaste el archivo correcto?")
    check_cancel()

    # 2) Kardex: primera pasada --------------------------------------------------
    report("Analizando kardex…", None)
    src = _open_workbook(kardex_path, "kardex", read_only=True, data_only=True)
    try:
        shape = _analyze(src.active, layout, cancel, report)
        if shape.groups == 0:
            raise KardexError(f"«{kardex_path.name}» no parece un kardex: no tiene filas «Consecutivo General».",
                              "Revisa que seleccionaste el archivo correcto.")
        # E (requisición) ya sale como columna propia; S y T son las que escribía la v1.
        excluded = {layout.kx_req_col, layout.kx_date_col, layout.legend_col}
        orig_cols = sorted(c for c in (shape.data_cols | set(shape.headers)) if c not in excluded)
        orig_titles = [shape.headers.get(c) or f"Columna {get_column_letter(c)}" for c in orig_cols]

        columns = [("Estado", 24), ("Días desde la requisición", 13), ("Fecha requisición", 14),
                   ("Requisición", 15), ("Orden", 14)]
        num_styles: list[Optional[str]] = ["days", None, None, None]
        if shape.has_group_info:
            columns.append(("Datos de la orden", 34))
            num_styles.append(None)
        columns += [(t, _width_for(t)) for t in orig_titles]
        num_styles += ["dec" if c in shape.decimal_cols else "int" for c in orig_cols]

        # 3) Libro de salida (streaming) ------------------------------------------
        out = openpyxl.Workbook(write_only=True)
        styles = _styles(out)
        summary = out.create_sheet("Resumen")
        sheets = {k: _DataSheet(out, STATES[k].sheet, columns, styles, STATES[k].strong, num_styles)
                  for k in ACTION_STATES}
        all_sheet = _DataSheet(out, ALL_SHEET, columns, styles, NAVY, num_styles)

        counts = {k: 0 for k in STATES}
        missing: dict[str, None] = {}
        dates_written = 0
        order = order_info = group_req = None
        carry = None
        seen_group = False
        total = max(shape.rows, 1)
        step = max(1, total // 300)
        label_i = layout.kx_label_col - 1
        req_i = layout.kx_req_col - 1
        m_i, n_i, r_i = layout.kx_qty_bought_col - 1, layout.kx_qty_received_col - 1, layout.kx_pending_col - 1
        info_skip = {layout.kx_label_col, layout.kx_req_col}

        def at(row, i):
            return row[i] if i < len(row) else None

        for n, row in enumerate(src.active.iter_rows(values_only=True), start=1):
            if n % step == 0:
                check_cancel()
                report(f"Generando reporte: fila {n:,} de {total:,}…".replace(",", "."), n / total)
            label = _norm_text(at(row, label_i))
            code = normalize_req(at(row, req_i), layout.req_prefix)

            if label.startswith(GROUP_LABEL):
                seen_group = True
                order = _order_label(at(row, label_i))
                order_info = " · ".join(str(v).strip() for c, v in enumerate(row, 1)
                                        if c not in info_skip and v not in (None, "")) or None
                group_req = code
                if code is not None and code not in req.dates:
                    missing[code] = None
                if group_scoped_filldown:
                    carry = req.dates.get(code) if code else None
                elif code and code in req.dates:
                    carry = req.dates[code]
                continue
            if label == HEADER_LABEL or not seen_group:
                continue
            if all(v in (None, "") for v in row):
                continue

            line_req = code or group_req
            date = None
            if code is not None:
                date = req.dates.get(code)
                if date is None:
                    missing[code] = None
            if date is None:
                date = carry
            else:
                carry = date
            if date is not None:
                dates_written += 1

            state = row_state(at(row, m_i), at(row, n_i), at(row, r_i))
            counts[state] += 1
            days = (today - date).days if isinstance(date, dt.date) else None

            values = [days, date, line_req, order]
            if shape.has_group_info:
                values.append(order_info)
            values += [at(row, c - 1) for c in orig_cols]

            all_sheet.append(state, values)
            if state in sheets:
                sheets[state].append(state, values)

        for s in (*sheets.values(), all_sheet):
            s.finish()
        check_cancel()

        _write_summary(summary, styles, counts, sorted(missing), requisitions_path, kardex_path)

        # 4) Guardar --------------------------------------------------------------
        report("Guardando reporte…", 1.0)
        _atomic_save(out, target)
    finally:
        src.close()

    result = KardexResult(
        output_path=target, backup_path=None, requisitions_loaded=len(req.dates),
        kardex_rows=shape.rows, dates_written=dates_written, rows_filled_down=0,
        missing_requisitions=sorted(missing), status_counts=counts,
        elapsed_s=time.perf_counter() - t0)
    log.info("Reporte generado: %s", result)
    return result


def default_report_path(kardex_path: Path, now: Optional[dt.datetime] = None) -> Path:
    now = now or dt.datetime.now()
    return kardex_path.with_name(f"Reporte Kardex - {kardex_path.stem} - {now:%Y-%m-%d %H%M}.xlsx")


# --------------------------------------------------------------------------- hoja Resumen
def _plural(n: int, one: str, many: str) -> str:
    return f"{n:,}".replace(",", ".") + " " + (one if n == 1 else many)


def _write_summary(ws, styles, counts, missing, req_path: Path, kx_path: Path) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    widths = {"A": 3, "B": 26, "C": 11, "D": 9, "E": 50, "F": 46, "G": 3}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    text_chars = int(sum(widths[c] for c in "BCDEF") * 1.05)   # caracteres por línea en B:F

    def C(value, style=None):
        c = WriteOnlyCell(ws, value)
        if style:
            c.style = styles[style].name
        return c

    row = [0]

    def line(*cells, height=None):
        if height:
            ws.row_dimensions[row[0] + 1].height = height   # en write_only va antes de escribir la fila
        ws.append([None, *cells])
        row[0] += 1

    def para(text, style="text"):
        """Párrafo que ocupa B:F, con altura según su largo."""
        lines = max(1, -(-len(text) // text_chars))
        line(C(text, style), height=16 * lines + 4)
        ws.merged_cells.add(f"B{row[0]}:F{row[0]}")

    def heading(text):
        line()
        line(C(text, "h2"), *(C(None, "h2") for _ in range(4)), height=22)

    total = sum(counts.values())
    problems = sum(counts[k] for k in ACTION_STATES)
    fmt = lambda n: f"{n:,}".replace(",", ".")

    line(C("Reporte de Kardex", "title"), height=34)
    line(C(f"Generado el {dt.datetime.now():%d/%m/%Y a las %H:%M} con Kardexpress {__version__}", "subtitle"))
    line(C(f"Requisiciones: {req_path.name}   ·   Kardex: {kx_path.name}", "subtitle"))

    heading("En pocas palabras")
    if problems:
        para(f"Se revisaron {_plural(total, 'línea', 'líneas')} del kardex y {fmt(problems)} "
             f"({problems / total:.0%}) necesitan seguimiento. La tabla de abajo explica qué significa cada "
             "color y qué hay que hacer en cada caso.", "bold")
    else:
        para(f"Se revisaron {_plural(total, 'línea', 'líneas')} y ninguna necesita seguimiento.", "bold")
    line()

    line(C("Estado", "hdr"), C("Líneas", "hdr"), C("%", "hdr"), C("¿Qué significa?", "hdr"),
         C("¿Qué hay que hacer?", "hdr"), height=24)
    keys = [k for k in (*ACTION_STATES, "completo", "sin_datos") if k != "sin_datos" or counts[k]]
    first_data = row[0] + 1
    for key in keys:
        st = STATES[key]
        name = C(st.title, f"st_{key}")
        if st.sheet:
            name.hyperlink = f"#{quote_sheetname(st.sheet)}!A1"
        line(name, C(counts[key], "big"), C(counts[key] / total if total else 0, "pct"),
             C(st.meaning, "text"), C(st.action, "bold" if key in ACTION_STATES else "text"), height=36)
    last_data = row[0]
    line(C("Total", "bold"), C(total, "big"))

    heading("Cómo usar este reporte")
    for s in (
        "1.  Las pestañas de colores (abajo) son listas de trabajo: cada una tiene solo las líneas de ese "
        "estado. También puedes hacer clic en el nombre del estado en la tabla de arriba para ir directo.",
        f"2.  La pestaña «{ALL_SHEET}» tiene todo el kardex. Para ver solo un estado, haz clic en la flecha ▼ "
        "del título «Estado» y marca el que necesites, o usa «Filtrar por color».",
        "3.  La columna «Días desde la requisición» dice cuánto tiempo lleva esperando cada línea: "
        "mientras más alto el número, más urgente.",
        "4.  Las fechas se muestran como día-mes-año (por ejemplo 15-mar-2023) para que no haya confusiones.",
    ):
        para(s)

    if missing:
        heading("Requisiciones sin fecha")
        n = len(missing)
        para(f"{_plural(n, 'requisición aparece', 'requisiciones aparecen')} en el kardex pero no en el "
             f"reporte de requisiciones, por eso {'no tiene' if n == 1 else 'no tienen'} fecha. Normalmente pasa cuando el reporte de "
             "requisiciones no cubre todo el período del kardex.")
        per_line = max(1, text_chars // 13)
        for i in range(0, n, per_line):
            line(C("  ".join(missing[i:i + per_line]), "code"))
            ws.merged_cells.add(f"B{row[0]}:F{row[0]}")

    # Gráfico de barras con los colores de cada estado
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Líneas por estado"
    chart.legend = None
    chart.y_axis.majorGridlines = None
    chart.y_axis.delete = True
    chart.x_axis.scaling.orientation = "maxMin"
    chart.add_data(Reference(ws, min_col=3, min_row=first_data, max_row=last_data), titles_from_data=False)
    chart.set_categories(Reference(ws, min_col=2, min_row=first_data, max_row=last_data))
    series = chart.series[0]
    for i, key in enumerate(keys):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = STATES[key].strong
        pt.graphicalProperties.line.solidFill = STATES[key].strong
        series.dPt.append(pt)
    series.dLbls = DataLabelList()
    series.dLbls.showVal = True
    chart.gapWidth = 60
    chart.height = 6.5
    chart.width = 18
    heading("Gráfico")
    ws.add_chart(chart, f"B{row[0] + 2}")
