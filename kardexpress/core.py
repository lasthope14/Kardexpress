"""
Motor de Kardexpress (sin interfaz gráfica).

Reproduce el comportamiento de la v1 (ver docs/kardex_v1_reconstruido.py) pero:
  * nunca deja el proceso "colgado": todo error se convierte en KardexError con un
    mensaje entendible y la UI siempre recupera el control;
  * informa progreso real y admite cancelación;
  * no destruye el archivo original (guarda una copia, o respalda antes de sobrescribir);
  * recorre la hoja una sola vez en lugar de seis;
  * tolera datos sucios (espacios, números guardados como texto, fechas como texto o
    número de serie de Excel, celdas no textuales en la columna de códigos).
"""
from __future__ import annotations

import datetime as dt
import logging
import os
import re
import shutil
import tempfile
import threading
import time
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

log = logging.getLogger(__name__)

ProgressFn = Callable[[str, Optional[float]], None]
"""callback(mensaje, fracción 0..1 o None si es indeterminado)."""


# --------------------------------------------------------------------------- layout
@dataclass(frozen=True)
class Layout:
    """Posiciones (1-based) de los datos en ambos reportes. Valores tomados de la v1."""

    # Reporte de requisiciones
    req_code_col: int = 2            # B: "REQUI-0001"
    req_date_col: int = 23           # W: fecha de creación...
    req_date_row_offset: int = 1     # ...en la fila siguiente al código
    req_prefix: str = "REQUI-"

    # Kardex
    kx_req_col: int = 5              # E: número de requisición
    kx_label_col: int = 2            # B: "Consecutivo General" / "Código"
    kx_date_col: int = 19            # S: aquí se escribe la fecha
    kx_first_row: int = 2
    kx_style_from_row: int = 7       # formato de la columna de fecha
    kx_color_from_row: int = 8       # semáforo
    kx_qty_bought_col: int = 13      # M: cantidad comprada
    kx_qty_received_col: int = 14    # N: cantidad entrada
    kx_pending_col: int = 18         # R: pendiente
    legend_col: int = 20             # T
    legend_first_row: int = 3


DEFAULT_LAYOUT = Layout()

GROUP_LABEL = "consecutivo general"
HEADER_LABEL = "codigo"

COLOR_YELLOW = "FFFF00"
COLOR_GREEN = "00FF00"
COLOR_RED = "FF0000"

STATUS_LABELS = {
    "sin_oc": "SIN ORDEN DE COMPRA",
    "sin_entrada": "SIN ENTRADA/REMISION",
    "pendiente": "PENDIENTE PROVEEDOR",
}
STATUS_COLORS = {"sin_oc": COLOR_YELLOW, "sin_entrada": COLOR_GREEN, "pendiente": COLOR_RED}


# --------------------------------------------------------------------------- errores
class KardexError(Exception):
    """Error esperado, con mensaje listo para mostrar al usuario."""

    def __init__(self, message: str, hint: str = ""):
        super().__init__(message)
        self.message = message
        self.hint = hint


class Cancelled(Exception):
    pass


# --------------------------------------------------------------------------- resultados
@dataclass
class RequisitionScan:
    path: Path
    dates: dict[str, object]                 # código normalizado -> date | str
    codes_without_date: list[str] = field(default_factory=list)
    unparsed_dates: int = 0
    rows: int = 0


@dataclass
class KardexScan:
    path: Path
    rows: int
    groups: int
    requisitions: int


@dataclass
class KardexResult:
    output_path: Path
    backup_path: Optional[Path]
    requisitions_loaded: int
    kardex_rows: int
    dates_written: int
    rows_filled_down: int
    missing_requisitions: list[str]
    status_counts: dict[str, int]
    elapsed_s: float


# --------------------------------------------------------------------------- utilidades
def _norm_text(value) -> str:
    """Texto en minúsculas, sin tildes ni espacios sobrantes (para comparar etiquetas)."""
    if value is None:
        return ""
    s = unicodedata.normalize("NFKD", str(value))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.split()).casefold()


def normalize_req(value, prefix: str = DEFAULT_LAYOUT.req_prefix) -> Optional[str]:
    """Devuelve el código de requisición normalizado ("REQUI-0001") o None."""
    if value is None:
        return None
    s = str(value).strip().upper()
    if not s.startswith(prefix.upper()) or len(s) <= len(prefix):
        return None
    if re.search(r"\s", s):
        return None
    return s


def to_number(value) -> Optional[float]:
    """Número de la celda; acepta texto numérico ("1.234,5", "12"). None si no es número."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip().replace(" ", "")
        if not s:
            return None
        if "," in s and "." in s:                      # 1.234,5 -> 1234.5
            s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return None


_DATE_FORMATS = ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%d/%m/%Y %H:%M:%S",
                 "%d/%m/%Y %H:%M", "%d-%m-%Y")


def to_date(value):
    """Convierte la celda a datetime.date. Si es texto no reconocible lo devuelve tal cual.

    Devuelve None si la celda está vacía.
    """
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            converted = from_excel(value)
            return converted.date() if isinstance(converted, dt.datetime) else converted
        except Exception:  # noqa: BLE001 - número fuera de rango
            return str(value)
    s = str(value).strip()
    for fmt in _DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return s


def classify_row(bought, received, pending) -> Optional[str]:
    """Semáforo de la v1. Gana el último que aplica: rojo > verde > amarillo.

    * amarillo  (sin_oc):       comprada == 0
    * verde     (sin_entrada):  comprada != 0 y entrada == 0
    * rojo      (pendiente):    entrada != 0 y pendiente != 0 (ambos numéricos)
    """
    m, n, r = to_number(bought), to_number(received), to_number(pending)
    status = None
    if m == 0:
        status = "sin_oc"
    if m != 0 and n == 0:
        status = "sin_entrada"
    if n not in (None, 0) and r not in (None, 0):
        status = "pendiente"
    return status


def _check_extension(path: Path, role: str) -> None:
    ext = path.suffix.lower()
    if ext == ".xls":
        raise KardexError(f"El archivo de {role} está en formato antiguo (.xls).",
                          "Ábrelo en Excel y usa «Guardar como» → Libro de Excel (.xlsx).")
    if ext not in (".xlsx", ".xlsm"):
        raise KardexError(f"El archivo de {role} no es un libro de Excel (.xlsx).")


def _open_workbook(path: Path, role: str, **kwargs):
    if not path.exists():
        raise KardexError(f"No se encontró el archivo de {role}:\n{path}")
    _check_extension(path, role)
    try:
        return openpyxl.load_workbook(path, **kwargs)
    except PermissionError as exc:
        raise KardexError(f"No se pudo abrir el archivo de {role} porque está en uso.",
                          "Ciérralo en Excel e inténtalo de nuevo.") from exc
    except Exception as exc:  # noqa: BLE001
        raise KardexError(f"El archivo de {role} no se pudo leer: parece dañado o no es un Excel válido.",
                          str(exc)) from exc


def is_file_locked(path: Path) -> bool:
    """True si otro programa (normalmente Excel) tiene el archivo abierto en exclusiva."""
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def default_output_path(kardex_path: Path, now: Optional[dt.datetime] = None) -> Path:
    now = now or dt.datetime.now()
    return kardex_path.with_name(f"{kardex_path.stem} - procesado {now:%Y-%m-%d %H%M}.xlsx")


# --------------------------------------------------------------------------- escaneos rápidos
def scan_requisitions(path: Path, layout: Layout = DEFAULT_LAYOUT,
                      cancel: Optional[threading.Event] = None) -> RequisitionScan:
    """Lee el reporte de requisiciones y devuelve {código: fecha}."""
    path = Path(path)
    wb = _open_workbook(path, "requisiciones", read_only=True, data_only=True)
    scan = RequisitionScan(path=path, dates={})
    code_i, date_i = layout.req_code_col - 1, layout.req_date_col - 1
    hi = max(layout.req_code_col, layout.req_date_col)
    offset = layout.req_date_row_offset

    def record(code, raw):
        date = to_date(raw)
        if date is None:
            scan.codes_without_date.append(code)
        else:
            if isinstance(date, str):
                scan.unparsed_dates += 1
            scan.dates[code] = date

    # Lectura en streaming: nunca se carga la hoja completa en memoria.
    pending: Optional[tuple[str, int]] = None   # (código, fila donde está su fecha)
    try:
        for n, row in enumerate(wb.active.iter_rows(min_col=1, max_col=hi, values_only=True)):
            scan.rows = n + 1
            if n % 2000 == 0 and cancel is not None and cancel.is_set():
                raise Cancelled()
            if pending is not None:
                if n < pending[1]:
                    continue                       # filas intermedias: se saltan, como la v1
                record(pending[0], row[date_i] if date_i < len(row) else None)
                pending = None
                continue
            code = normalize_req(row[code_i] if code_i < len(row) else None, layout.req_prefix)
            if code is not None:
                pending = (code, n + offset)
        if pending is not None:                    # código en la última fila, sin fecha
            record(pending[0], None)
    finally:
        wb.close()
    return scan


def scan_kardex(path: Path, layout: Layout = DEFAULT_LAYOUT) -> KardexScan:
    """Conteo rápido (solo lectura) para validar que el archivo parece un Kardex."""
    path = Path(path)
    wb = _open_workbook(path, "kardex", read_only=True, data_only=True)
    try:
        ws = wb.active
        hi = max(layout.kx_label_col, layout.kx_req_col)
        rows = groups = reqs = 0
        for row in ws.iter_rows(min_col=1, max_col=hi, values_only=True):
            rows += 1
            if _norm_text(row[layout.kx_label_col - 1]).startswith(GROUP_LABEL):
                groups += 1
            if normalize_req(row[layout.kx_req_col - 1], layout.req_prefix):
                reqs += 1
    finally:
        wb.close()
    return KardexScan(path=path, rows=rows, groups=groups, requisitions=reqs)


def guess_file_role(path: Path, layout: Layout = DEFAULT_LAYOUT) -> Optional[str]:
    """'requisiciones', 'kardex' o None. Sirve para avisar si se intercambiaron los archivos."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:  # noqa: BLE001
        return None
    try:
        ws = wb.active
        req_hits = kx_hits = 0
        for n, row in enumerate(ws.iter_rows(min_col=1, max_col=layout.kx_req_col, values_only=True)):
            if n > 400:
                break
            if _norm_text(row[layout.kx_label_col - 1]).startswith(GROUP_LABEL):
                kx_hits += 1
            if normalize_req(row[layout.req_code_col - 1], layout.req_prefix):
                req_hits += 1
        if kx_hits:
            return "kardex"
        if req_hits:
            return "requisiciones"
        return None
    finally:
        wb.close()


# --------------------------------------------------------------------------- proceso principal
def process_kardex(
    requisitions_path: Path,
    kardex_path: Path,
    output_path: Optional[Path] = None,
    *,
    overwrite: bool = False,
    group_scoped_filldown: bool = True,
    layout: Layout = DEFAULT_LAYOUT,
    progress: Optional[ProgressFn] = None,
    cancel: Optional[threading.Event] = None,
) -> KardexResult:
    """Cruza requisiciones con el kardex, aplica formato/semáforo y guarda el resultado.

    overwrite=True  -> guarda sobre el kardex original (antes crea una copia de respaldo).
    overwrite=False -> guarda en `output_path` (o un nombre nuevo junto al original).
    group_scoped_filldown -> la fecha solo se arrastra dentro de la misma orden
                             (la v1 la arrastraba también a la orden siguiente).
    """
    t0 = time.perf_counter()
    report = progress or (lambda _m, _f: None)
    requisitions_path, kardex_path = Path(requisitions_path), Path(kardex_path)

    def check_cancel():
        if cancel is not None and cancel.is_set():
            raise Cancelled()

    if requisitions_path.resolve() == kardex_path.resolve():
        raise KardexError("Seleccionaste el mismo archivo para requisiciones y para kardex.")

    target = kardex_path if overwrite else Path(output_path or default_output_path(kardex_path))
    if target.exists() and is_file_locked(target):
        raise KardexError(f"«{target.name}» está abierto en otro programa.",
                          "Cierra el archivo en Excel e inténtalo de nuevo.")

    # 1) Requisiciones ---------------------------------------------------------
    report("Leyendo requisiciones…", None)
    req = scan_requisitions(requisitions_path, layout, cancel)
    if not req.dates:
        raise KardexError(
            f"No se encontró ninguna requisición con fecha en «{requisitions_path.name}».",
            f"Se esperan códigos «{layout.req_prefix}…» en la columna "
            f"{get_column_letter(layout.req_code_col)} y la fecha en la columna "
            f"{get_column_letter(layout.req_date_col)} de la fila siguiente. "
            "¿Seleccionaste el archivo correcto?")
    log.info("Requisiciones con fecha: %d (sin fecha: %d)", len(req.dates), len(req.codes_without_date))
    check_cancel()

    # 2) Kardex ----------------------------------------------------------------
    report("Abriendo kardex (puede tardar en archivos grandes)…", None)
    wb = _open_workbook(kardex_path, "kardex")
    try:
        ws = wb.active
        last_row = ws.max_row
        max_col = ws.max_column
        if last_row < layout.kx_first_row:
            raise KardexError(f"El kardex «{kardex_path.name}» está vacío.")

        date_style_font = Font(name="Courier New", size=8)
        center = Alignment(horizontal="center", vertical="center")
        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_font = Font(name="Calibri", size=11, bold=True)
        fills = {k: PatternFill(start_color=c, end_color=c, fill_type="solid")
                 for k, c in STATUS_COLORS.items()}

        dates_written = filled_down = 0
        missing: dict[str, None] = {}
        counts = {k: 0 for k in STATUS_COLORS}
        carry = None
        total = last_row - layout.kx_first_row + 1
        step = max(1, total // 200)

        dc = layout.kx_date_col
        for idx, cells in enumerate(ws.iter_rows(min_row=layout.kx_first_row, max_row=last_row,
                                                 max_col=max(max_col, dc))):
            r = layout.kx_first_row + idx
            if idx % step == 0:
                check_cancel()
                report(f"Procesando fila {r:,} de {last_row:,}…".replace(",", "."), idx / total)

            label = _norm_text(cells[layout.kx_label_col - 1].value)
            date_cell = cells[dc - 1]

            # a) Fecha desde la requisición
            code = normalize_req(cells[layout.kx_req_col - 1].value, layout.req_prefix)
            if code is not None:
                found = req.dates.get(code)
                if found is not None:
                    date_cell.value = found
                    dates_written += 1
                else:
                    missing[code] = None

            # b) Encabezados de grupo
            if label.startswith(GROUP_LABEL):
                if group_scoped_filldown:
                    carry = date_cell.value if code is not None and code in req.dates else None
                else:
                    carry = date_cell.value or carry
                date_cell.value = None
                continue
            if label == HEADER_LABEL:
                date_cell.value = "Fecha"
                date_cell.font = header_font
                date_cell.alignment = center
                if r >= layout.kx_style_from_row:
                    date_cell.border = border
                continue

            # c) Arrastrar fecha hacia abajo (desde la fila 3, como la v1)
            if r > layout.kx_first_row:
                if date_cell.value in (None, ""):
                    if carry is not None:
                        date_cell.value = carry
                        filled_down += 1
                else:
                    carry = date_cell.value
            elif date_cell.value not in (None, ""):
                carry = date_cell.value

            if isinstance(date_cell.value, (dt.date, dt.datetime)):
                date_cell.number_format = "mm/dd/yyyy"

            # d) Formato de la columna de fecha
            if r >= layout.kx_style_from_row:
                date_cell.font = date_style_font
                date_cell.alignment = center
                date_cell.border = border

            # e) Semáforo
            if r >= layout.kx_color_from_row:
                status = classify_row(cells[layout.kx_qty_bought_col - 1].value,
                                      cells[layout.kx_qty_received_col - 1].value,
                                      cells[layout.kx_pending_col - 1].value)
                if status:
                    counts[status] += 1
                    fill = fills[status]
                    for c in cells[:max_col]:
                        c.fill = fill

        # 3) Leyenda -----------------------------------------------------------
        for offset, key in enumerate(("sin_oc", "sin_entrada", "pendiente")):
            c = ws.cell(row=layout.legend_first_row + offset, column=layout.legend_col)
            c.value = STATUS_LABELS[key]
            c.font = header_font
            c.alignment = center
            c.fill = fills[key]
        ws.column_dimensions[get_column_letter(layout.legend_col)].width = 25
        ws.column_dimensions[get_column_letter(dc)].width = max(
            ws.column_dimensions[get_column_letter(dc)].width or 0, 12)

        check_cancel()

        # 4) Guardar de forma segura (archivo temporal + reemplazo atómico) ------
        report("Guardando resultado…", 1.0)
        backup = None
        if overwrite:
            backup = kardex_path.with_name(
                f"{kardex_path.stem} - respaldo {dt.datetime.now():%Y-%m-%d %H%M%S}{kardex_path.suffix}")
            shutil.copy2(kardex_path, backup)
        _atomic_save(wb, target)
    finally:
        wb.close()

    result = KardexResult(
        output_path=target, backup_path=backup, requisitions_loaded=len(req.dates),
        kardex_rows=total, dates_written=dates_written, rows_filled_down=filled_down,
        missing_requisitions=sorted(missing), status_counts=counts,
        elapsed_s=time.perf_counter() - t0)
    log.info("Proceso terminado: %s", result)
    return result


def _atomic_save(wb, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".kardexpress-", suffix=".xlsx", dir=target.parent)
    os.close(fd)
    try:
        wb.save(tmp)
        try:
            os.replace(tmp, target)
        except PermissionError as exc:
            raise KardexError(f"No se pudo guardar «{target.name}» porque está abierto en otro programa.",
                              "Cierra el archivo en Excel e inténtalo de nuevo.") from exc
    except PermissionError as exc:
        raise KardexError(f"No hay permiso para escribir en la carpeta:\n{target.parent}",
                          "Elige otra carpeta de destino.") from exc
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
