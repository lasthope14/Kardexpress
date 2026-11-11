import datetime as dt
import threading

import openpyxl
import pytest
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from kardexpress.core import (Cancelled, KardexError, classify_row, guess_file_role, normalize_req,
                              process_kardex, scan_kardex, scan_requisitions, to_date, to_number)

D1, D2, D3 = dt.datetime(2023, 3, 1), dt.datetime(2023, 4, 15), dt.datetime(2023, 5, 20)


# --------------------------------------------------------------------------- fixtures
def make_requisitions(path, entries):
    """entries: [(codigo, fecha)] -> código en B, fecha en W de la fila siguiente."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=2, value="Reporte de requisiciones")
    r = 3
    for code, date in entries:
        ws.cell(row=r, column=2, value=code)
        ws.cell(row=r + 1, column=23, value=date)
        r += 3
    wb.save(path)
    return path


def make_kardex(path, groups):
    """groups: [(req_en_fila_grupo, [(M, N, R, req_en_item)])]"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="KARDEX DE COMPRAS")
    r = 6
    for req, items in groups:
        ws.cell(row=r, column=2, value="Consecutivo General 00%d" % r)
        ws.cell(row=r, column=5, value=req)
        r += 1
        ws.cell(row=r, column=2, value="Código")
        ws.cell(row=r, column=13, value="Cant. comprada")
        r += 1
        for m, n, pend, item_req in items:
            ws.cell(row=r, column=2, value=f"ART-{r}")
            ws.cell(row=r, column=5, value=item_req)
            ws.cell(row=r, column=13, value=m)
            ws.cell(row=r, column=14, value=n)
            ws.cell(row=r, column=18, value=pend)
            r += 1
    wb.save(path)
    return path


@pytest.fixture
def files(tmp_path):
    req = make_requisitions(tmp_path / "req.xlsx", [("REQUI-0001", D1), ("REQUI-0002", D2), ("REQUI-0003", D3)])
    kx = make_kardex(tmp_path / "kardex.xlsx", [
        ("REQUI-0001", [(0, 0, 0, None), (5, 0, 5, None), (5, 2, 3, None)]),
        ("REQUI-0009", [(5, 5, 0, None)]),                      # requisición inexistente
        ("REQUI-0003", [(3, 3, 0, None), (0, None, None, None)]),
    ])
    return req, kx


# --------------------------------------------------------------------------- utilidades
def test_normalize_req():
    assert normalize_req(" requi-0001 ") == "REQUI-0001"
    assert normalize_req("REQUI-12345") == "REQUI-12345"
    assert normalize_req("REQUI-") is None
    assert normalize_req(1234) is None
    assert normalize_req(None) is None


def test_to_number():
    assert to_number(0) == 0
    assert to_number("12") == 12
    assert to_number("1.234,5") == 1234.5
    assert to_number("abc") is None
    assert to_number(None) is None
    assert to_number(True) is None


def test_to_date():
    assert to_date(D1) == D1.date()
    assert to_date(45000) == dt.date(2023, 3, 15)
    assert to_date("2023-03-01") == dt.date(2023, 3, 1)
    assert to_date("no es fecha") == "no es fecha"
    assert to_date("  ") is None


@pytest.mark.parametrize("m,n,r,expected", [
    (0, 0, 0, "sin_oc"),
    (5, 0, 5, "sin_entrada"),
    (5, 2, 3, "pendiente"),
    (5, 5, 0, None),
    (0, 2, 3, "pendiente"),            # rojo gana a amarillo, como en la v1
    ("0", "x", None, "sin_oc"),        # número guardado como texto
    (None, None, None, None),
])
def test_classify(m, n, r, expected):
    assert classify_row(m, n, r) == expected


# --------------------------------------------------------------------------- escaneos
def test_scan_requisitions(files):
    scan = scan_requisitions(files[0])
    assert scan.dates == {"REQUI-0001": D1.date(), "REQUI-0002": D2.date(), "REQUI-0003": D3.date()}


def test_scan_kardex_and_guess(files):
    scan = scan_kardex(files[1])
    assert scan.groups == 3 and scan.requisitions == 3
    assert guess_file_role(files[0]) == "requisiciones"
    assert guess_file_role(files[1]) == "kardex"


# --------------------------------------------------------------------------- proceso
def _fill(ws, r):
    c = ws.cell(row=r, column=1).fill
    return c.start_color.rgb[-6:] if c and c.fill_type == "solid" else None


def test_process_copy(files, tmp_path):
    req, kx = files
    before = kx.read_bytes()
    out = tmp_path / "salida.xlsx"
    res = process_kardex(req, kx, out)
    assert kx.read_bytes() == before, "el original no debe tocarse"
    assert res.output_path == out and res.backup_path is None
    assert res.missing_requisitions == ["REQUI-0009"]
    assert res.status_counts == {"sin_oc": 2, "sin_entrada": 1, "pendiente": 1}

    ws = openpyxl.load_workbook(out).active
    # grupo 1: filas 6 (consecutivo), 7 (código), 8..10 items
    assert ws.cell(row=6, column=19).value is None
    assert ws.cell(row=7, column=19).value == "Fecha"
    for r in (8, 9, 10):
        assert ws.cell(row=r, column=19).value == D1
    assert [_fill(ws, r) for r in (8, 9, 10)] == ["FFFF00", "00FF00", "FF0000"]
    # grupo 2 (REQUI-0009 inexistente): sin fecha, NO hereda la del grupo anterior
    assert ws.cell(row=13, column=19).value is None
    # grupo 3
    assert ws.cell(row=16, column=19).value == D3
    assert ws.cell(row=17, column=19).value == D3
    assert ws.cell(row=3, column=20).value == "SIN ORDEN DE COMPRA"


def test_classic_filldown_matches_v1(files, tmp_path):
    req, kx = files
    out = tmp_path / "clasico.xlsx"
    process_kardex(req, kx, out, group_scoped_filldown=False)
    ws = openpyxl.load_workbook(out).active
    assert ws.cell(row=13, column=19).value == D1   # la v1 arrastraba la fecha anterior


def test_overwrite_creates_backup(files):
    req, kx = files
    before = kx.read_bytes()
    res = process_kardex(req, kx, overwrite=True)
    assert res.output_path == kx
    assert res.backup_path is not None and res.backup_path.read_bytes() == before


def test_cancel(files, tmp_path):
    ev = threading.Event()
    ev.set()
    out = tmp_path / "x.xlsx"
    with pytest.raises(Cancelled):
        process_kardex(files[0], files[1], out, cancel=ev)
    assert not out.exists()


def test_errors(files, tmp_path):
    req, kx = files
    with pytest.raises(KardexError):
        process_kardex(req, req, tmp_path / "x.xlsx")              # mismo archivo
    with pytest.raises(KardexError, match="requisición"):
        process_kardex(kx, req, tmp_path / "x.xlsx")               # intercambiados
    bad = tmp_path / "viejo.xls"
    bad.write_bytes(b"xx")
    with pytest.raises(KardexError, match=".xls"):
        process_kardex(req, bad, tmp_path / "x.xlsx")
    corrupt = tmp_path / "roto.xlsx"
    corrupt.write_bytes(b"no soy un zip")
    with pytest.raises(KardexError, match="dañado"):
        process_kardex(req, corrupt, tmp_path / "x.xlsx")


def test_dirty_data_does_not_crash(tmp_path):
    """La v1 fallaba con números en la columna B o fechas como texto."""
    req = tmp_path / "req.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=2, value=12345)                    # número en B (v1: AttributeError)
    ws.cell(row=2, column=2, value="REQUI-0001")
    ws.cell(row=3, column=23, value="2023-03-01")            # fecha como texto (v1: AttributeError)
    wb.save(req)
    kx = make_kardex(tmp_path / "k.xlsx", [(" requi-0001 ", [(1, "abc", None, None)])])
    res = process_kardex(req, kx, tmp_path / "o.xlsx")
    assert res.dates_written == 1


# --------------------------------------------------------------------------- equivalencia con la v1
def run_v1(path_1, path_2):
    """Lógica de run_script() de la v1 (docs/kardex_v1_reconstruido.py), sin la UI."""
    sheet = openpyxl.load_workbook(path_1).active
    result_data, row_index = [], 1
    while row_index <= sheet.max_row:
        v = sheet.cell(row=row_index, column=2).value
        if v and v.startswith("REQUI-") and len(v) == 10:
            d = sheet.cell(row=row_index + 1, column=23).value
            if d:
                result_data.append((v, d.strftime("%m/%d/%Y")))
            row_index += 2
        else:
            row_index += 1
    dates = dict(result_data)
    wb = openpyxl.load_workbook(path_2)
    s = wb.active
    last = s.max_row
    for i in range(2, last + 1):
        d = dates.get(s.cell(row=i, column=5).value)
        if d:
            s.cell(row=i, column=19).value = d
    for i in range(3, last + 1):
        if not s.cell(row=i, column=19).value:
            s.cell(row=i, column=19).value = s.cell(row=i - 1, column=19).value
    for i in range(1, last + 1):
        v = s.cell(row=i, column=2).value
        if v and v.startswith("Consecutivo General"):
            s.cell(row=i, column=19).value = None
            continue
        if v and v == "Código":
            s.cell(row=i, column=19).value = "Fecha"
    colors = {}
    for i in range(8, last + 1):
        m, n, r = (s.cell(row=i, column=c).value for c in (13, 14, 18))
        if m == 0:
            colors[i] = "FFFF00"
        if m != 0 and n == 0:
            colors[i] = "00FF00"
        if (n != 0 and n is not None and not isinstance(n, str)
                and r != 0 and r is not None and not isinstance(r, str)):
            colors[i] = "FF0000"
    return s, colors


def test_equivalence_with_v1(files, tmp_path):
    req, kx = files
    v1_sheet, v1_colors = run_v1(req, kx)
    out = tmp_path / "v2.xlsx"
    process_kardex(req, kx, out, group_scoped_filldown=False)
    ws = openpyxl.load_workbook(out).active
    for r in range(1, v1_sheet.max_row + 1):
        v1 = v1_sheet.cell(row=r, column=19).value
        v2 = ws.cell(row=r, column=19).value
        if isinstance(v2, dt.datetime):
            v2 = v2.strftime("%m/%d/%Y")             # v2 guarda fecha real; v1 guardaba texto
        assert v1 == v2, f"fila {r}: v1={v1!r} v2={v2!r}"
        assert v1_colors.get(r) == _fill(ws, r), f"color fila {r}"
