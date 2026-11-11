import datetime as dt
import threading

import openpyxl
import pytest

from kardexpress.core import Cancelled, KardexError
from kardexpress.report import ACTION_STATES, ALL_SHEET, STATES, build_report, row_state

from .test_core import D1, D3, make_kardex, make_requisitions, run_v1

TODAY = dt.date(2023, 6, 1)


@pytest.fixture
def files(tmp_path):
    req = make_requisitions(tmp_path / "req.xlsx", [("REQUI-0001", D1), ("REQUI-0003", D3)])
    kx = make_kardex(tmp_path / "kardex.xlsx", [
        ("REQUI-0001", [(0, 0, 0, None), (5, 0, 5, None), (5, 2, 3, None), (5, 5, 0, None)]),
        ("REQUI-0009", [(5, 0, 5, None)]),
        ("REQUI-0003", [(3, 2.5, 0.5, None)]),
    ])
    return req, kx


def test_row_state():
    assert row_state(5, 5, 0) == "completo"
    assert row_state(None, None, None) == "sin_datos"
    assert row_state(5, 2, 3) == "pendiente"


def test_report_structure(files, tmp_path):
    req, kx = files
    before = kx.read_bytes()
    out = tmp_path / "reporte.xlsx"
    res = build_report(req, kx, out, today=TODAY)
    assert kx.read_bytes() == before
    assert res.status_counts == {"pendiente": 2, "sin_entrada": 2, "sin_oc": 1, "completo": 1, "sin_datos": 0}
    assert res.missing_requisitions == ["REQUI-0009"]

    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["Resumen", *(STATES[k].sheet for k in ACTION_STATES), ALL_SHEET]

    ws = wb[ALL_SHEET]
    head = [c.value for c in ws[1]]
    assert head[:5] == ["Estado", "Días desde la requisición", "Fecha requisición", "Requisición", "Orden"]
    assert "Código" in head and "Cant. comprada" in head
    assert ws.freeze_panes == "B2"
    assert ws.auto_filter.ref == f"A1:{openpyxl.utils.get_column_letter(len(head))}7"
    rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
    assert [r[0] for r in rows] == ["Sin orden de compra", "Sin entrada / remisión", "Pendiente proveedor",
                                   "Completo", "Sin entrada / remisión", "Pendiente proveedor"]
    assert rows[0][1] == (TODAY - D1.date()).days
    assert rows[0][2] == D1 and rows[0][3] == "REQUI-0001"
    assert rows[4][2] is None and rows[4][3] == "REQUI-0009"     # no hereda fecha de otra orden
    # La celda Estado lleva el color del estado (para «Filtrar por color»)
    assert ws["A2"].fill.start_color.rgb.endswith(STATES["sin_oc"].strong)
    # Pestañas de trabajo solo con su estado
    pend = wb[STATES["pendiente"].sheet]
    assert [r[0].value for r in pend.iter_rows(min_row=2)] == ["Pendiente proveedor"] * 2

    summary = [c for r in wb["Resumen"].iter_rows(values_only=True) for c in r if c is not None]
    assert "REQUI-0009" in " ".join(map(str, summary))


def test_report_matches_v1_colors(files, tmp_path):
    """Mismo semáforo que la v1, línea por línea."""
    req, kx = files
    _, v1_colors = run_v1(req, kx)
    color_to_state = {"FFFF00": "sin_oc", "00FF00": "sin_entrada", "FF0000": "pendiente"}
    src = openpyxl.load_workbook(kx).active
    expected = []
    for r in range(8, src.max_row + 1):
        b = src.cell(row=r, column=2).value
        if b and (str(b).startswith("Consecutivo") or b == "Código"):
            continue
        expected.append(color_to_state.get(v1_colors.get(r)))
    out = tmp_path / "r.xlsx"
    build_report(req, kx, out, today=TODAY)
    got = [r[0] for r in openpyxl.load_workbook(out)[ALL_SHEET].iter_rows(min_row=2, values_only=True)]
    title_to_state = {s.title: k for k, s in STATES.items()}
    assert [title_to_state[g] if title_to_state[g] in ACTION_STATES else None for g in got] == expected


def test_report_errors_and_cancel(files, tmp_path):
    req, kx = files
    with pytest.raises(KardexError, match="ninguna requisición"):
        build_report(kx, req, tmp_path / "x.xlsx")                      # intercambiados
    other = tmp_path / "otro.xlsx"
    other.write_bytes(req.read_bytes())
    with pytest.raises(KardexError, match="no parece un kardex"):
        build_report(req, other, tmp_path / "x.xlsx")                   # «kardex» que no lo es
    ev = threading.Event()
    ev.set()
    with pytest.raises(Cancelled):
        build_report(req, kx, tmp_path / "y.xlsx", cancel=ev)
    assert not (tmp_path / "y.xlsx").exists()
