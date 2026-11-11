# Kardexpress — notas para desarrollo

Aplicación de escritorio para Windows (Python 3.12, customtkinter, openpyxl). Hidroobras la usa en el almacén para cruzar el reporte de requisiciones con el kardex de compras. El usuario final no es técnico: los textos de la interfaz y del reporte deben ser simples, en español.

## Estructura
- `kardexpress/core.py`: lógica base, sin interfaz.
  - `Layout`: columnas de ambos reportes.
  - `scan_requisitions`: lee el reporte de requisiciones en streaming.
  - `classify_row`: el semáforo de la v1.
  - `process_kardex`: modo clásico. Pinta el kardex original y lo carga completo en memoria (pesado).
- `kardexpress/report.py`: `build_report`, el modo por defecto. Lee el kardex con `read_only` y escribe con `write_only`, así que la memoria es constante.
  - Genera las hojas Resumen, una por cada estado de `ACTION_STATES` y «Todas las líneas».
  - Los textos de cada estado están en `STATES`.
  - En `write_only`, `row_dimensions`, `freeze_panes` y `column_dimensions` se fijan **antes** de hacer `append` de la fila.
- `kardexpress/ui.py`: interfaz.
  - El trabajo pesado corre en un hilo y se comunica con la interfaz mediante `queue` y `after()`. Nunca hay que tocar widgets desde el hilo.
- `kardexpress/settings.py`: preferencias en `%APPDATA%\Kardexpress`, log en `%LOCALAPPDATA%\Kardexpress\logs`.
- `docs/kardex_v1_reconstruido.py`: código de la v1 recuperado de `kardex.exe` (solo como referencia).
- `tests/`: incluye pruebas de equivalencia con la v1 (`run_v1`). Si cambias la lógica del semáforo o de las fechas, esas pruebas deben seguir pasando, salvo que el cambio sea intencional.

## Reglas de negocio (v1)
- **Requisiciones:** el código `REQUI-…` está en la columna B y la fecha en la columna W de la fila siguiente.
- **Kardex:** cada grupo empieza con una fila «Consecutivo General» en B, seguida de la fila de títulos «Código». La requisición está en E.
- **Semáforo** (gana el último que aplique):
  - Amarillo: M = 0.
  - Verde: M ≠ 0 y N = 0.
  - Rojo: N ≠ 0 y R ≠ 0.

## Comandos
```powershell
.\.venv\Scripts\python.exe -m pytest -q          # pruebas
.\.venv\Scripts\python.exe run.py                # ejecutar la app
powershell -ExecutionPolicy Bypass -File build.ps1   # pruebas + exe + instalador (requiere Inno Setup 6)
```
Para publicar una versión, sube `__version__` en `kardexpress/__init__.py`.

## Cuidado
- No edites archivos con `Get-Content`/`Set-Content` de PowerShell 5.1: estropea las tildes. Usa herramientas que respeten UTF-8.
- No cambies el `AppId` de `installer/kardexpress.iss`: identifica la app entre versiones.
- El logo se genera con `tools/make_assets.py`.
