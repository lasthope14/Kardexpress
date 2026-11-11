# Kardexpress 2.0

Cruza el **reporte de requisiciones** con el **kardex de compras**. Escribe en el kardex la fecha de creación de cada requisición y colorea cada línea según su estado.

| Color | Estado | Regla (columnas del kardex) |
|---|---|---|
| Amarillo | Sin orden de compra | M (comprada) = 0 |
| Verde | Sin entrada/remisión | M ≠ 0 y N (entrada) = 0 |
| Rojo | Pendiente proveedor | N ≠ 0 y R (pendiente) ≠ 0 |

## Qué entrega (v2.1)

Por defecto genera un **reporte nuevo** y no modifica el kardex:

| Pestaña | Contenido |
|---|---|
| **Resumen** | En lenguaje sencillo: cuántas líneas hay en cada estado, qué significa cada uno y qué hay que hacer. También incluye un gráfico y la lista de requisiciones que quedaron sin fecha. |
| **Pendiente proveedor** · **Sin entrada** · **Sin orden de compra** | Una lista de trabajo por estado. Cada pestaña lleva el color de su estado. |
| **Todas las líneas** | El kardex completo como tabla plana, con filtros (incluido *Filtrar por color*), encabezado fijo y la columna *Estado* siempre visible. |

Cada línea incluye *Estado*, *Días desde la requisición*, *Fecha requisición* (formato `15-mar-2023`), *Requisición*, *Orden* y las columnas originales del kardex, con los títulos que tienen en el archivo. Está preparado para imprimir en horizontal, ajustado al ancho de la página.

El modo anterior (kardex original pintado con colores, en una copia o sobrescribiendo el archivo) sigue disponible en el paso 3.

### Archivos grandes

El reporte se lee y se escribe por partes (streaming), así que la memoria no crece con el tamaño del archivo. Con un kardex de 192.000 filas, los dos modos se comparan así:

| | Kardex pintado (modo anterior) | Reporte (v2.1) |
|---|---|---|
| Memoria | ~1.500 MB | ~20 MB |
| Tiempo | 42 s | 38 s |
| Interfaz | Se traba hasta 1 s | Fluida |

## Formato que espera

- **Requisiciones** (hoja activa): el código `REQUI-…` está en la columna **B** y su fecha en la columna **W** de la fila siguiente.
- **Kardex** (hoja activa): está agrupado por filas `Consecutivo General…` y `Código` en la columna **B**. La requisición va en la columna **E**. La fecha se escribe en la columna **S** y la leyenda en **T3:T5**.

Estas posiciones se definen en `Layout`, dentro de [kardexpress/core.py](kardexpress/core.py).

## Uso

- **Instalar en un equipo:** basta con enviar el archivo `dist\Kardexpress-Instalador-<versión>.exe` y abrirlo con doble clic.
  - No pide permisos de administrador. Se instala en `%LOCALAPPDATA%\Programs\Kardexpress`.
  - Crea un acceso en el menú Inicio y, si se marca la opción, otro en el escritorio.
  - Para desinstalar: *Configuración → Aplicaciones*.
  - **Actualizar:** se ejecuta el instalador de la versión nueva encima del anterior. Las preferencias se conservan.
  - El instalador no está firmado digitalmente. Windows puede mostrar «Windows protegió su PC»; en ese caso se elige *Más información → Ejecutar de todas formas*.
- **Desde el código:**
  ```powershell
  .\.venv\Scripts\python.exe run.py
  ```
- **Pruebas:**
  ```powershell
  .\.venv\Scripts\python.exe -m pytest -q
  ```
- **Generar una versión nueva:**
  1. Sube el número en `kardexpress/__init__.py`.
  2. Ejecuta:
     ```powershell
     powershell -ExecutionPolicy Bypass -File build.ps1
     ```
  El script ejecuta las pruebas, crea el ejecutable y luego el instalador. Si alguna prueba falla, se detiene. Requiere [Inno Setup 6](https://jrsoftware.org/isinfo.php), que se instala con `winget install JRSoftware.InnoSetup`.

Las preferencias se guardan en `%APPDATA%\Kardexpress`. El registro de actividad está en `%LOCALAPPDATA%\Kardexpress\logs`.

## Cambios respecto a la v1

La v1 se reconstruyó a partir de `kardex.exe`. El código recuperado está en [docs/kardex_v1_reconstruido.py](docs/kardex_v1_reconstruido.py).

| Problema en la v1 | En la v2 |
|---|---|
| Si ocurría cualquier error (archivo abierto en Excel, celda numérica en B, fecha como texto), el hilo moría y la barra giraba para siempre | Todos los errores llegan a la interfaz con un mensaje claro y el botón vuelve a activarse |
| Desde la segunda ejecución nunca avisaba que había terminado: el evento no se reiniciaba y dejaba de comprobarlo | Cada ejecución tiene su propio ciclo de vida |
| Si faltaba un archivo, la interfaz quedaba bloqueada sin decir nada | El botón solo se activa cuando ambos archivos son válidos |
| Sobrescribía el kardex sin respaldo | Por defecto guarda una copia nueva; si eliges sobrescribir, primero crea un respaldo. El guardado es atómico |
| La barra de progreso era una animación falsa | Muestra el progreso real por fila y permite cancelar |
| Recorría la hoja 6 veces | Hace una sola pasada (40.000 filas en unos 8 s) |
| La fecha de una orden sin requisición encontrada se heredaba de la orden anterior (dato incorrecto) | La fecha solo se arrastra dentro de cada orden. Puedes desactivarlo para obtener el comportamiento de la v1 |
| No informaba qué requisiciones faltaban | Muestra un resumen con los conteos y la lista de requisiciones no encontradas, que se puede copiar |
| Sin validación de archivos | Valida los archivos al seleccionarlos y detecta si están intercambiados |

Otras diferencias: la fecha se guarda como fecha real de Excel con formato `mm/dd/yyyy` (la v1 la guardaba como texto con la misma apariencia). Además, se aceptan números guardados como texto y códigos con espacios o en minúsculas.
