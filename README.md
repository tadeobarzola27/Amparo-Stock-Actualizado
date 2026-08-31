# Amparo · Inventario semanal

Dashboard en Streamlit que lee automáticamente una **Hoja de Google central**
administrada por la oficina (solapa "Base Articulos") y muestra el stock
actualizado, con filtros por Rubro, Familia, Marca y Proveedor. Las tiendas
consultan la información desde el enlace publicado y no pueden subir archivos.

## Configurar el archivo central

La aplicación ya tiene configurada la Hoja de Google central. La planilla debe
estar compartida como **Cualquier persona con el enlace: Lector** para que el
servidor pueda descargarla.

Opcionalmente, se puede reemplazar la fuente sin modificar el código desde
Streamlit Community Cloud, en **App settings → Secrets**:

```toml
DATA_FILE_URL = "https://docs.google.com/spreadsheets/d/ID_DE_LA_PLANILLA/edit"
```

La aplicación vuelve a consultar el archivo automáticamente cada cinco minutos.
También ofrece un botón **Actualizar datos** para forzar una lectura inmediata.
Para actualizar el inventario, la oficina edita la Hoja de Google central.

## Cómo correrlo

```bash
pip install -r requirements.txt
streamlit run amparo_dashboard.py
```

Se abre en el navegador (`http://localhost:8501`) y toma los datos del archivo
central configurado en `DATA_FILE_URL`.

Cada semana: la oficina actualiza los movimientos en la solapa "Ingresos y
Egresos semanal" de la Hoja de Google central. El dashboard publicado mostrará
la nueva información sin que las tiendas deban subir nada.

## Qué se automatizó en el Excel

En `Base_articulos.xlsx`, las columnas **F (Ingresos)** y **G (Egresos)**
de la solapa "Base Articulos" ahora son fórmulas (`SUMIFS`) que buscan cada
Código en "Ingresos y Egresos semanal" y traen el movimiento de esa semana
automáticamente. La columna **H (StkActual)**, que ya tenía la fórmula
`=E+F-G`, se recalcula sola con eso. No hace falta tocar nada a mano salvo
cargar la semana en "Ingresos y Egresos semanal".

La columna **E (StkInicial)** sigue siendo de carga manual — no se
automatizó porque no se pidió y porque es la que define el punto de
partida del stock (ver más abajo).

## Dos cosas para revisar en tus datos (no son errores del dashboard)

1. **5 códigos con movimientos que no están en tu catálogo**: 3072, 393,
   6690862, 6691866 y 604 registraron Ingresos/Egresos en la carga semanal
   pero no existen en "Base Articulos", así que ese movimiento (4589
   unidades de Egresos en total) no se descontó de ningún stock. Están
   detallados en la solapa nueva **"Códigos a revisar"** del Excel y en un
   aviso dentro del dashboard. Para que se sumen solos, agregalos como fila
   nueva en "Base Articulos".

2. **StkInicial (columna E) está en 0 para los 3.621 productos**: hoy no
   hay ningún punto de partida cargado, así que StkActual (H) es solo el
   neto de Ingresos menos Egresos desde que se empezó a cargar la solapa
   semanal — no el stock físico real (por eso ves tantos productos "Sin
   stock" o en negativo). En cuanto cargues el stock inicial real de cada
   producto en la columna E, H va a quedar exacto.

## Qué se ve en el dashboard (y qué no)

Columnas mostradas, en el mismo orden que en el Excel: **Rubro, Código,
Descripción, UxB, Stock (StkActual/H), Familia, Subfamilia, Marca,
Proveedor** — más un indicador de **Estado** (OK / Bajo / Sin stock /
Negativo) calculado a partir del Stock. Las columnas E, F y G (StkInicial,
Ingresos, Egresos) no se muestran en la tabla principal, tal como se pidió;
sólo aparecen, por separado, en el aviso de "códigos a revisar" — porque
ese aviso es justamente sobre movimientos sin capturar.

## Archivos

- `amparo_dashboard.py` — la app (interfaz).
- `amparo_logic.py` — la lógica de lectura y cálculo, sin Streamlit, para
  poder testearla sola.
- `test_amparo_logic.py` — pruebas contra el archivo real (43 checks,
  incluye los dos hallazgos de arriba). Correr con `python3
  test_amparo_logic.py`.
- `add_automation.py` / `add_orphan_check.py` — los scripts que se usaron
  para automatizar F/G y agregar el chequeo de códigos a
  `Base_articulos.xlsx`. Quedan acá como referencia; no hace falta
  volver a correrlos.
