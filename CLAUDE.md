# CLAUDE.md

Memoria de arranque para Claude Code en este repo. Se carga sola al empezar cada sesión: léela antes de explorar o tocar nada — evita repetir exploración (tokens) y errores ya conocidos.

**Orden de lectura:**
1. Este archivo — mapa rápido, reglas, estado y registro de avances/errores.
2. `sys-a-kyte/MEMORIA_PROYECTO.md` — memoria completa: reglas de negocio detalladas, datos del inventario, decisiones de Kino, especificación de la v2 y prompt de continuación. Para cualquier cambio en la lógica de cruce, léelo entero primero.

## Qué es el proyecto

"SYS a Kyte": app que sube el inventario del SYS (Excel), cruza por SKU contra el catálogo de Kyte y devuelve los 4 archivos de Kyte actualizados (stock + precio) más un reporte de cambios. Reemplaza un flujo manual con XLOOKUP.

Corre como Claude Artifact (un HTML + `tabla.js`, sin backend) **y también** como app Node/Express desplegable en Railway (agregado 21 sep 2026 — ver "Estado actual" abajo).

Idioma de trabajo: **español**. El usuario (Kino) escribe informal y rápido; prefiere respuestas cortas, directas y prácticas.

## Flujo de git

**Se trabaja directo sobre `main`. Sin ramas por feature.** (Decisión de Kino, 21 sep 2026: quiere ver cada cambio en vivo — probablemente porque Railway despliega desde `main` — y prefiere verificar él mismo en vez de pasar por rama+PR.) Commitear y pushear a `main` cuando el cambio esté probado localmente (correr `test/run.py` antes si se tocó la lógica de cruce). No crear ramas `claude/...` salvo que Kino lo pida para algo puntual.

## Reglas que no se rompen (resumen — detalle en MEMORIA_PROYECTO.md §0)

1. El SYS manda: es la fuente de verdad de stock y precio base.
2. Stock actual de Kyte = existencia del SYS (negativo → 0; blanco en Kyte + 0 en SYS se deja en blanco).
3. Precio\* = 2 × precio del SYS (con casilla para apagarlo). Pisa precios propios de Kyte.
4. Lo que no cruza no se toca — la app nunca crea, borra ni oculta productos por su cuenta.
5. Cruce flexible (por formato o por nombre) solo si hay una única coincidencia posible.
6. SKU nuevos del SYS: nunca se crean solos, solo se avisan y se listan.
7. **Antes de tocar código, correr `sys-a-kyte/test/run.py`**; debe terminar en `TODO OK`.
8. No cambiar estas reglas ni las decisiones de la v2 (MEMORIA_PROYECTO.md §4) sin preguntarle a Kino.

## Arquitectura

```
server.js, package.json        servidor Express (deploy Railway) — sirve sys-a-kyte/app/
sys-a-kyte/
├── MEMORIA_PROYECTO.md        memoria completa del proyecto
├── app/
│   ├── sys-a-kyte.html        FUENTE — motor de negocio completo, se publica como Claude Artifact
│   ├── index.html             generado de sys-a-kyte.html vía tools/build_index.py
│   └── tabla.js                tabla base de Kyte (datos): window.__TABLA__
├── tools/
│   ├── build_index.py         sys-a-kyte.html → index.html
│   └── generar_tabla.py       Excel formato Kyte → tabla.js
├── test/run.py                 test end-to-end (Playwright) contra un cálculo independiente
└── datos/                      Excel de ejemplo (SYS + hojas Kyte, 18 sep 2026)
```

Toda la lógica vive en `sys-a-kyte/app/sys-a-kyte.html` (funciones clave: `leerSys`, `leerTabla`, `procesar`, `libroHoja`/`libroReporte`, estado en el objeto `S`). Flujo para cambiar lógica: editar ese archivo → `build_index.py` → `test/run.py`.

## Comandos

```bash
# Deploy Node/Railway
npm install && npm start                 # server.js sirve sys-a-kyte/app/ en $PORT (3000 por defecto)

# Tests (necesita Python con playwright+openpyxl, y Node en sys-a-kyte/test)
cd sys-a-kyte/test && npm install
pip install playwright openpyxl && playwright install chromium
cd ../.. && python3 sys-a-kyte/tools/build_index.py
python3 sys-a-kyte/test/run.py            # ~1 min, debe terminar "RESUMEN: TODO OK"
SYS_XLSX=/ruta/otro.xlsx python3 sys-a-kyte/test/run.py   # probar con otro Excel

# Regenerar la tabla base de Kyte desde un Excel
python3 sys-a-kyte/tools/generar_tabla.py "mi_excel.xlsx"
```

## Estado actual (avances)

- **v1**: publicada (18 sep 2026) como Claude Artifact. Cruce exacto + por formato. Funcionando en producción.
- **v2**: especificada (decisiones de Kino del 19 sep 2026 — MEMORIA_PROYECTO.md §4-6) pero **sin construir todavía**: cruce por nombre cuando falta el Código, manejo de repetidos/copias/huérfanas, SKU nuevos clasificados. Criterios de aceptación ya definidos en §6.
- **Deploy Railway** (21 sep 2026): se agregó `server.js` + `package.json` (Express) para servir la app fuera de Claude Artifacts. **Pendiente:** no estaba documentado (ya se agregó nota en MEMORIA_PROYECTO.md §5), falta probarlo en Railway real, y decidir si `datos/` (825 KB de Excel) debe excluirse del bundle de producción.

## Errores y trampas conocidas (resumen — detalle en MEMORIA_PROYECTO.md §8)

- pandas convierte `NA`/`N/A`/`nan` en vacío al leer Excel → usar `openpyxl` o `keep_default_na=False` para analizar datos de Kyte/SYS.
- Cruce por formato solo es válido si el resultado es único (hay ~101 grupos de códigos parecidos que son productos distintos).
- Stock en blanco (Kyte) + 0 (SYS) = sin cambio; si no se maneja así, se inflan los "cambios" con falsos positivos.
- CSV: leer con `raw: true` para no perder ceros a la izquierda (ej. `0044`).
- Saltos de línea `\r\n` del Excel se normalizan a `\n` al leer la tabla.
- El Artifact solo carga scripts desde `cdnjs` y `cdn.jsdelivr.net/npm/`; el sandbox bloquea `fetch` y descargas directas (`tabla.js` se carga con `<script src>`, no `fetch`; las descargas usan `downloads.save`, una a la vez).
- No se declaró la capacidad `db` en el Artifact a propósito (lo haría exclusivo de la organización).
- **Sesiones de Claude Code en la nube:** el Chromium pre-instalado del sandbox puede no coincidir con la versión que trae el `pip install playwright` más reciente (error `Executable doesn't exist at .../chromium_headless_shell-XXXX`). No correr `playwright install` (no hay red para bajar el browser o se pisa el pre-instalado). Ver qué revisión hay en `/opt/pw-browsers/*.json` o en el Playwright de Node global (`/opt/node22/lib/node_modules/playwright/package.json`) e instalar esa misma versión con pip, ej. `pip install playwright==1.56.0`.

## Limpieza hecha el 21 sep 2026

Había en la raíz un archivo `SYS a Kyte — Memoria del proyecto` (sin extensión, texto plano sin formato) con el mismo contenido pero desactualizado respecto a `sys-a-kyte/MEMORIA_PROYECTO.md`. Se eliminó por ser un duplicado inferior — la memoria completa vive solo en `sys-a-kyte/MEMORIA_PROYECTO.md`. Sigue disponible en el historial de git (commit inicial) si hace falta.

## Registro de avances y errores

Agregar una línea arriba de todo (más reciente primero) cada vez que se resuelve algo importante, se descubre un error/bug, o se cierra un avance grande. Mantenerlo corto — el detalle largo va en `MEMORIA_PROYECTO.md`.

| Fecha | Tipo | Nota |
|---|---|---|
| 2026-09-21 | avance | Reporte de cambios → hoja "Nuevos en SYS": 3 columnas nuevas (Marca, Accesorio, Familia ya en tu tabla) clasificando cada SKU nuevo con stock. Era parte de la v2 (§6.4), no construida hasta ahora; los 4 archivos de Kyte siguen sin tocarse (Kino eligió solo el reporte, no pantalla ni archivos). |
| 2026-09-21 | avance | Tarjetas por hoja (`#chain`) ahora muestran cuánto de "cambian de stock" y "de precio" es suba vs. baja (`↑N` verde / `↓N` rojo), reusando `.delta up/down` ya existente — sin colores nuevos. Cambio solo visual: `procesar()` suma `st.subeS/bajaS/subeP/bajaP`, no toca los archivos que se descargan (test/run.py sigue en TODO OK, celda por celda). |
| 2026-09-21 | decisión | Flujo de git cambiado: se pushea directo a `main`, sin ramas por feature (Kino verifica en vivo). Ver sección "Flujo de git" arriba. |
| 2026-09-21 | avance | Creado `CLAUDE.md` como memoria de arranque para Claude Code, para no re-explorar el proyecto en cada sesión. |
| 2026-09-21 | avance | Documentado en MEMORIA_PROYECTO.md el deploy Node/Express + Railway (`server.js`, `package.json`); seguía sin aparecer ahí. Pendiente probar el deploy real. |
| 2026-09-21 | limpieza | Eliminado duplicado `SYS a Kyte — Memoria del proyecto` en la raíz (texto plano, contenido desactualizado). |
| 2026-09-19 | avance | v2 especificada: decisiones de Kino sobre cruce por nombre, repetidos, copias y SKU nuevos (MEMORIA_PROYECTO.md §3-6). Todavía sin construir. |
| 2026-09-18 | avance | v1 publicada como Claude Artifact. |
