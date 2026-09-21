# SYS a Kyte — Memoria del proyecto

> Léelo primero. Aquí está qué hace la app, las reglas que no se rompen, las decisiones de Kino y qué falta construir.
> Estado: **v1 publicada (18 sep 2026)** · **v2 especificada y sin construir (decisiones del 19 sep 2026)**.

- **App (Artifact, privado):** https://claude.ai/artifact/KyChVS1ZYewEqiAPecFE15
- **Claves y accesos:** ninguno. No hay contraseñas, API keys ni servidor. Todo se procesa en el navegador.
- **Idioma de trabajo:** español. Kino escribe informal y rápido; prefiere respuestas cortas, directas y prácticas.

---

## 0. Clave del proyecto (las reglas que no se rompen)

1. **Qué es:** subes el archivo del SYS, la app cruza por SKU y bajas los archivos de Kyte (KYTE CON SYS INV, CURREN, JOEFOX, POEDAGAR) con stock y precio al día, más un reporte. Reemplaza el Excel con XLOOKUP.
2. **El SYS manda:** es la fuente de verdad del stock y del precio base.
3. **Stock actual** de Kyte = existencia del SYS. Negativo pasa a 0. Si en Kyte estaba en blanco y el SYS dice 0, se deja en blanco.
4. **Precio\*** = 2 × precio del SYS (casilla para apagarlo). Pisa los precios propios de Kyte: el 18 sep eran 100 productos.
5. **Lo que no cruza no se toca.** La app nunca crea, borra ni oculta productos por su cuenta.
6. **Cruce flexible solo sin duda:** por formato o por nombre únicamente cuando hay una sola coincidencia posible.
7. **Formato de salida:** las 14 columnas de Kyte, sin la columna auxiliar "SYS".
8. **SKU nuevos: nada automático.** Se avisan y se listan; Kino decide qué hacer.
9. **Se ve antes de bajar:** tabla "antes → después" y alerta si cruza menos del 50 % de los productos con código.
10. **No cambiar estas reglas sin preguntarle a Kino.** Antes de tocar código, correr `test/run.py`; debe terminar en `TODO OK`.

---

## 1. Uso diario (para el equipo)

1. Abrir la app y mirar arriba a la derecha: **Tabla base: 18 sep 2026** (la fecha de la tabla guardada).
2. **Subir el SYS:** arrastrar el Excel a la casilla punteada "SYS INV" o hacer clic. Sirve el Excel del SYS (hoja con `Artículo`, `Descripción`, `Exist. unidades`, `Precio`), el libro completo con una hoja "SYS INV", o un CSV (coma o punto y coma).
3. **Revisar:**
   - Cada hoja muestra cuántos productos **cambian de stock** y de precio.
   - Aviso amarillo o rojo = el archivo probablemente no es el del SYS. No subir nada a Kyte.
   - En "Cambios: antes → después" mirar sobre todo los **precios** (se pisan los que tenían precio propio).
   - La casilla **"Actualizar el precio (2×)"** se apaga si un día se quiere solo stock.
4. **Bajar:** cada hoja por separado (`Bajar .xlsx`), **"Bajar los 4 (.zip)"** o **"Reporte de cambios"**. Cada descarga pide confirmación en pantalla; con el zip es una sola.
5. **Subir a Kyte** los archivos descargados.
6. Volver a la app y pulsar **"Guardar como tabla base"** (solo quien edita el Artifact). Así la próxima vez los cambios se comparan contra lo que ya está en Kyte.
7. **Cuando cambie el catálogo en Kyte** (productos nuevos, nombres, categorías): armar un Excel con las hojas en formato Kyte y usar **"Cambiar tabla"** → "Guardar en la app".

Quién puede qué: cualquiera con acceso al Artifact sube el SYS y baja archivos. **Guardar la tabla** requiere permiso de edición; para el resto de personas del equipo hay que compartir el Artifact desde su menú *Share*.

Pendiente de confirmar en la práctica: no se ha probado si Kyte acepta la importación del archivo tal cual. Probar primero con una hoja pequeña.

---

## 2. Reglas de la v1 (publicada)

**Cruce (Código de Kyte ↔ Artículo del SYS)**

| Nivel | Regla |
|---|---|
| Exacto | Iguales, sin distinguir mayúsculas y sin espacios en los bordes (igual que el XLOOKUP del Excel). |
| Formato | Si no hay exacto: quitar espacios y `- _ : . / \` de ambos lados. Vale solo si **un único** artículo del SYS coincide. Ej.: `DR-256-1` (Kyte) = `DR256-1` (SYS). |
| Sin cruce | No se toca. Va al reporte con el motivo: "Sin código" o "El código no está en el SYS". |

**Qué se escribe en las filas que cruzan:** `Stock actual` y, si la casilla está activa y el SYS trae precio mayor que 0, `Precio*` = `round(precio SYS × 2)`. Nada más cambia (ID, nombre, categoría, costo, promoción, "Mostrar en el catálogo", etc. quedan igual).

**Hojas:** `KYTE CON SYS INV` (maestra) y `CURREN`, `JOEFOX`, `POEDAGAR`. Las tres últimas son subconjuntos de la maestra (mismo ID, celdas idénticas). Cada hoja se procesa por separado.

**Archivos que baja:** `<HOJA>_<AAAA-MM-DD>.xlsx`, `Kyte_actualizado_<fecha>.zip` (las 4 hojas) y `Reporte_SYS_<fecha>.xlsx` (hojas `Resumen`, `Cambios`, `Sin cruzar`, `Nuevos en SYS` —solo con stock mayor que 0—, `Por formato`).

**Columnas Kyte (en este orden):** ID · Nombre\* · Categoría · Código · Descripción · Costo unitario · Precio\* · Precio de promoción · Mostrar en el catálogo · Destacar · Fracción? · Controlar stock · Stock actual · Stock mínimo.

---

## 3. Datos clave (inventario del 18 sep 2026)

**Lo que hace la v1 hoy**

| Dato | Valor |
|---|---|
| Artículos en el SYS | 9.993 (3.140 con stock, 245 en negativo, 0 repetidos exactos) |
| Productos en KYTE CON SYS INV | 5.792 |
| Cruzan con el SYS | 1.953 (1.943 exactos + 10 por formato) |
| Sin cruzar | 3.839 (3.410 sin Código, 429 con Código que el SYS no trae) |
| Cambian de stock | 803 KYTE · 87 CURREN · 143 JOEFOX · 256 POEDAGAR |
| Cambian de precio | 100 KYTE · 20 CURREN · 10 JOEFOX · 5 POEDAGAR |
| Artículos del SYS que no están en la tabla | 8.060 (1.535 con stock), **cifra inflada**: ver abajo |

**Hallazgo grande: el SKU está en el nombre.** Muchos productos de Kyte tienen el `Código` vacío pero el SKU al inicio del `Nombre*` (ej. `1401-4 POEDAGAR CRONOGRAFO AUTOMATICO METAL HOMBRE`). Hoy salen `#N/A` y nunca se actualizan. Cruzándolos por nombre, la cobertura sube de 1.953 a 4.190 filas y los "nuevos" bajan de 1.535 a 114 con stock.

**Lo que dará la v2 con las reglas elegidas** (KYTE CON SYS INV, precio ×2 activo)

| Dato | Valor |
|---|---|
| Sin Código, SKU en el nombre y existe en el SYS | 2.472 |
| — nivel alto (token con dígito y precio Kyte = 2 × SYS) | 2.237 → se actualizan |
| — nivel medio (precio distinto o token sin dígitos) | 235 → "Confirmar", no se tocan |
| Filas que reciben actualización | **4.190** de 5.792 (1.943 + 10 + 2.237) |
| Filas con stock distinto al de Kyte | 1.612 |
| Sin cruce (huérfanas) | 1.367 (429 con Código) |
| Huérfanas visibles en catálogo con stock en Kyte | 68 (3.380 uds) |
| SKU repetidos entre las filas que se actualizan | 24 SKU · 58 filas · 12 con stock · **2.162 uds duplicadas** |
| Copias | 28 con "(copia)" en el nombre · 17 nombres idénticos (34 filas) |
| SKU nuevos del SYS (que ninguna fila usa) | 5.627: **114 con stock (45.929 uds)**, 5.442 con stock 0, 71 negativos |

- Mayores nuevos: `TP-04` TAPA (33.196 uds), `1203-TP` (2.570), `FRASE` adhesiva (2.081), `RY-45` estuche (2.000). Casi todo son accesorios y piezas, no relojes.
- Repetidos con stock: `S-640+S-647` (815 uds, 2 filas), `PUL-5` (422), `DR256-1` (253), `V0028` (221), `K-673E` (120).
- **Grupo 1203:** 20 filas de Kyte llamadas `1203 <personaje> INFANTIL MUSICAL/LUZ/TAPA` (19 visibles) comparten el SKU genérico `1203` (SYS: 8.152 uds, $7.500). Su precio en Kyte es $13.000 y no $15.000, así que caen en "Confirmar" y **no reciben stock**. Si algún día se confirman, se duplicarían 19 × 8.152 unidades.
- **Trampa del SYS:** 101 grupos de códigos parecidos que son **productos distintos** (`1179` SKMEI LED binario vs `11-79` SENKO; `2385-1` JOEFOX vs `238-5-1` DISITE). Por eso el cruce flexible exige coincidencia única.

> Nota sobre cifras: en el chat del 19 sep aparecieron números un poco distintos (2.478 / 2.242 / 234, "158.000 unidades fantasma", 117 nuevos, 73 huérfanas). Venían de un primer cálculo con pandas, que convierte textos como `NA` en vacío, y que además contaba las filas de "Confirmar". **Valen las cifras de esta sección.** Lo importante: con las reglas elegidas la exposición por repetidos es de unas 2.162 unidades, no de 158.000.

---

## 4. Decisiones de Kino para la v2 (19 sep 2026)

| Tema | Decisión |
|---|---|
| Productos sin Código con el SKU en el nombre | **Cruzar sin tocar el Código.** Nivel alto: se actualiza stock y precio. Nivel medio: a "Confirmar" sin tocar. La columna `Código` del archivo queda como está (seguirán cruzando por nombre cada vez). |
| SKU repetido en varias filas de Kyte | **Todas igual, marcadas.** Todas las filas del grupo reciben el mismo stock del SYS y quedan marcadas como repetidas. Riesgo asumido: sobreventa por unidades duplicadas. |
| SKU nuevos con stock | **Solo lista y aviso.** Destacados en pantalla y en el reporte, sin archivo aparte. Nada se crea solo. |
| Copias (`(copia)`) y nombres idénticos | **Solo marcar y listar** en el reporte. No se cambia nada en los archivos. |

**Incluido en la propuesta que Kino vio y no objetó** (confirmar al construir): niveles de confianza como columna en el reporte, lista "Confirmar" para el nivel medio, y alerta aparte para las huérfanas visibles con stock (68 el 18 sep).

**Sugerencia de implementación (no decidida):** mostrar en pantalla la cifra "unidades duplicadas" de los repetidos, para que la decisión de "todas igual" se tome viendo el número.

---

## 5. Arquitectura y archivos

**Un solo HTML + un archivo de datos. Sin servidor ni base de datos.**

```
sys-a-kyte/
├── MEMORIA_PROYECTO.md        este archivo
├── app/
│   ├── sys-a-kyte.html        FUENTE. Fragmento que se publica como Artifact (sin doctype/head/body)
│   ├── index.html             versión autónoma (fragmento + esqueleto); se genera con tools/build_index.py
│   └── tabla.js               tabla base: window.__TABLA__ = {v, guardada, origen, cols, hojas:[{nombre, rows}]}
├── tools/
│   ├── build_index.py         sys-a-kyte.html → index.html
│   └── generar_tabla.py       Excel con hojas Kyte → app/tabla.js
├── test/
│   ├── run.py                 verificación en navegador (Playwright) contra un cálculo independiente
│   ├── shots.py               capturas (escritorio, móvil, modo oscuro)
│   └── package.json           xlsx 0.18.5 y jszip 3.10.1 para probar sin internet
└── datos/
    └── INVENTARIO_KINO_18_de_sep.xlsx   Excel original (SYS INV + las 4 hojas Kyte)
```

**Dentro de `sys-a-kyte.html`**
- Constantes: `COLS` (14 columnas), `I` (índices), `REQ` (columnas obligatorias de una hoja Kyte).
- Lectura: `leerSys(wb)` detecta la hoja y las columnas del SYS; `leerTabla(wb)` toma cada hoja con formato Kyte y normaliza saltos de línea.
- Motor: `procesar(tabla, sys, opts)` devuelve `{hojas, cambios, sin, formato, nuevos, tot, pct, …}`.
- Salida: `libroHoja(h)`, `libroReporte()`, `entregar()` y `conBloqueo()` (una descarga a la vez).
- Flujo: `cargarSys`, `prepararTabla`, `proponerBase`, `guardarPend`, `correr`. Pintado: `render*`.
- Estado en un solo objeto `S` (`tabla`, `sys`, `res`, `opts`, `filtro`, `pend`, …).

**Librerías** (por `cdn.jsdelivr.net/npm`): SheetJS `xlsx@0.18.5` y `jszip@3.10.1`.

**Capacidades del Artifact declaradas:** `artifact` (guardar la tabla) y `downloads` (bajar archivos).
- El sandbox del Artifact bloquea `<a download>`: por eso se usa `downloads.save`, que pide confirmación y admite una a la vez.
- Guardar la tabla = `artifact.publish({'tabla.js': …})` (publicación por archivos, solo quien edita). Si falla, la tabla se usa solo en esa sesión y la app lo avisa.
- No se declaró `db` a propósito: haría el Artifact exclusivo de la organización.

**Fuera de Claude:** abrir `app/index.html` con `tabla.js` al lado (necesita internet para las librerías). Las descargas usan `<a download>`; "Guardar en la app" no guarda, así que para cambiar la tabla se regenera `tabla.js` con `tools/generar_tabla.py`.

---

## 6. Especificación de la v2 (lo que falta construir)

Todo esto va en `procesar()` y en el pintado; los archivos Kyte solo cambian por el nuevo nivel de cruce.

1. **Nivel "nombre".** Si `Código` está vacío, tomar el primer token del `Nombre*` en mayúsculas. Si existe como Artículo del SYS:
   - **Alta:** el token tiene al menos un dígito **y** `Precio*` = `round(precio SYS × multiplicador)`. Se trata como un cruce normal (stock y precio). **No escribir el Código.**
   - **Media:** cualquier otro caso. No se toca. Va a la lista **Confirmar** con: ID, token, nombre, precio Kyte, precio SYS×2, stock Kyte, stock SYS.
   - Orden de precedencia: Código exacto → Código por formato → nombre. Un artículo del SYS usado por cualquier nivel (incluido "media") **no** es nuevo.
2. **Repetidos.** Agrupar por artículo del SYS las filas que reciben actualización (exacto, formato, nombre-alta). Grupos de 2 o más: todas reciben el stock (decisión), se marcan "Repetido ×N", y se cuenta `unidades duplicadas = (filas − 1) × max(0, stock SYS)`. Aviso aparte si un SKU en "Confirmar" tiene 2 o más filas (caso 1203).
3. **Copias.** `(copia)` en el nombre (sin distinguir mayúsculas) y nombres idénticos (con trim, sin distinguir mayúsculas) con 2 o más filas → hoja `Copias` del reporte, agrupadas. No modifican archivos.
4. **SKU nuevos.** Los del SYS que ningún nivel usó. Con stock mayor que 0: tarjeta destacada "SKU nuevos con stock: N · U uds", lista ordenada por unidades y hoja del reporte. Sugerencia de clasificación: marca por palabra clave (POEDAGAR, JOEFOX, CURREN, SANSE, KARDOO, SKMEI, T5, JAKCOM, LUGANOS, CASIO, DISITE, EXTRI), accesorio (ESTUCHE, CAJA, BOLSA, CORREA, PILA, TAPA, EXHIBIDOR), y "familia ya existe en tu tabla" (código sin el sufijo numérico final: `1401-3` → `1401`). Stock 0 y negativos: solo conteo. **Sin archivo aparte y sin agregarlos a ningún archivo.**
5. **Huérfanas.** Filas sin cruce (y que no estén en "Confirmar") con `Mostrar en el catálogo` = S y `Stock actual` mayor que 0 → alerta "N visibles con stock que el SYS no trae (U uds)" y hoja `Huérfanas`. No se tocan.
6. **Reporte.** `Resumen` con una columna por nivel (Exacto, Formato, Nombre alta, Confirmar, Sin cruce) y hojas nuevas `Confirmar`, `Repetidos`, `Copias`, `Huérfanas`; `Nuevos en SYS` con marca y familia.
7. **Pantalla.** Nuevas tarjetas en el resumen (Repetidos, Confirmar, Nuevos con stock, Copias, Huérfanas) con color de advertencia, y un filtro "Repetidos" en la tabla de cambios.

**Criterios de aceptación** (Excel del 18 sep, hoja KYTE CON SYS INV, precio ×2 activo). Ampliar `load_expected()` de `test/run.py` con este nivel y comprobar:

| Comprobación | Esperado |
|---|---|
| Filas que reciben actualización | 4.190 (1.943 exactas + 10 formato + 2.237 nombre-alta) |
| Confirmar | 235 |
| Sin cruce | 1.367 (429 con Código) |
| Filas con stock distinto | 1.612 |
| Repetidos | 24 SKU · 58 filas · 12 con stock · 2.162 uds duplicadas |
| Copias | 28 con "(copia)" · 17 nombres / 34 filas |
| Huérfanas visibles con stock | 68 · 3.380 uds |
| SKU nuevos con stock | 114 · 45.929 uds |
| Archivos descargados | celda por celda igual al cálculo aparte; la columna `Código` no cambia |

---

## 7. Publicar, probar y cambiar la tabla

**Probar** (necesita Python con `playwright` y `openpyxl`, y Node):
```
cd test && npm install
pip install playwright openpyxl && playwright install chromium
cd .. && python3 tools/build_index.py
python3 test/run.py          # ~1 min; debe terminar en "RESUMEN: TODO OK"
```
`run.py` carga el SYS real en un navegador, baja los 4 archivos, el zip y el reporte, y compara **celda por celda** contra un cálculo independiente en Python. También prueba: precio apagado, "Cambiar tabla", CSV, archivo equivocado, alerta de pocos cruces y móvil. Los conteos iniciales están escritos para el Excel del 18 sep; con otros datos hay que actualizarlos. Para usar otro Excel: `SYS_XLSX=/ruta/archivo.xlsx python3 test/run.py`.

**Cambiar la tabla:** en la app, "Cambiar tabla" → elegir el Excel → "Guardar en la app". O desde la línea de comandos: `python3 tools/generar_tabla.py "mi_excel.xlsx"` y republicar.

**Republicar con Claude:** publicar `app/sys-a-kyte.html` con el archivo de datos `tabla.js` (`files: {"tabla.js": "app/tabla.js"}`), capacidades `{artifact: {}, downloads: true}`, y pasando la `url` del Artifact para conservar el enlace.

---

## 8. Trampas y aprendizajes

- **pandas convierte `NA`, `N/A`, `nan` en vacío** en columnas de texto. Para analizar datos de Kyte o del SYS usar `openpyxl` (o `keep_default_na=False`). Esto causó diferencias de unas 6 filas en cifras del chat.
- **Cruce por formato solo si es único.** El SYS tiene 101 grupos de códigos parecidos que son productos distintos.
- **Stock en blanco en Kyte y 0 en el SYS = sin cambio.** Si no, se inflan los "cambios" con miles de falsos positivos.
- **CSV:** leerlo con `raw: true` para no perder ceros a la izquierda (`0044`). Al escribir xlsx usar `compression: true` (1 MB contra 2,9 MB).
- **Saltos de línea:** el Excel trae `\r\n` en las descripciones; SheetJS los conserva y openpyxl los normaliza. La app los normaliza a `\n` al leer la tabla.
- **XLOOKUP de Excel no distingue mayúsculas;** el motor tampoco.
- **Artifact:** solo se cargan scripts de cdnjs y `cdn.jsdelivr.net/npm/`; el sandbox bloquea `fetch` y descargas directas. La tabla se lee con `<script src="tabla.js">`, no con `fetch`.
- **Las 3 hojas de marca** son idénticas a la maestra para los mismos ID. Si un día difieren, cada hoja se procesa por separado, así que no se pierde nada.

---

## 9. Prompt para continuar (pegar junto con este archivo y el zip)

```
Lee MEMORIA_PROYECTO.md y app/sys-a-kyte.html. Implementa la v2 de la sección 6
sin romper la v1. Corre test/run.py antes y después de cada cambio y amplía
load_expected() con el nivel "nombre" para cumplir los criterios de aceptación.
No cambies las reglas de la sección 0 ni las decisiones de la sección 4 sin
preguntarme. Al terminar, republica el Artifact conservando la URL y dime en
una frase qué cambió.
```
