# Embalaje predeterminado por SKU — Diseño

**Fecha:** 2026-08-10
**Autor:** Leo (con asistencia de Claude)

## Resumen

La operación pasa a usar embalajes con tamaños predeterminados (aproximadamente 5 cajas y 5 bolsas).
Cada SKU —individual o combo— tiene asignado uno de esos embalajes, y el operario necesita verlo
en la etiqueta al momento de empaquetar.

El alcance es **informativo para el operario**. Explícitamente **fuera de alcance**:

- Las dimensiones y el peso que se suben a MercadoLibre no cambian. Se siguen cargando a mano en el
  Excel de medidas y se siguen subiendo las columnas +20% como hoy.
- El embalaje **no** reemplaza ni valida las medidas cargadas.
- No hay asignación automática de embalaje por medidas del producto.

## Contexto del código

- **Excel madre de medidas** (`MedidasExcelManager`, `MedidaSku`): hoja única con columnas
  `SKU | PRODUCTO | Ancho | Alto | Profundidad | Peso | Ancho +20% | Alto +20% | Profundidad +20% | Peso +20% | SUBIDO | ERROR`
  (índices 0..11). La app agrega los SKU pendientes de medición, el usuario carga las medidas a mano
  y el botón "Subir medidas a ML" envía los `SELLER_PACKAGE_*`.
  Está detrás del checkbox `medidasEnabledCheck`; si está apagado, el módulo entero no interviene.
- **Inyección ZPL** (`MainController.injectZplHeaders`): por cada etiqueta inyecta, en un bloque
  `^LH0,0` al inicio, el número de posición `#N` en `^FO45,30` con fuente 35 y triple pasada;
  opcionalmente el banner `MEDIR: <sku>` en video inverso ocupando `^FO200,15^GB580,65,65`.
  Después inyecta `ZONA:` (anclada al texto "Unidad") y `COD.EXT.:` (anclada al texto "SKU:").
  Las anclas de texto son frágiles: dependen del formato de etiqueta de ML y ya está previsto
  loguear una advertencia si cambian.
- **Aviso post-proceso** (`MainController.mostrarMensajeSkusFaltantes`): diálogo con los SKU sin
  medidas detectados en el lote. Se invoca desde los dos caminos (`Archivo Local` y `API ML`).

## Diseño

### 1. Excel: hoja de SKUs

Se agrega la columna **`EMBALAJE`** al final, en el índice 12 (después de `ERROR`).

Va al final —y no junto a las columnas de dimensiones— para no correr los índices existentes ni
romper las fórmulas que el usuario ya tiene cargadas (`BUSCARX` en PRODUCTO, `base*1.2` en las +20%).

- `HEADERS` suma `"EMBALAJE"`; nueva constante `COL_EMBALAJE = 12`.
- `asegurarColumnaEmbalaje(workbook, sheet)`, análogo al `asegurarColumnaError` existente: si el
  header no está, lo escribe. Se llama desde `agregarPendientes` y desde `marcarResultados`.
- En `agregarPendientes`, la celda `EMBALAJE` de un SKU nuevo se deja vacía con el estilo amarillo
  tenue de "falta cargar" (`crearEstiloCeldaFaltante`), igual que las celdas de medidas.
- Se aplica **validación de datos** sobre la columna `EMBALAJE`, desde la fila 2 hasta la última
  fila con SKU: lista desplegable apuntando a la columna de códigos de la hoja `EMBALAJES`
  (`DataValidationHelper` de POI, con `createFormulaListConstraint("EMBALAJES!$A$2:$A$100")`).
  El `$A$100` es un tope holgado fijo: si el catálogo creciera más allá de 99 filas, las de más
  no aparecerían en el desplegable, pero sí serían válidas al leerlas.
  La validación se declara como advertencia, no como bloqueo, para no impedir editar el archivo
  a mano si hiciera falta.
- La lectura en `leerMedidas` reconoce el header `EMBALAJE` por nombre normalizado, como el resto.

### 2. Excel: hoja `EMBALAJES` (catálogo)

Hoja nueva creada **solo con encabezados**, sin filas de ejemplo. El usuario carga las que necesite.

| CÓDIGO | TIPO | Ancho cm | Alto cm | Profundidad cm |
|--------|------|----------|---------|----------------|

- Se crea si no existe, al mismo tiempo que se asegura la columna `EMBALAJE`. Nunca se pisan filas
  existentes.
- `CÓDIGO` es el texto que se imprime en la etiqueta (ej. `CAJA 3`, `BOLSA CHICA`).
- `TIPO` y las medidas son documentación del catálogo: **la app no las usa**. Existen para que el
  catálogo esté descrito en un solo lugar.
- No hay límite de 5+5; se lee lo que haya.

### 3. Modelo y lectura

**`ar.com.leo.etiquetas.model.Embalaje`** (record nuevo):

```java
public record Embalaje(String codigo, String tipo, Double anchoCm, Double altoCm, Double profundidadCm) {}
```

**`MedidaSku`**: suma el campo `String embalaje` — el código crudo leído de la celda, `""` si está
vacía. No se agrega ningún método derivado: la clasificación la hace `EmbalajeResolver` (§4), que
necesita distinguir "vacío" de "código que no existe", cosa que un `tieneEmbalaje()` no captura.

`EmbalajeResolver` quedó como clase de utilidad con métodos estáticos (`indexar`, `resolver`,
`normalizar`) en vez de una instancia: no tiene estado propio.

**`MedidasExcelManager`**: nuevo método público

```java
public Map<String, Embalaje> leerCatalogoEmbalajes(Path excelPath) throws Exception
```

- Devuelve mapa `código normalizado → Embalaje`, preservando el código tal como fue escrito dentro
  del record (es lo que se imprime).
- Si el archivo o la hoja no existen, devuelve mapa vacío sin lanzar excepción.
- Toma el `fileLock` existente, como el resto de las operaciones sobre el archivo.

**Normalización de códigos**: para comparar SKU↔catálogo se usa `trim` + mayúsculas + colapso de
espacios internos (la misma `normalizarHeader` ya existente sirve). Así `caja 3`, `CAJA  3` y
`Caja 3` resuelven al mismo embalaje. En la etiqueta se imprime el código tal como figura en el
**catálogo**, no como lo escribió el usuario en la fila del SKU, para que la impresión sea uniforme.

### 4. Resolución SKU → texto de embalaje

Clase nueva **`ar.com.leo.etiquetas.parser.EmbalajeResolver`** (lógica pura, testeable, fuera de
`MainController`):

```java
public record ResultadoEmbalaje(String textoEtiqueta, Estado estado) {
    public enum Estado { OK, SIN_ASIGNAR, CODIGO_INVALIDO }
}
public ResultadoEmbalaje resolver(String codigoCrudo, Map<String, Embalaje> catalogo)
```

| Caso | `textoEtiqueta` | `estado` |
|------|-----------------|----------|
| Código presente y en catálogo | código del catálogo | `OK` |
| Celda vacía / SKU ausente del Excel | `-` | `SIN_ASIGNAR` |
| Código presente pero no está en el catálogo | `-` | `CODIGO_INVALIDO` |

Un código inválido se imprime como `-` (nunca el texto errado) y se reporta aparte, para que un
typo no pase inadvertido.

### 5. Inyección en la etiqueta ZPL

Se agrega la línea `EMBALAJE: <texto>` **debajo del `#N`**, en el mismo bloque `^LH0,0` que ya se
inyecta al inicio de cada etiqueta, con coordenadas absolutas:

```
^FO45,85^A0N,30,30^FDEMBALAJE: CAJA 3^FS
^FO46,85^A0N,30,30^FDEMBALAJE: CAJA 3^FS
```

Doble pasada con 1px de offset para simular negrita, igual que `ZONA` y `COD.EXT.`.

**Por qué ahí.** El espacio de picking va de y≈120 a y≈394 (línea de corte + ícono de tijera) y está
lleno: `COD.EXT.` queda en y≈347, a 47px de la línea de corte. En cambio la franja y 85–125 del
margen superior izquierdo está libre: el `#N` termina en y=65, el banner `MEDIR` termina en y=80
(y arranca en x=200) y el `Pack ID` empieza en y=130. Con fuente 30 la línea ocupa y 85–115, sin
tocar nada. Además, al ir en el bloque de coordenadas absolutas ya inyectado, **no depende de las
anclas de texto de ML** ("Unidad", "SKU:"), a diferencia de `ZONA` y `COD.EXT.`.

**Alcance:** solo zonas distintas de `CARROS`. Las etiquetas de carro listan varios productos y
quedan sin cambios.

**Siempre presente:** la línea se imprime siempre (con `-` si falta el dato), para que la ausencia
sea visible y no se confunda con un problema de impresión.

**Si el módulo de medidas está desactivado** (checkbox apagado, ruta vacía o archivo inexistente):
no se inyecta la línea y no se avisa nada. La etiqueta sale exactamente como hoy.

**Dónde se carga el catálogo:** en `MainController.loadCatalogoEmbalajes()`, hermano de
`loadMedidasMap()` y con el mismo criterio de activación (checkbox encendido, ruta válida; si no,
devuelve `null` y la línea no se inyecta). Ambos se pasan a `injectZplHeaders`.

Esto abre el Excel dos veces por lote en vez de una. Se eligió así por simplicidad: unificar las
dos lecturas obligaba a refactorizar `leerMedidasInterno` para compartir el `Workbook`, y el costo
de la segunda apertura de un archivo local chico es despreciable frente a la descarga de etiquetas.

### 6. Aviso en la app

`injectZplHeaders` acumula, además de los SKU pendientes de medición, un mapa
`sku → Estado` con los `SIN_ASIGNAR` y `CODIGO_INVALIDO` (solo para las etiquetas donde la línea
aplica, es decir zonas distintas de CARROS).

`mostrarMensajeSkusFaltantes` suma una sección al diálogo existente:

```
3 SKU(s) sin embalaje asignado en este lote:
  1241212, 1241255, 998877

1 SKU(s) con un código de embalaje que no existe en el catálogo:
  1240001 → "CAJA3"
```

Se muestra en el mismo diálogo (no uno nuevo) porque es la misma clase de aviso post-proceso y el
usuario ya lo tiene incorporado al flujo. Si no hay faltantes de medición pero sí de embalaje, el
diálogo se muestra igual.

## Testing

**`MedidasExcelManagerTest`** (nuevo, con workbooks armados en memoria y volcados a un `@TempDir`):

- `leerCatalogoEmbalajes` con hoja ausente → mapa vacío, sin excepción.
- `leerCatalogoEmbalajes` con hoja poblada → códigos y medidas leídos; filas con código vacío se
  ignoran.
- `leerMedidas` sobre un Excel **sin** la columna `EMBALAJE` (archivo viejo) → `embalaje` vacío,
  el resto de los campos intactos. Es el caso de compatibilidad hacia atrás.
- `agregarPendientes` sobre un Excel viejo → agrega el header `EMBALAJE` sin alterar las columnas
  ni las fórmulas existentes.

**`EmbalajeResolverTest`** (nuevo): los tres estados, más normalización (`caja 3` ≡ `CAJA 3`) y la
regla de imprimir el código del catálogo y no el escrito por el usuario.

La inyección ZPL en sí no se testea: vive en un método privado de `MainController` (~1900 líneas) y
extraerla completa excede este cambio. Lo que se extrae y sí queda cubierto es la resolución
`SKU → texto`, que es donde está la lógica. La construcción del string ZPL es una concatenación de
coordenadas fijas, verificable a ojo en el archivo `Etiquetas/etiquetas_ordenadas_*.txt` que la app
guarda en cada corrida.

## Archivos afectados

| Archivo | Cambio |
|---------|--------|
| `etiquetas/model/Embalaje.java` | **Nuevo** — record del catálogo |
| `etiquetas/model/MedidaSku.java` | Campo `embalaje` + `tieneEmbalaje()` |
| `etiquetas/parser/EmbalajeResolver.java` | **Nuevo** — resolución SKU → texto/estado |
| `etiquetas/parser/MedidasExcelManager.java` | Columna `EMBALAJE`, hoja `EMBALAJES`, validación de datos, `leerCatalogoEmbalajes` |
| `ui/MainController.java` | Lee el catálogo, inyecta la línea ZPL, acumula y reporta los faltantes |
| `test/.../MedidasExcelManagerTest.java` | **Nuevo** |
| `test/.../EmbalajeResolverTest.java` | **Nuevo** |

## Compatibilidad

Un Excel de medidas existente sigue funcionando sin tocarlo: se lee sin la columna `EMBALAJE`
(todos los SKU quedan `SIN_ASIGNAR`) y la columna junto con la hoja `EMBALAJES` se crean la próxima
vez que la app escriba el archivo. Ninguna medida ni fórmula cargada se modifica.
