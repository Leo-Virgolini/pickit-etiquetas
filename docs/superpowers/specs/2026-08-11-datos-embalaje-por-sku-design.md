# Datos de embalaje por SKU en la etiqueta — Diseño

**Fecha:** 2026-08-11
**Autor:** Leo (con asistencia de Claude)

Reemplaza a [2026-08-10-embalaje-por-sku-design.md](2026-08-10-embalaje-por-sku-design.md), que
modelaba el embalaje como un único código elegido de un catálogo.

## Resumen

El embalaje deja de ser un código de catálogo y pasa a ser un conjunto de datos que el usuario carga
a mano en el Excel de medidas: qué caja o bolsa va, si lleva pluribol, si lleva rollo inflado y
cuántos paños, más observaciones libres. La etiqueta muestra solo los datos que ese SKU tenga.

**La app solo lee.** No crea columnas, no inserta, no valida, no genera hojas. Eso elimina la hoja
catálogo, el desplegable, la validación de datos y toda la migración de estructura introducidos por
el diseño anterior — que es de donde salieron los defectos de las rondas de review.

Fuera de alcance: las dimensiones y el peso que se suben a MercadoLibre no cambian.

## Columnas

Van después de `SUBIDO` y **las crea y carga el usuario**:

| Columna | Contenido |
|---|---|
| `N° Bolsa` | Número de bolsa |
| `Nombre Caja` | Nombre de la caja (ej. `GRANDE`) |
| `N° Caja` | Número de caja |
| `PLURIBOL` | `SI` / vacío |
| `CANT PLURIBOL` | Vueltas de pluribol |
| `ROLLO INFLABLE` | Tipo de rollo (ej. `DIAMANTE`, `CUADRADO`) |
| `CANT PAÑOS` | Cantidad de paños |
| `OBSERVACIONES` | Texto libre |

Se ubican **por su encabezado**, no por índice: el usuario puede reordenarlas o intercalar columnas
propias. Si una falta, ese dato no se muestra y el resto sigue funcionando.

Esa resolución es **compartida por la lectura y la escritura** (`resolverColumnas`, que devuelve un
record `Columnas`). `agregarPendientes` no puede escribir por índice fijo: si el usuario intercaló
sus columnas entre las de la app, limpiar el rango 2..9 le borraría datos y el `NO` de `SUBIDO`
caería en una columna de medidas.

Los patrones de medida (`ANCHO`, `ALTO`, `PROFUN`, `PESO`) se evalúan **antes** que los de embalaje:
un encabezado como `Ancho caja cm` es natural —lo que se mide es la caja— y con los de embalaje
primero terminaría asignado a la columna de caja, dejando la dimensión sin leer.

**La hoja de medidas se ubica por su columna `SKU`**, no por posición: el usuario puede tener hojas
propias antes de ella.

Reconocimiento sobre el header normalizado (mayúsculas, espacios colapsados), en este orden —
el orden importa porque los patrones se solapan:

1. contiene `CANT` y `PLURIBOL` → cantidad de pluribol
2. contiene `PLURIBOL` → pluribol
3. contiene `NOMBRE` y `CAJA` → nombre de caja
4. contiene `CAJA` → número de caja
5. contiene `BOLSA` → número de bolsa
6. contiene `ROLLO` → tipo de rollo
7. contiene `PAÑO` o `PANO` → cantidad de paños
8. contiene `OBSERV` → observaciones

`PAÑO`/`PANO` se aceptan ambos porque el encabezado puede escribirse sin la eñe.

## Modelo

**`ar.com.leo.etiquetas.model.DatosEmbalaje`** (record nuevo), con los ocho campos como `String` —
son texto que se imprime tal cual, no se opera con ellos:

```java
public record DatosEmbalaje(String nroBolsa, String nombreCaja, String nroCaja,
                            String pluribol, String cantPluribol,
                            String rollo, String cantPanos, String observaciones) {
    public static final DatosEmbalaje VACIO = ...;
    public boolean tieneCajaOBolsa();
}
```

**`MedidaSku`**: el campo `String embalaje` pasa a `DatosEmbalaje embalaje`, nunca null (`VACIO`
cuando el SKU no tiene ninguna columna cargada).

## Render de las líneas

**`ar.com.leo.etiquetas.parser.EmbalajeRenderer`** (reemplaza a `EmbalajeResolver`), lógica pura:

```java
public static List<String> lineas(DatosEmbalaje datos)
public static String campoZpl(List<String> lineas)
```

| Condición | Línea |
|---|---|
| N° Caja y Nombre Caja cargados | `CAJA: 3 - GRANDE` |
| solo uno de los dos | `CAJA: 3` / `CAJA: GRANDE` |
| N° Bolsa cargado | `BOLSA: 5` |
| ni caja ni bolsa | `NO ESTANDARIZADO` |
| PLURIBOL cargado, con cantidad | `PLURIBOL: SI - 2 vueltas` |
| PLURIBOL cargado, sin cantidad | `PLURIBOL: SI` |
| ROLLO cargado, con cantidad | `ROLLO: DIAMANTE - 3 paños` |
| ROLLO cargado, sin cantidad | `ROLLO: DIAMANTE` |
| OBSERVACIONES cargado | `OBS: Colchon + Tapa` |

La primera línea siempre está: si el SKU no tiene ni caja ni bolsa se imprime `NO ESTANDARIZADO`,
para que el operario distinga "todavía no le cargaron el embalaje" de una etiqueta generada sin esta
función, en vez de embalar a criterio propio.

El orden es siempre ese. **Caja y bolsa son excluyentes en el render**: si las dos están cargadas
—pasa cuando cambia el embalaje y queda el número de bolsa viejo— gana la caja. Sin esa regla
saldrían cinco líneas y la quinta se imprimiría sobre el separador de la zona de picking y el
título del producto. Así el máximo es siempre de 4 líneas.

Los valores se colapsan antes de imprimirse (saltos de línea y espacios repetidos): el usuario
puede haber usado Alt+Enter en OBSERVACIONES, y un LF crudo dentro de un `^FD` pega las palabras
porque la impresora lo ignora.

`campoZpl` neutraliza `^` y `~` en todos los valores: son prefijos de comando de ZPL y dentro de un
`^FD` cortarían el campo, haciendo que el resto de la etiqueta se interprete como comandos. El texto
lo tipea el usuario en el Excel, así que no es confiable.

## Inyección en la etiqueta

Bloque de hasta 4 líneas en el margen superior derecho, dentro del mismo bloque `^LH0,0` que ya se
inyecta al inicio de cada etiqueta:

```
^FO410,83^A0N,22,22^FB380,1,0,L^FDCAJA: 3 - GRANDE^FS
^FO410,107^A0N,22,22^FB380,1,0,L^FDPLURIBOL: SI - 2 vueltas^FS
^FO410,131^A0N,22,22^FB380,1,0,L^FDROLLO: DIAMANTE - 3 paños^FS
^FO410,155^A0N,22,22^FB380,1,0,L^FDOBS: Colchon + Tapa^FS
```

`x=410`, primera línea en `y=20`, paso de 24px, fuente 22. Entran hasta 6 líneas: la última llega a
y≈162, justo antes del separador de la zona de picking (y=180).

**La última línea recibe todo el alto libre que sobra** (`MAX_LINEAS - índice`), porque es la de
observaciones, el único texto que puede no entrar en una línea: con `^FB` de una sola línea ZPL no
descarta el sobrante, lo reimprime encima de la misma y queda ilegible.

**El banner MEDIR se reubica** debajo del `#N` (`^FO20,70^GB380,52,52`, o sea x 20–400 y 70–122),
entre el número de posición (termina en y=65) y el `Pack ID` de ML (empieza en y=129). Antes ocupaba
el margen superior derecho, que ahora se necesita entero para las observaciones largas. El `^FB340,1` acota cada línea al ancho
disponible: un valor largo se recorta en vez de derramarse fuera del área imprimible.

**Por qué ahí.** A la derecha de x=410 no hay ningún campo de ML entre y=0 y y=180 salvo el texto
que se elimina. El límite por la izquierda lo marca el bloque `Pack ID: ...`, cuyo número llega
hasta x≈400 en las filas y 129–160.

**Hay que borrar el texto de ML** *"Recortá esta parte de la etiqueta para que tu paquete viaje
seguro"* (`^FO450,30` bajo `^LH0,90`, o sea y≈120–160), que cae justo en esa zona y no le sirve al
operario. Se elimina localizando el campo por el fragmento `ecort` —sin acentos ni mayúsculas, para
que no dependa de cómo ML codifique la tilde— y cortando desde su `^FO` hasta su `^FS`.

Si ML cambia ese texto y el ancla no aparece, se registra una advertencia en el log y las líneas se
dibujan igual, superpuestas a ese texto. Es visible y no silencioso: preferible a no mostrar el
embalaje.

**Alcance:** solo zonas distintas de `CARROS` y solo SKU numéricos, la misma guarda que usa el banner
MEDIR. Los SKU no numéricos son sentinelas del parser (`SKU INVALIDO: ...`) que nunca llegan al Excel
de medidas, así que no se les puede cargar un embalaje.

## Aviso de pendientes

Un SKU se reporta como pendiente cuando **no tiene ni caja ni bolsa**; pluribol, rollo y
observaciones no cuentan. El diálogo post-proceso suma:

```
3 SKU(s) sin caja/bolsa asignada en este lote:
  1241212, 1241255, 998877
```

Desaparece la categoría de "código inexistente": existía solo por el catálogo.

Si el módulo de medidas está apagado, la ruta está vacía o el archivo no es usable, no se inyecta
nada y no se avisa nada — la etiqueta sale como antes de esta función.

Si el Excel tiene el módulo activo pero le faltan las columnas de embalaje, todas las etiquetas
salen con `NO ESTANDARIZADO` y el aviso lista el lote entero. Es intencional: el archivo de
producción ya tiene las columnas, y un aviso ruidoso es preferible a que la función se apague sola
sin que nadie lo note.

## Qué se elimina

Del diseño anterior desaparecen por completo:

- `Embalaje` (record del catálogo) y `EmbalajeResolver`.
- Hoja `EMBALAJES`: `HOJA_EMBALAJES`, `HEADERS_EMBALAJES`, `asegurarHojaEmbalajes`,
  `crearEstiloHeaderCatalogo`, `leerCatalogoDe`.
- Validación y desplegable: `aplicarValidacionEmbalaje`, `yaTieneValidacion`, `quitarValidaciones`,
  `MAX_FILAS_CATALOGO`, `MAX_FILAS_VALIDACION`.
- Migración de estructura: `asegurarEstructuraEmbalajes`, `desplazarColumnasALaDerecha`,
  `copiarCelda`, `asegurarColumnaEmbalaje`, `hojaMedidasParaEscribir`, `tieneEstructuraCompleta`,
  y el campo `estructuraCompleta` de `DatosMedidas`.
- La columna `EMBALAJE` y su constante `COL_EMBALAJE`.

`normalizarHeader` deja de delegar en `EmbalajeResolver` y recupera su implementación propia.

**El usuario limpia su Excel a mano**: borra la columna `EMBALAJE`, la hoja `EMBALAJES` y la
validación, y agrega las 8 columnas nuevas. La app no lo hace por él.

Al crear un Excel **desde cero** (archivo inexistente), `HEADERS` incluye las 8 columnas nuevas
entre `SUBIDO` y `ERROR`, para que el archivo salga completo y en el orden pedido. `COL_ERROR` pasa
a 19 en consecuencia. No es migración: no toca archivos existentes, y como `ERROR` se ubica por su
encabezado, un archivo donde quedó en otra posición sigue funcionando.

`agregarPendientes` **no escribe nada** en las 8 columnas al insertar la fila de un SKU nuevo: son
de carga manual. Las celdas quedan vacías, sin estilo de "falta cargar", para no ensuciar el archivo
con formato sobre columnas que la app no administra.

## Testing

**`EmbalajeRendererTest`** (nuevo): cada línea por separado, caja con y sin nombre, cantidades
vacías, orden de las líneas, SKU sin ningún dato (lista vacía), sanitizado de `^`/`~`, y el `^FB` en
el campo ZPL.

**`MedidasExcelManagerTest`** (se reescribe la parte de embalaje): lectura de las ocho columnas,
archivo sin ninguna de ellas (`DatosEmbalaje.VACIO`), columnas en orden distinto, y variantes de
encabezado (`PAÑOS`/`PANOS`, mayúsculas y espacios).

Se eliminan los tests de migración, validación, desplegable, hoja catálogo y desplazamiento de
columnas.

Como antes, el punto donde el fragmento ZPL se concatena dentro de `injectZplHeaders` no queda
cubierto: vive en un método privado de `MainController` que necesita JavaFX inicializado. Lo que sí
se cubre es la generación del fragmento. El borrado del texto de ML tampoco: depende del mismo
método.

## Archivos afectados

| Archivo | Cambio |
|---|---|
| `etiquetas/model/DatosEmbalaje.java` | **Nuevo** |
| `etiquetas/model/Embalaje.java` | **Se elimina** |
| `etiquetas/model/MedidaSku.java` | `embalaje` pasa de `String` a `DatosEmbalaje` |
| `etiquetas/parser/EmbalajeRenderer.java` | **Nuevo** — reemplaza a `EmbalajeResolver` |
| `etiquetas/parser/EmbalajeResolver.java` | **Se elimina** |
| `etiquetas/parser/MedidasExcelManager.java` | Lee las 8 columnas; se le quita toda la escritura de estructura |
| `ui/MainController.java` | Inyecta el bloque, borra el texto de ML, reporta los SKU sin caja ni bolsa |
| `test/.../EmbalajeRendererTest.java` | **Nuevo** |
| `test/.../EmbalajeResolverTest.java` | **Se elimina** |
| `test/.../MedidasExcelManagerTest.java` | Se reescribe la parte de embalaje |
| `README.md` | Columnas nuevas, líneas en la etiqueta, se quita la hoja `EMBALAJES` |
