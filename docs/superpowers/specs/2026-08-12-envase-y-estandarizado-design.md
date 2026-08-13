# Envase único y columna ESTANDARIZADO — Diseño

**Fecha:** 2026-08-12
**Autor:** Leo (con asistencia de Claude)

Ajusta a [2026-08-11-datos-embalaje-por-sku-design.md](2026-08-11-datos-embalaje-por-sku-design.md)
sobre el Excel real de producción (`TAMAÑO EMBALAJE ML.xlsx`).

## Resumen

Las tres columnas de bolsa y caja se reemplazan por una sola, `ENVASE`, que guarda un código
(`BOL-1`, `CAJ-1`) resuelto contra una hoja de estandarización. Aparece además una columna
`ESTANDARIZADO`, calculada por una fórmula del usuario, que pasa a ser la única fuente de verdad
sobre si el embalaje de un SKU está cargado.

## Estructura real del Excel

**Hoja de medidas** (17 columnas):

`SKU | PRODUCTO | Largo cm | Ancho cm | Alto cm | Peso físico (empaque + producto) kg |
Largo +20% | Ancho +20% | Alto +20% | Peso físico (empaque + producto) +5% |
SUBIDO | ESTANDARIZADO | ENVASE | TIPO DE ROLLO | CANT PAÑOS | OBSERVACIONES | ERROR`

**Hoja `ESTANDARIZACION`**: `N° | LARGO | ANCHO | ALTO | INSCRIPCION | … | ROLLOS`, donde `N°` es el
código (`CAJ-1`, `BOL-1`) e `INSCRIPCION` el texto escrito en el envase físico (`9Y`, `AYUDIN`).

### Detección de las columnas de margen

El peso ya no lleva `+20%` sino `+5%`, así que buscar el literal `+20` deja de servir: esa columna
pasaría a leerse como el peso base y lo pisaría, rompiendo la subida a ML.

No alcanza con buscar `+`: el encabezado del peso base es `Peso físico (empaque + producto) kg` y
también lo tiene. Se usa el patrón `\+\d+\s*%`, que reconoce `+20%` y `+5%` y no matchea el `+` de
"empaque + producto".

## Modelo

`DatosEmbalaje` pasa a:

```java
public record DatosEmbalaje(String envase, String inscripcion, String rollo,
                            String cantPanos, String observaciones, boolean estandarizado)
```

`inscripcion` se resuelve al leer, buscando `envase` en la hoja `ESTANDARIZACION`. Si el código no
figura ahí, queda vacía y la etiqueta muestra solo el código: es un dato de referencia, no una
validación, y no vale la pena frenar por eso.

`tieneCajaOBolsa()` desaparece; su lugar lo toma `estandarizado`.

## Etiqueta

| Condición | Línea |
|---|---|
| `ESTANDARIZADO` ≠ SI | `NO ESTANDARIZADO`, en grande y negrita, y **es la única línea** |
| `ENVASE` con inscripción conocida | `ENVASE: BOL-1 - 9Y` |
| `ENVASE` sin inscripción en la hoja | `ENVASE: BOL-1` |
| `ENVASE` = `NO` | `ENVASE: NO` |
| inscripción `-` en la hoja | se ignora: las bolsas y la fila `NO` la traen así |
| `TIPO DE ROLLO` cargado, con paños > 0 | `ROLLO: DIAMANTES - 2 paños` (un solo paño va en singular) |
| `TIPO DE ROLLO` cargado, paños 0 o vacío | `ROLLO: DIAMANTES` / `ROLLO: NO` |
| `CANT PAÑOS` con texto libre (`2 Y 1`) | `ROLLO: MIXTO - 2 Y 1 paños`; si el texto ya dice "paño", no se repite |
| sin columna `ESTANDARIZADO` en el archivo | no se imprime nada y no se reclama nada |
| `OBSERVACIONES` cargado | `OBS: Colchon + Tapa` |

`ENVASE: NO` se imprime como cualquier otro valor: que el producto no lleve envase es una decisión
tomada, no un dato faltante. Lo que marca lo faltante es `ESTANDARIZADO`.

`TIPO DE ROLLO` = `NO` también se imprime: el operario tiene que ver que la decisión está tomada.

**Los paños solo aparecen si son mayores a cero.** En el Excel real la columna trae `0` en las filas
sin rollo, y `ROLLO: NO - 0 paños` sería ruido.

## Subida a ML

A las condiciones actuales (`SUBIDO` = NO) se suma que las cuatro columnas de margen sean **celdas
numéricas** y mayores a cero.

Numérica incluye **fórmula con resultado numérico**, que es el caso real: en el archivo de
producción las cuatro columnas son fórmulas. Se excluye el texto, aunque sea parseable como número,
y se excluyen las fórmulas con error — hoy hay una fila así en el archivo.

## Tabla y aviso

La columna `Embalaje` de la tabla pasa a llamarse **`Estandarizado`** y muestra el valor de esa
celda con los mismos `✓ SI` / `✘ NO` / `—`. El aviso al final del lote lista los SKU con `NO`.

## Testing

`EmbalajeRendererTest`: envase con y sin inscripción, `ENVASE: NO`, rollo en `NO`, paños en cero y
mayores a cero, y el aviso como única línea cuando no está estandarizado.

`MedidasExcelManagerTest`: las columnas nuevas leídas por encabezado, el reconocimiento de `+5%`
frente al `+` del peso base, la hoja `ESTANDARIZACION` ausente o con códigos que no matchean, y que
una celda de texto o una fórmula con error no habiliten la subida.

`EstadoDatoTest`: el estado sale de `estandarizado` y no de si hay envase.

## Archivos afectados

| Archivo | Cambio |
|---|---|
| `etiquetas/model/DatosEmbalaje.java` | Campos nuevos, sin `tieneCajaOBolsa` |
| `etiquetas/model/MedidaSku.java` | `tieneMedidasParaSubir` exige celdas numéricas |
| `etiquetas/parser/MedidasExcelManager.java` | Columnas nuevas, patrón de margen, hoja `ESTANDARIZACION` |
| `etiquetas/parser/EmbalajeRenderer.java` | Líneas de envase y rollo |
| `etiquetas/ui/EstadoDato.java` | El estado sale de `estandarizado` |
| `ui/MainController.java` | Columna renombrada, aviso |
| `resources/ar/com/leo/ui/MainView.fxml` | Columna renombrada |
| `README.md` | Columnas y líneas nuevas |

## Adenda 2026-08-13 — Alcance de la sección de embalaje

### Solo etiquetas de un producto suelto

La sección se imprime únicamente en etiquetas de **1 SKU y cantidad 1**. Con dos o más unidades el
operario no está embalando un producto suelto, así que la instrucción de envase no le sirve y le
saca lugar al resto de la etiqueta. Es el mismo criterio que ya usaba el banner MEDIR.

Antes se resolvía una vez por grupo y se inyectaba en todas sus etiquetas. Ahora las líneas se
siguen armando una vez por grupo —son iguales para todas— pero la inyección se decide por etiqueta
con `EstadoDato.llevaEmbalaje(zona, sku, cantidad)`.

Las etiquetas **turbo** siguen llevándola: se tratan como cualquier otra zona.

### SKU que no figura en el Excel

Un SKU elegible que todavía no tiene fila imprime `NO ESTANDARIZADO`. Antes no imprimía nada, que
era indistinguible de tener el embalaje resuelto.

Se lo agrega al Excel en ese mismo lote —eso ya funcionaba, por el flujo de pendientes de medición—
pero recién después alguien va a poder cargarle el envase, así que mientras tanto el operario tiene
que enterarse.

Con una salvaguarda: si el Excel no tiene la columna `ESTANDARIZADO`, la función no está en uso y no
hay nada que reclamar. Como un SKU ausente no trae ninguna fila de donde deducirlo, el dato se saca
de las filas que sí se leyeron (`EstadoDato.moduloEmbalajeActivo`), una vez por lote.

### Un solo origen para los dos consumidores

`EstadoDato.embalajeDe(medida, moduloActivo)` devuelve el `DatosEmbalaje` a usar, y tanto la etiqueta
como la columna de la tabla parten de ese objeto: `estandarizadoDe` pasa a recibirlo en vez del
`MedidaSku`. La constante nueva `DatosEmbalaje.SIN_CARGAR` (`aplica = true`, `estandarizado = false`)
representa al SKU ausente, frente a `VACIO`, que representa a la función apagada.

La tabla **no** mira la cantidad: una fila es un grupo que puede mezclar etiquetas de una unidad y de
varias, y la columna informa sobre el SKU, no sobre una etiqueta puntual.

### Testing

`EstadoDatoTest`: la cantidad decide (`1` sí, `2` no), un carro nunca lleva; el módulo activo se
deduce de las filas leídas; un SKU ausente da `SIN_CARGAR` con el módulo activo y `VACIO` sin él.
`EmbalajeRendererTest`: `SIN_CARGAR` produce `NO ESTANDARIZADO`.

La inyección en sí sigue sin cobertura automática —`injectZplHeaders` es privado y necesita JavaFX
inicializado—, así que hace falta una impresión de prueba real.

## Adenda 2026-08-13 (2) — Referencia, banner y geometría

Corrige la adenda anterior sobre la etiqueta real.

### Cantidad > 1: referencia en vez de nada

La regla de "solo etiquetas de una unidad" se afloja: con dos o más unidades la sección **sí** sale,
encabezada por `REFERENCIA` en negrita, porque el envase cargado sirve de orientación aunque el
operario no esté embalando un producto suelto. Lo que no sale es el reclamo: si ese SKU no tiene el
embalaje cargado, la etiqueta queda limpia en vez de imprimir el aviso.

La decisión se mudó de `EstadoDato.llevaEmbalaje` —que se elimina— a
`EmbalajeRenderer.lineas(datos, cantidad)`, que es donde se puede testear.

### El aviso final sigue lo impreso

`EmbalajeRenderer.avisaSinEstandarizar(lineas)` decide qué SKU entra al diálogo del final del lote:
se pregunta por lo que se imprimió y no por lo que dice el Excel. Un SKU cuyas etiquetas son todas
de 2+ unidades no aparece, porque en papel nunca se reclamó.

### `NO ESTANDARIZADO` como recuadro

Pasa al formato que tenía el banner MEDIR: `^GB390,52,52` relleno y el texto en video inverso (`^FR`)
centrado, fuente **38**. No 42: son 16 caracteres en mayúscula que a esa altura ocupan ~403 de los
390 disponibles, y `^FB` no descarta lo que no entra sino que lo reimprime encima.

### Banner MEDIR fuera de uso

Queda detrás de `MainController.BANNER_MEDIR = false`, con el código intacto. La **detección** de
pendientes sigue activa: es la que agrega los SKU nuevos al Excel y la que alimenta el diálogo.

### Geometría

El bloque se queda a la derecha: el `Pack ID: …` / `Venta ID: …` de ML ocupa `x=30..380, y=129..160`
y no se toca. Se corre lo que el margen permite y se agranda la fuente.

| | Antes | Ahora |
|---|---|---|
| x | 410 | 400 |
| ancho | 380 | 390 |
| fuente | 22 | 24 |
| alto de línea | 24 | 26 |
| caracteres por línea | 30 | 29 |
| líneas | 6 | 6 |

Los 20px entre el fin del número de ML (x≈380 con sus 11 dígitos en fuente 30) y el borde del bloque
son el margen por si alguna vez viene de 12 dígitos.

## Adenda 2026-08-13 (3) — Correcciones de review

### La columna Medidas sale de la tabla y del aviso

El banner MEDIR ya no se imprime, así que informar el estado de las medidas dejó de tener uso.
`EstadoDato.medidasDe` y la columna desaparecen. El diálogo del final del lote conserva cuántos SKU
nuevos se agregaron al Excel —son las filas que hay que completar con el envase— y lista los SKU sin
estandarizar. La detección de pendientes y el alta en el Excel siguen intactas.

### "¿La función está en uso?" deja de inferirse

`EstadoDato.moduloEmbalajeActivo` deducía la respuesta de las filas leídas, así que un mapa vacío
contestaba "apagada". Un Excel recién creado por la app trae la columna `ESTANDARIZADO` y ninguna
fila: en el primer lote —justo donde todos los SKU son nuevos y el aviso es el que más importa— no
se habría impreso ningún `NO ESTANDARIZADO`.

`MedidasExcelManager.leerMedidas` pasa a devolver `Medidas(porSku, embalajeEnUso)`, con el flag
tomado de donde realmente se sabe: `cols.estandarizado() != -1`. La inferencia se elimina.

### Otros dos

- `lineas(datos, cantidad)` descarta el encabezado `REFERENCIA` si quedó solo: la fórmula del usuario
  puede decir que sí sin que el archivo tenga las columnas de envase, rollo ni observaciones.
- Las anclas `Unidad` y `SKU:` se buscan solo en el ZPL de ML, a partir del final de lo inyectado.
  `OBS` es texto libre del Excel y una observación del tipo "2 Unidades por caja" hacía que ZONA se
  posicionara tomando como referencia la línea de embalaje.

## Adenda 2026-08-13 (4) — Formato del encabezado REFERENCIA

`REFERENCIA` pasa a la fuente residual **B** de Zebra (`^ABN,22,14`), subrayada, además de la negrita
que ya tenía. Se lee como un rótulo y no como un dato más.

Las fuentes bitmap de Zebra escalan solo en múltiplos enteros de su base (B es 11×7), así que 22×14
es el tamaño más cercano al cuerpo del bloque: dos píxeles más bajo que los 24 de `^A0`, con lo que
entra en el mismo alto de línea y no corre nada.

El subrayado se dibuja con `^GB`, porque ZPL no lo tiene como atributo. Con la fuente proporcional
del resto habría que estimar el ancho de la palabra; la bitmap es monoespaciada, así que sale exacto
—10 caracteres × 14 = 140— y la línea termina justo donde termina el texto.

```
^FO400,20^ABN,22,14^FDREFERENCIA^FS
^FO401,20^ABN,22,14^FDREFERENCIA^FS
^FO400,42^GB140,2,2^FS
^FO400,46^A0N,24,24^FB390,1,0,L^FDENVASE: BOL-1 - 9Y^FS
```

## Adenda 2026-08-13 (5) — Formato de la línea de envase

El separador `" - "` entre el código y la inscripción se reemplaza por comillas alrededor de la
inscripción: `ENVASE: CAJ-1 "9Y"` en vez de `ENVASE: CAJ-1 - 9Y`. Las comillas marcan que eso es lo
que está escrito en el envase físico, que es como lo busca el operario.

Sin inscripción va el código solo, sin comillas vacías colgando. El `" - "` de la línea de rollo
—`ROLLO: DIAMANTES - 2 paños`— no cambia.

## Adenda 2026-08-13 (6) — Fuente 26

El bloque se queda donde está —el `Pack ID` / `Venta ID` de ML sigue ocupando la franja de la
izquierda— y solo sube la fuente.

| | Antes | Ahora |
|---|---|---|
| fuente | 24 | 26 |
| alto de línea | 26 | 28 |
| caracteres por línea | 29 | 27 |
| líneas | 6 | 5 |

`x=400` no se toca: bajarlo se comería el margen reservado por si un `Pack ID` viene de 12 dígitos.

Con el Excel real —740 filas, 28 estandarizadas, 65 líneas— entra todo: se generaron los 56 bloques
posibles (cada SKU en etiqueta de 1 y de 2+ unidades) y ninguno se recorta. El caso más ajustado es
`OBS: 3 COLCHON 1 TAPA. AJUSTAR PAÑO SEGÚN CANT VENDIDA`, de 54 caracteres, que en una etiqueta de
2+ unidades —donde `REFERENCIA` ocupa una línea— entra en 2 filas de 27 justo al límite. Una
observación de 55 caracteres en ese mismo caso sí se recortaría.

`campoZpl` corta ahora las líneas que no entrarían antes del separador de la zona de picking. Con
cinco disponibles y cuatro como máximo real nunca se activa; está por si la cuenta cambia, porque
perder una línea es preferible a imprimir sobre la zona de picking.

`REFERENCIA` sigue en `^ABN,22,14`: las fuentes bitmap escalan en múltiplos enteros y el siguiente
—33×21— no entra en el alto de línea.

## Adenda 2026-08-13 (7) — Correcciones de review

### El recorte se hace por filas, no por caracteres

`truncar(texto, maxLineas · MAX_CARACTERES)` acotaba **caracteres**, pero `^FB` corta por
**palabras**: un texto que entra por cuenta de caracteres puede necesitar una fila más, y lo que no
entra ZPL lo reimprime encima de la última en vez de descartarlo — justo la mancha que ese recorte
existía para evitar.

La adenda (6) lo volvió alcanzable: al pasar el alto de línea de 26 a 28, en una etiqueta de 2+
unidades `OBS` bajó de 3 filas permitidas a 2. `OBS: 3 COLCHON 1 TAPA. AJUSTAR PAÑO SEGÚN CANT
VENDIDA` mide 54 = 2 × 27, así que pasaba el tope, pero al cortarse por palabras ocupa tres filas:
`OBS: 3 COLCHON 1 TAPA.` / `AJUSTAR PAÑO SEGÚN CANT` / `VENDIDA`.

`acotar(texto, maxFilas)` simula ahora el corte de `^FB` (`envolver`) y recorta por filas. La
verificación contra el Excel real se rehízo contando filas y no caracteres: de los 56 bloques
posibles, **0 desbordan** y **1 se recorta** —esa misma observación en etiqueta de 2+ unidades, que
pierde "VENDIDA"—. La verificación anterior, que solo miraba si aparecía `...`, no podía detectarlo.

### Otros

- Un alta fallida en el Excel dejaba el lote sin ningún diálogo: `agregarPendientes` traga la
  excepción y devuelve 0, y el diálogo no se muestra con 0 agregados y sin avisos. El caso habitual
  es tener el Excel abierto. Ahora se avisa.
- `MAX_PALABRA` se dimensiona por `"OBS: "` y no por el rótulo más largo (`"ENVASE: "`) a propósito:
  las palabras que hay que partir son códigos y URLs, y aparecen en las observaciones, que es donde
  el alto escasea. Queda documentado con su costo.
- El ancho del subrayado de `REFERENCIA` sale de `AVANCE_REFERENCIA`, que asume que la matriz 11x7
  de la fuente B ya incluye el espacio entre caracteres. Si no lo incluyera, el avance sería 18 y el
  subrayado quedaría corto por unos dos caracteres. **Confirmar en la impresión de prueba.**
  *(Confirmado, y la suposición era errónea: ver la adenda 8.)*
- Locales muertos en el hilo de descarga, y los tests nuevos que habían quedado bajo el separador de
  helpers.

## Adenda 2026-08-13 (8) — Ancho del subrayado, medido en papel

La impresión de prueba mostró el subrayado llegando hasta la **C** de `REFEREN`**C**`IA`, o sea
cubriendo 8 de los 10 caracteres. Eso resuelve la duda de la adenda anterior: la matriz 11x7 de la
fuente B es **el glifo solo**, y el espacio entre caracteres son 2 dots aparte, escalados por la
misma magnificación (4 a ×2).

La cuenta cierra exacta: 8 glifos más sus 7 huecos son `8 · 14 + 7 · 4 = 140`, que era justo el
ancho que se estaba dibujando.

El ancho pasa a contar glifos y huecos, sin el hueco que sobraría al final:

```
10 · 14 + 9 · 4 = 176
```

`AVANCE_REFERENCIA` se reemplaza por `GAP_REFERENCIA = 4`, que es el dato que faltaba.

### Sobre usar `**` en vez del subrayado

Se descartó. Es convención de markdown: significa "negrita" para quien conoce el formato, y en una
etiqueta impresa son dos asteriscos y nada más. `REFERENCIA` ya va en negrita de verdad, así que
estaría marcando algo ya marcado. El asterisco sí funciona en papel como delimitador
(`*** REFERENCIA ***`, convención de tickets), pero eso es otra cosa que lo que se buscaba.

## Adenda 2026-08-13 (9) — El reparto en filas desperdiciaba ancho

Una impresión de prueba con una observación de dígitos seguidos mostró todas las filas cortando
antes del margen derecho.

La causa: `partirPalabrasLargas` cortaba las palabras largas en pedazos de `MAX_PALABRA` = 22 y los
unía con espacios, y como `^FB` corta por palabras metía **un pedazo por fila**. Cada fila de
continuación usaba 22 de los 27 caracteres disponibles.

`partirPalabrasLargas` y `truncar` se reemplazan por `envolver`, que reparte el texto en filas del
ancho real: la primera pieza comparte fila con el rótulo y entra lo que sobre, y las siguientes se
quedan con la fila entera. `MAX_PALABRA` desaparece.

Una palabra se parte **solo si no entra ni en una fila entera** —códigos, URLs—. Una palabra común
que no entra en lo que queda se baja entera: partir "AJUSTAR" en "AJUS TAR" para no dejar cuatro
caracteres de margen se lee mucho peor.

Las filas se unen con espacios y `^FB` reproduce el mismo reparto: una fila solo se cierra cuando lo
que sigue no entraba, así que la impresora tampoco lo puede subir.

Reverificado contra el Excel real: 58 bloques, **0 desbordes**.

### Lo que sigue quedando de margen, y por qué

`MAX_CARACTERES` es un promedio para una fuente proporcional, dimensionado para **mayúsculas**, que
es lo que tienen las observaciones reales. Los dígitos son bastante más angostos, así que una
observación de puros números deja margen visible a la derecha. Subir el número por ese caso haría
que las observaciones en mayúscula desbordaran, que es el error caro: `^FB` no descarta lo que no
entra, lo reimprime encima.

El espacio libre abajo del bloque tampoco es aprovechable: en una etiqueta de 2+ unidades a `OBS` le
quedan 22 puntos hasta el separador de picking y una fila necesita 26.

## Adenda 2026-08-13 (10) — 29 caracteres por fila

Medido sobre una impresión real a fuente 26: un dígito ocupa ~13,4 puntos, así que en los 390 de
ancho entran **29**; una mayúscula ocupa ~15, así que de esas entrarían 25.

Un número plano no puede servir a los dos. Se elige el de los dígitos —el que aprovecha el ancho—
sabiendo que una fila de puras mayúsculas puede pasarse, y que `^FB` no descarta lo que sobra sino
que lo reimprime encima.

Verificado con el Excel real: de las **284 filas** que produce, **ninguna supera los 390 puntos**. La
más ancha da 389 y es una observación de puros dígitos. Las observaciones reales, aunque estén en
mayúscula, llevan espacios —7 puntos contra 15— y eso baja bastante el promedio:
`AJUSTAR PAÑO SEGÚN CANT` son 321 puntos.

El caso que quedaría al filo es una tirada larga de mayúsculas sin espacios, que se parte en pedazos
del ancho de la fila: 29 mayúsculas serían 435 puntos. No es un riesgo nuevo de este cambio —con 27
ya eran 405— pero es más ancho. Si alguna vez aparece una fila ilegible, `MAX_CARACTERES` es el
número a bajar.
