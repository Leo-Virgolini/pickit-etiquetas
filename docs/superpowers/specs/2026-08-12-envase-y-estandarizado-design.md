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
