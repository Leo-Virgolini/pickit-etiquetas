# Columnas MEDIDAS y EMBALAJE en la tabla de etiquetas — Diseño

**Fecha:** 2026-08-12
**Autor:** Leo (con asistencia de Claude)

## Resumen

La tabla de etiquetas descargadas suma dos columnas, `MEDIDAS` y `EMBALAJE`, que muestran de un
vistazo qué SKU del lote tienen su información cargada en el Excel de medidas. Hoy eso solo se sabe
al final del proceso, en un diálogo modal que se cierra y no se puede volver a mirar.

Es solo presentación: no cambia lo que se imprime, lo que se sube a ML ni lo que se escribe en el
Excel.

## Estado por SKU

Enum **`ar.com.leo.etiquetas.ui.EstadoDato`** con tres valores y su texto:

| Estado | Texto | Cuándo |
|---|---|---|
| `SI` | `✓ SI` | El dato está cargado |
| `NO` | `✘ NO` | Falta cargarlo |
| `NO_APLICA` | `—` | Ese SKU nunca lleva ese dato |

`NO_APLICA` cubre las filas de zona `CARROS` (listan varios productos), los SKU no numéricos
—sentinelas `SKU INVALIDO: ...` del parser, que nunca llegan al Excel— y el caso de módulo de
medidas apagado.

Criterios, los mismos que ya usan la etiqueta y el aviso post-proceso:

- **MEDIDAS**: `MedidaSku.estaMedido()` — las cuatro columnas base cm/kg cargadas.
- **EMBALAJE**: `DatosEmbalaje.tieneCajaOBolsa()` — los agregados (pluribol, rollo, observaciones)
  no cuentan, igual que en la línea `NO ESTANDARIZADO`.

Un SKU que no figura en el Excel cuenta como `NO` en ambas: es elegible, solo que todavía no está
cargado.

## Regla de elegibilidad compartida

La condición de "SKU elegible" (no CARROS, no vacío, sin saltos de línea, numérico) hoy vive dentro
de `injectZplHeaders`. Se extrae a un método que usan tanto la inyección en la etiqueta como el
cálculo de la tabla.

No es refactor gratuito: es la duplicación que el último code review marcó entre los bloques MEDIR y
embalaje, que ya había divergido una vez. Si la tabla tuviera su propia copia, podría decir `NO`
sobre un SKU que la etiqueta considera no elegible, o al revés.

## Tabla

`LabelTableRow` suma dos campos `EstadoDato` con sus getters y properties, siguiendo el patrón de
los campos existentes.

`displayResult` pasa a recibir el mapa de medidas —los dos llamadores ya lo tienen a mano— y
resuelve el estado de cada grupo al construir la fila.

**FXML** (`MainView.fxml`): dos `TableColumn` nuevas al final de `labelTable`, después de
`Cantidad`, con ancho fijo chico (min 70, pref 90). Van al final para no correr las columnas de
contenido, que son las que se leen al operar.

**Celdas**: `cellFactory` propia, centrada, que pinta el `NO` con fondo rosa pálido (`#FEE2E2`) y
texto rojo oscuro (`#991B1B`) en negrita — los mismos colores que la columna ERROR del Excel de
medidas, para que el código de color sea el mismo en la app y en la planilla. El `SI` va en verde
oscuro y el `—` en gris claro, ambos sin fondo.

## Testing

**`EstadoDatoTest`** (nuevo): el texto de cada estado, y la resolución a partir de un `MedidaSku`
—medido y sin medir, con caja, con bolsa, sin ninguna de las dos— más el caso de SKU ausente del
Excel y el de módulo apagado.

El pintado de la celda no se testea: es JavaFX y necesita el toolkit inicializado, igual que el
resto de la UI.

## Archivos afectados

| Archivo | Cambio |
|---|---|
| `etiquetas/ui/EstadoDato.java` | **Nuevo** — enum y resolución del estado |
| `etiquetas/ui/LabelTableRow.java` | Dos campos nuevos |
| `ui/MainController.java` | Elegibilidad compartida, `displayResult` con medidas, cell factory |
| `resources/ar/com/leo/ui/MainView.fxml` | Dos columnas nuevas |
| `test/.../EstadoDatoTest.java` | **Nuevo** |
| `README.md` | Las dos columnas en la descripción de la tabla |
