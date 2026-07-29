# Mejoras sección "Etiquetas ML" — Diseño

**Fecha:** 2026-07-29
**Autor:** Leo (con asistencia de Claude)

## Resumen

Cuatro cambios en la sección **Etiquetas ML** de la app JavaFX:

0. Renombrar la columna "Despacho" → "Despacho límite" en la tabla de órdenes.
1. Agregar una columna "#" en la tabla de etiquetas con el número físico de impresión.
2. Reordenar los carros: primero los de ≤2 SKUs distintos (juntos), luego los de 3+ (juntos).
3. Agregar filtros por tipo de envío (Flex / Colecta / Turbo) en la UI.

Ninguno requiere cambios de arquitectura; son extensiones puntuales sobre el flujo existente.

## Contexto del código

- **Sub-tab "API MercadoLibre"**: `onFetchMeliOrders` trae órdenes → `displayOrders` puebla `orderTable` (`OrderTableRow`). El usuario selecciona órdenes → "Descargar Etiquetas" → `labelSorter.sort` + `injectZplHeaders` → `displayResult` puebla `labelTable` (`LabelTableRow`).
- **Sub-tab "Archivo Local"**: parsea ZPL → mismo pipeline de sort → `labelTable`.
- **Orden físico de impresión**: `interleaveForPrint(result.sortedFlatList())`. El intercalado acordeón está diseñado para que, tras plegar y cortar, las etiquetas queden en orden secuencial. Por lo tanto el orden físico final = orden de los grupos aplanados (`sortedFlatList`), que es el mismo orden en que `displayResult` puebla la tabla.
- `injectZplHeaders` preserva el orden de los grupos que produce `LabelSorter.sort` (itera `result.groups()` en orden y reconstruye `newGroups` en el mismo orden). Entonces el orden final lo define `LabelSorter.sort`.
- El tipo de envío proviene del shipment (`GET /shipments/{id}`), que la app **ya descarga** en `MercadoLibreAPI.obtenerSla` para detectar turbo y SLA. La respuesta (sin header `X-Format-New`) trae el campo plano `logistic_type`. Cero requests adicionales.

## Cambio 0 — Renombrar columna "Despacho" → "Despacho límite"

**Archivo:** `src/main/resources/ar/com/leo/ui/MainView.fxml`

- `orderSlaCol`: cambiar `text="Despacho"` → `text="Despacho límite"`.
- Es la única tabla con esa columna. Los labels "Despacho ML:" de los combos/radios no se tocan.

## Cambio 1 — Columna "#" (número físico de impresión)

**Semántica:** número de la etiqueta en la pila física impresa (orden secuencial en `sortedFlatList`). Cada fila de `labelTable` es un **grupo** que puede contener varias etiquetas físicas (`group.labels().size()`):
- Grupo de 1 etiqueta → un número (ej. `7`).
- Grupo de N etiquetas → rango (ej. `5–12`).

El número es sobre el **lote completo** (igual que el archivo auto-guardado). Nota conocida: si en el diálogo de impresión se eligen solo algunas zonas, los números siguen identificando la etiqueta pero no serán contiguos. Aceptable.

**Archivos:**

- `src/main/java/ar/com/leo/etiquetas/ui/LabelTableRow.java`
  - Agregar campo `printNumber` (`StringProperty`), su getter `getPrintNumber()`, `printNumberProperty()`, y sumarlo al constructor como primer parámetro.
- `src/main/resources/ar/com/leo/ui/MainView.fxml`
  - Agregar como **primera** columna de `labelTable`: `<TableColumn fx:id="printNumCol" text="#" minWidth="45" maxWidth="70" prefWidth="55" />`.
- `src/main/java/ar/com/leo/ui/MainController.java`
  - Declarar `@FXML private TableColumn<LabelTableRow, String> printNumCol;`.
  - En `initialize()`: `printNumCol.setCellValueFactory(new PropertyValueFactory<>("printNumber"));` + `printNumCol.setCellFactory(col -> centeredCell());`.
  - En `displayResult(SortResult result)`: llevar un contador acumulado. Para cada grupo, `count = group.labels().size()`; `start = acumulado + 1`; `end = acumulado + count`; `printNumber = (count == 1) ? String.valueOf(start) : (start + "–" + end)`; `acumulado += count`. Pasar `printNumber` como primer argumento a `new LabelTableRow(...)`.

**Testing:** verificar visualmente que la suma de rangos cubre 1..totalEtiquetas sin huecos ni solapamientos, y que grupos de 1 etiqueta muestran un solo número.

## Cambio 2 — Reordenar carros (≤2 SKUs juntos, luego 3+ juntos)

**Definición:** cantidad de **códigos SKU distintos** del carro. Dentro de la zona CARROS: bloque de ≤2 distintos primero, luego bloque de 3+. El resto de las zonas no se altera.

**Archivos:**

- `src/main/java/ar/com/leo/etiquetas/sorter/LabelSorter.java`
  - Agregar helper `private int carrosBucket(SortedLabelGroup g)`: si `!"CARROS".equals(g.zone())` → `0` (no diferencia). Si es CARROS: contar líneas distintas no-blancas de `g.sku()` (`sku.split("\n")`, trim, filtrar vacías, `distinct().count()`); devolver `0` si `distinct <= 2`, `1` si `>= 3`.
  - En el `Comparator` de `sort(...)`, insertar `.thenComparingInt(this::carrosBucket)` **después** de `zoneGroupPriority` y del `zone` string, y **antes** del comparator por `sku`. Así solo subdivide dentro de CARROS y mantiene el orden por SKU dentro de cada bloque.
- `src/main/java/ar/com/leo/ui/MainController.java`
  - En el comparator de `orderTable` (~línea 1429), agregar la misma lógica. Como `OrderTableRow` no expone `labels`, contar SKUs distintos desde `r.getSku()` (multilínea, mismo criterio). Insertar `.thenComparingInt(r -> carrosBucketOrders(r))` en la posición análoga (tras zona-prioridad y zona-string, antes de `getSku`). Helper local que devuelve 0 para no-CARROS.

**Efecto:** al arreglar `LabelSorter.sort` se corrige simultáneamente `labelTable`, impresión directa (`onPrintDirect` usa `currentResult.groups()`) y archivo auto-guardado (`sortedFlatList`). El comparator de `orderTable` cubre la tabla de órdenes.

**Testing:** con un lote que tenga carros de 2 y de 3+ SKUs, verificar que en ambas tablas y en el archivo generado aparecen todos los de 2 juntos y luego todos los de 3+ juntos.

## Cambio 3 — Filtros por tipo de envío (Flex / Colecta / Turbo)

**Ubicación:** sub-tab "API MercadoLibre", junto a los combos Estado/Despacho. Filtro **en vivo** (client-side) sobre `orderTable`, combinado con el buscador. No re-consulta ML.

**Mapeo (estricto):**
- **Flex** = `logistic_type == "self_service"` **y** no turbo.
- **Turbo** = tag `turbo` (ya detectado en `SlaInfo.turbo()`).
- **Colecta** = `logistic_type == "cross_docking"` (solo cross_docking; Places/`xd_drop_off` NO se incluye).

**Semántica del filtro:**
- **Ningún checkbox tildado = sin filtro → muestra TODOS** los tipos (incluidos Places, Drop off/`drop_off`, Full/`fulfillment`).
- **Uno o más tildados** → muestra solo las filas cuyo tipo coincide con algún checkbox tildado; el resto se oculta.
- **Default: ninguno tildado** (muestra todo). Solo se filtra cuando el usuario lo pide explícitamente.

**Archivos:**

- `src/main/java/ar/com/leo/api/ml/MercadoLibreAPI.java`
  - `record SlaInfo(String status, OffsetDateTime expectedDate, boolean turbo)` → agregar `String logisticType`.
  - En `obtenerSla`, del mismo JSON del shipment (`root`) que ya se lee para el tag turbo, leer `root.path("logistic_type").asString("")` y pasarlo al `new SlaInfo(...)`. Actualizar todos los sitios que construyen `SlaInfo`.
- `src/main/java/ar/com/leo/etiquetas/ui/OrderTableRow.java`
  - Agregar campo `shippingType` (`String`, plano — no necesita property; solo se usa para filtrar). Valores del enum lógico: `FLEX`, `COLECTA`, `TURBO`, `OTRO`. Getter `getShippingType()`. Sumarlo al constructor.
- `src/main/java/ar/com/leo/ui/MainController.java`
  - Helper `private String resolveShippingType(SlaInfo sla)`: si `sla == null` → `OTRO`; si `sla.turbo()` → `TURBO`; si `"self_service".equals(sla.logisticType())` → `FLEX`; si `"cross_docking".equals(sla.logisticType())` → `COLECTA`; else `OTRO`.
  - En `displayOrders` (donde se arma cada `OrderTableRow`), calcular el `shippingType` a partir del `slaMap` del/los shipment(s) del grupo y pasarlo al constructor. (Un grupo/orderTableRow = un shipment en el flujo normal; si hubiera varios, tomar el del primer shipment con SLA, consistente con cómo ya se toma status/slaDate.)
  - FXML: agregar 3 `CheckBox` (`filterFlexCheck`, `filterColectaCheck`, `filterTurboCheck`) en el `HBox` del sub-tab API (línea ~111-117), con labels "⚡ Turbo" / "🚚 Colecta" / "📦 Flex" (íconos a definir), **sin** `selected` (default destildados).
  - Declarar los 3 `@FXML CheckBox` + listeners que llamen a un único `applyOrderFilters()`.
  - Refactor del predicate: hoy el `searchField` listener setea `filteredOrders.setPredicate(...)`. Extraer la lógica combinada a `applyOrderFilters()` que arma un predicate = `matchesSearch(row) && matchesShippingFilter(row)`:
    - `matchesSearch`: la lógica actual del buscador.
    - `matchesShippingFilter`: si los 3 checkboxes están destildados → `true` (sin filtro). Si no, `true` sólo si el `shippingType` de la fila está en el conjunto de tipos tildados.
    - Llamar `applyOrderFilters()` desde el listener del `searchField` y desde los 3 listeners de checkboxes.
  - Deshabilitar los 3 checkboxes durante loading (junto a `estadoFilterCombo.setDisable(loading)` ~línea 1296), por consistencia.

**Testing:** con un lote mixto, verificar: (a) sin tildar nada → aparecen todos incluidos Places/Full; (b) tildar solo Turbo → solo turbo; (c) tildar Flex+Colecta → ambos, ocultando turbo y otros; (d) el buscador sigue funcionando combinado con los checkboxes.

## Orden de implementación sugerido

1. Cambio 0 (rename) — trivial, aislado.
2. Cambio 1 (columna "#") — `LabelTableRow` + FXML + `displayResult`.
3. Cambio 2 (carros) — `LabelSorter` + comparator `orderTable`.
4. Cambio 3 (filtros) — `SlaInfo` + `OrderTableRow` + `displayOrders` + FXML + `applyOrderFilters`.

## Fuera de alcance

- Filtrar por tipo de envío en el request de ML (no soportado de forma confiable; el dato ya se obtiene client-side).
- Filtros de tipo de envío sobre `labelTable` (los grupos por SKU pueden mezclar tipos; el filtro va sobre órdenes).
- Colecta rápida (`xd_same_day`, solo Brasil).
