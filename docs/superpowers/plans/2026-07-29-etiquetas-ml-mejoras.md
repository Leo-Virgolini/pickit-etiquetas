# Mejoras Etiquetas ML — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Agregar a la sección Etiquetas ML: columna "#" con número físico de impresión, reordenamiento de carros por cantidad de SKUs, filtros por tipo de envío (Flex/Colecta/Turbo), y renombrar la columna "Despacho".

**Architecture:** Extensiones puntuales sobre el flujo existente (JavaFX + Maven). La lógica pura se extrae a clases utilitarias testeables con JUnit (`CarrosOrdering`, `PrintNumbering`, enum `ShippingType`); el cableado FXML/UI se verifica compilando y corriendo la app. Sin cambios de arquitectura.

**Tech Stack:** Java 25, JavaFX 26.0.1, Maven (surefire 3.5.5), Jackson 3, JUnit Jupiter 5.11.4 (nuevo, scope test).

## Global Constraints

- Java 25 (`maven.compiler.source/target=25`); JavaFX 26.0.1.
- Encoding UTF-8 en todo archivo fuente.
- No romper patrones existentes: helpers estáticos en `MainController`, records en `model`, cell factories con `centeredCell()`.
- Mapeo de tipos de envío (estricto, confirmado): Flex = `logistic_type == "self_service"` y no turbo; Turbo = tag `turbo`; Colecta = `logistic_type == "cross_docking"`; cualquier otro (`xd_drop_off`, `drop_off`, `fulfillment`, vacío) = `OTRO`.
- Semántica filtro: ningún checkbox tildado ⇒ muestra todos; uno o más tildados ⇒ solo esos tipos. Default: destildados.
- Carros: "cantidad de SKUs" = **códigos SKU distintos**. Bucket 0 = ≤2 distintos; bucket 1 = ≥3.
- Comandos: compilar `mvn -q -DskipTests compile`; compilar tests `mvn -q test-compile`; correr un test `mvn -q -Dtest=Clase#metodo test`; correr toda la suite `mvn -q test`; correr la app `mvn javafx:run`.
- ⚠️ **NUNCA descargar etiquetas de ML durante las pruebas.** El endpoint `GET /shipment_labels` (botón "Descargar Etiquetas") tiene como efecto colateral marcar las órdenes como impresas (`ready_to_print` → `printed`) en ML. En verificación manual: "Obtener Órdenes" es seguro (solo lectura: búsqueda + SLA); para probar la tabla de etiquetas usar el sub-tab **Archivo Local** con un `.txt` ZPL de ejemplo, nunca la descarga por API.

---

### Task 1: Renombrar columna "Despacho" → "Despacho límite"

**Files:**
- Modify: `src/main/resources/ar/com/leo/ui/MainView.fxml:163`

**Interfaces:**
- Consumes: nada.
- Produces: nada (cambio de texto aislado).

- [ ] **Step 1: Editar el texto de la columna**

En `MainView.fxml`, la columna `orderSlaCol`:

```xml
<TableColumn fx:id="orderSlaCol" minWidth="100" prefWidth="140" text="Despacho límite" />
```

(Antes decía `text="Despacho"`.)

- [ ] **Step 2: Compilar**

Run: `mvn -q -DskipTests compile`
Expected: BUILD SUCCESS.

- [ ] **Step 3: Commit**

```bash
git add src/main/resources/ar/com/leo/ui/MainView.fxml
git commit -m "feat(etiquetas): renombrar columna Despacho a Despacho limite"
```

---

### Task 2: Agregar JUnit + utilitario `CarrosOrdering`

**Files:**
- Modify: `pom.xml:79` (agregar dependencia después de guava, dentro de `<dependencies>`)
- Create: `src/main/java/ar/com/leo/etiquetas/sorter/CarrosOrdering.java`
- Test: `src/test/java/ar/com/leo/etiquetas/sorter/CarrosOrderingTest.java`

**Interfaces:**
- Consumes: nada.
- Produces:
  - `CarrosOrdering.distinctSkuCount(String skuMultiline) -> int` (cuenta códigos SKU distintos no-blancos separados por `\n`; `null`/vacío ⇒ 0).
  - `CarrosOrdering.bucket(String zone, String skuMultiline) -> int` (0 si `zone` no es exactamente `"CARROS"`; si es CARROS: 0 cuando distinct ≤ 2, 1 cuando ≥ 3).

- [ ] **Step 1: Agregar la dependencia de JUnit al pom**

Dentro de `<dependencies>`, después del bloque de guava (línea ~79):

```xml
        <!-- JUnit 5 (tests) -->
        <dependency>
            <groupId>org.junit.jupiter</groupId>
            <artifactId>junit-jupiter</artifactId>
            <version>5.11.4</version>
            <scope>test</scope>
        </dependency>
```

- [ ] **Step 2: Escribir el test que falla**

Create `src/test/java/ar/com/leo/etiquetas/sorter/CarrosOrderingTest.java`:

```java
package ar.com.leo.etiquetas.sorter;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

class CarrosOrderingTest {

    @Test
    void distinctSkuCountCuentaCodigosDistintos() {
        assertEquals(2, CarrosOrdering.distinctSkuCount("10\n20\n10"));
        assertEquals(3, CarrosOrdering.distinctSkuCount("10\n20\n30"));
        assertEquals(1, CarrosOrdering.distinctSkuCount("10"));
        assertEquals(0, CarrosOrdering.distinctSkuCount(""));
        assertEquals(0, CarrosOrdering.distinctSkuCount(null));
    }

    @Test
    void bucketDevuelve0ParaZonasNoCarros() {
        assertEquals(0, CarrosOrdering.bucket("J1", "10\n20\n30"));
        assertEquals(0, CarrosOrdering.bucket("TURBOS", "10"));
    }

    @Test
    void bucketSeparaCarrosPorCantidadDeSkus() {
        assertEquals(0, CarrosOrdering.bucket("CARROS", "10\n20"));      // 2 distintos
        assertEquals(0, CarrosOrdering.bucket("CARROS", "10\n20\n10"));  // 2 distintos (con dup)
        assertEquals(1, CarrosOrdering.bucket("CARROS", "10\n20\n30"));  // 3 distintos
        assertEquals(1, CarrosOrdering.bucket("CARROS", "10\n20\n30\n40"));
    }
}
```

- [ ] **Step 3: Correr el test para verificar que falla**

Run: `mvn -q -Dtest=CarrosOrderingTest test`
Expected: FAIL (error de compilación: `CarrosOrdering` no existe).

- [ ] **Step 4: Implementar `CarrosOrdering`**

Create `src/main/java/ar/com/leo/etiquetas/sorter/CarrosOrdering.java`:

```java
package ar.com.leo.etiquetas.sorter;

import java.util.Arrays;

/**
 * Lógica de sub-ordenamiento de carros por cantidad de SKUs distintos.
 * Compartida entre {@link LabelSorter} (tabla de etiquetas / impresión) y el
 * comparador de la tabla de órdenes en la UI.
 */
public final class CarrosOrdering {

    private CarrosOrdering() {}

    /** Cuenta códigos SKU distintos no-blancos separados por '\n'. */
    public static int distinctSkuCount(String skuMultiline) {
        if (skuMultiline == null || skuMultiline.isBlank()) return 0;
        return (int) Arrays.stream(skuMultiline.split("\n"))
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .distinct()
                .count();
    }

    /**
     * Bucket de orden para carros: 0 para no-CARROS (no altera el orden),
     * 0 para carros de ≤2 SKUs distintos, 1 para carros de ≥3.
     */
    public static int bucket(String zone, String skuMultiline) {
        if (!"CARROS".equals(zone)) return 0;
        return distinctSkuCount(skuMultiline) <= 2 ? 0 : 1;
    }
}
```

- [ ] **Step 5: Correr el test para verificar que pasa**

Run: `mvn -q -Dtest=CarrosOrderingTest test`
Expected: PASS (Tests run: 3, Failures: 0).

- [ ] **Step 6: Commit**

```bash
git add pom.xml src/main/java/ar/com/leo/etiquetas/sorter/CarrosOrdering.java src/test/java/ar/com/leo/etiquetas/sorter/CarrosOrderingTest.java
git commit -m "feat(etiquetas): agregar JUnit y utilitario CarrosOrdering"
```

---

### Task 3: Aplicar orden de carros en `LabelSorter`

**Files:**
- Modify: `src/main/java/ar/com/leo/etiquetas/sorter/LabelSorter.java:55-64`
- Test: `src/test/java/ar/com/leo/etiquetas/sorter/LabelSorterCarrosTest.java`

**Interfaces:**
- Consumes: `CarrosOrdering.bucket(String, String)` de Task 2; `LabelSorter.sort(List<ZplLabel>, Map<String,String>) -> SortResult` (existente).
- Produces: `SortResult.groups()` con carros de ≤2 SKUs antes que los de ≥3, dentro de la zona CARROS.

- [ ] **Step 1: Escribir el test que falla**

Create `src/test/java/ar/com/leo/etiquetas/sorter/LabelSorterCarrosTest.java`:

```java
package ar.com.leo.etiquetas.sorter;

import ar.com.leo.etiquetas.model.SortResult;
import ar.com.leo.etiquetas.model.SortedLabelGroup;
import ar.com.leo.etiquetas.model.ZplLabel;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

class LabelSorterCarrosTest {

    @Test
    void carrosDe2SkusVanAntesQueLosDe3() {
        // carroTres: 3 SKUs distintos → CARROS bucket 1
        ZplLabel carroTres = new ZplLabel("^XA1^XZ", "10\n20\n30", "d", "det", 1);
        // carroDos: 2 SKUs distintos → CARROS bucket 0
        ZplLabel carroDos = new ZplLabel("^XA2^XZ", "40\n50", "d", "det", 1);
        // normal: zona J1 (prioridad 0, antes de CARROS)
        ZplLabel normal = new ZplLabel("^XA3^XZ", "99", "d", "det", 1);

        Map<String, String> skuToZone = Map.of("99", "J1");

        SortResult result = new LabelSorter().sort(List.of(carroTres, carroDos, normal), skuToZone);
        List<SortedLabelGroup> groups = result.groups();

        assertEquals("J1", groups.get(0).zone());
        assertEquals("CARROS", groups.get(1).zone());
        assertEquals("40\n50", groups.get(1).sku());      // carro de 2 primero
        assertEquals("CARROS", groups.get(2).zone());
        assertEquals("10\n20\n30", groups.get(2).sku());  // carro de 3 después
    }
}
```

- [ ] **Step 2: Correr el test para verificar que falla**

Run: `mvn -q -Dtest=LabelSorterCarrosTest test`
Expected: FAIL (los dos carros salen en orden de inserción; el de 3 SKUs queda antes que el de 2).

- [ ] **Step 3: Insertar el comparador de bucket en `LabelSorter.sort`**

En `LabelSorter.java`, dentro del `.sorted(Comparator...)` (líneas 55-64), agregar `.thenComparingInt(...)` **después** de `.thenComparing(g -> g.zone().toUpperCase())` y **antes** del comparador por SKU:

```java
                .sorted(Comparator
                        .<SortedLabelGroup>comparingInt(g -> zoneGroupPriority(g.zone()))
                        .thenComparing(g -> g.zone().toUpperCase())
                        .thenComparingInt(g -> CarrosOrdering.bucket(g.zone(), g.sku()))
                        .thenComparing(g -> {
                            try {
                                return Long.parseLong(g.sku());
                            } catch (NumberFormatException e) {
                                return Long.MAX_VALUE;
                            }
                        }))
```

(`CarrosOrdering` está en el mismo paquete, no requiere import.)

- [ ] **Step 4: Correr el test para verificar que pasa**

Run: `mvn -q -Dtest=LabelSorterCarrosTest test`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/main/java/ar/com/leo/etiquetas/sorter/LabelSorter.java src/test/java/ar/com/leo/etiquetas/sorter/LabelSorterCarrosTest.java
git commit -m "feat(etiquetas): ordenar carros por cantidad de SKUs en LabelSorter"
```

---

### Task 4: Aplicar orden de carros en el comparador de `orderTable`

**Files:**
- Modify: `src/main/java/ar/com/leo/ui/MainController.java:1429-1441`

**Interfaces:**
- Consumes: `CarrosOrdering.bucket(String, String)` de Task 2; `OrderTableRow.getZone()`, `OrderTableRow.getSku()` (existentes; `getSku()` devuelve SKUs multilínea separados por `\n`).
- Produces: filas de `orderTable` con carros de ≤2 SKUs antes que los de ≥3.

- [ ] **Step 1: Agregar el import**

En `MainController.java`, junto a los imports de `ar.com.leo.etiquetas.sorter.*`:

```java
import ar.com.leo.etiquetas.sorter.CarrosOrdering;
```

- [ ] **Step 2: Insertar el comparador de bucket**

En el `rows.sort(Comparator...)` (~línea 1429), agregar `.thenComparingInt(...)` después de `.thenComparing(r -> r.getZone().toUpperCase())` y antes de `.thenComparing(OrderTableRow::getSku)`:

```java
        rows.sort(Comparator
                .<OrderTableRow, Integer>comparing(r -> {
                    String z = r.getZone().toUpperCase();
                    if (z.startsWith("J")) return 0;
                    if (z.startsWith("TURBOS")) return 4;
                    if (z.startsWith("T")) return 1;
                    if (z.startsWith("COMBOS")) return 2;
                    if (z.startsWith("CARROS")) return 3;
                    if (z.startsWith("RETIROS")) return 5;
                    return Integer.MAX_VALUE;
                })
                .thenComparing(r -> r.getZone().toUpperCase())
                .thenComparingInt(r -> CarrosOrdering.bucket(r.getZone(), r.getSku()))
                .thenComparing(OrderTableRow::getSku));
```

- [ ] **Step 3: Compilar**

Run: `mvn -q -DskipTests compile`
Expected: BUILD SUCCESS.

- [ ] **Step 4: Verificar en la app (manual)**

Run: `mvn javafx:run`
Verificar en la tabla de Órdenes (sub-tab API, tras "Obtener Órdenes") que los carros de 2 SKUs aparecen juntos antes de los de 3+. Cerrar la app.

- [ ] **Step 5: Commit**

```bash
git add src/main/java/ar/com/leo/ui/MainController.java
git commit -m "feat(etiquetas): ordenar carros por cantidad de SKUs en tabla de ordenes"
```

---

### Task 5: Columna "#" con número físico de impresión

**Files:**
- Create: `src/main/java/ar/com/leo/etiquetas/model/PrintNumbering.java`
- Test: `src/test/java/ar/com/leo/etiquetas/model/PrintNumberingTest.java`
- Modify: `src/main/java/ar/com/leo/etiquetas/ui/LabelTableRow.java`
- Modify: `src/main/resources/ar/com/leo/ui/MainView.fxml:168-177`
- Modify: `src/main/java/ar/com/leo/ui/MainController.java` (declaración de columna ~116, `initialize()` ~231-264, `displayResult()` ~1980-1990)

**Interfaces:**
- Consumes: `SortedLabelGroup.labels()` (existente).
- Produces:
  - `PrintNumbering.compute(List<Integer> groupSizes) -> List<String>` (numera acumulativamente; tamaño 1 ⇒ `"N"`; tamaño >1 ⇒ `"start–end"` con guion largo `–`).
  - `LabelTableRow` con primer parámetro `String printNumber` + getter `getPrintNumber()` + `printNumberProperty()`.

- [ ] **Step 1: Escribir el test que falla**

Create `src/test/java/ar/com/leo/etiquetas/model/PrintNumberingTest.java`:

```java
package ar.com.leo.etiquetas.model;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class PrintNumberingTest {

    @Test
    void gruposDeUnaEtiquetaNumeranSecuencial() {
        assertEquals(List.of("1", "2", "3"), PrintNumbering.compute(List.of(1, 1, 1)));
    }

    @Test
    void gruposMultiEtiquetaMuestranRango() {
        // grupo de 1 → "1"; grupo de 4 → "2–5"; grupo de 2 → "6–7"
        assertEquals(List.of("1", "2–5", "6–7"), PrintNumbering.compute(List.of(1, 4, 2)));
    }

    @Test
    void listaVaciaDevuelveVacio() {
        assertEquals(List.of(), PrintNumbering.compute(List.of()));
    }
}
```

- [ ] **Step 2: Correr el test para verificar que falla**

Run: `mvn -q -Dtest=PrintNumberingTest test`
Expected: FAIL (compilación: `PrintNumbering` no existe).

- [ ] **Step 3: Implementar `PrintNumbering`**

Create `src/main/java/ar/com/leo/etiquetas/model/PrintNumbering.java`:

```java
package ar.com.leo.etiquetas.model;

import java.util.ArrayList;
import java.util.List;

/**
 * Calcula el número físico de impresión de cada grupo de etiquetas.
 * El intercalado acordeón está diseñado para que, tras plegar y cortar, las
 * etiquetas queden en el orden secuencial de esta numeración.
 */
public final class PrintNumbering {

    private PrintNumbering() {}

    /** Numera acumulativamente los grupos según su cantidad de etiquetas. */
    public static List<String> compute(List<Integer> groupSizes) {
        List<String> result = new ArrayList<>(groupSizes.size());
        int acumulado = 0;
        for (int size : groupSizes) {
            int start = acumulado + 1;
            int end = acumulado + size;
            result.add(size <= 1 ? String.valueOf(start) : (start + "–" + end));
            acumulado = end;
        }
        return result;
    }
}
```

- [ ] **Step 4: Correr el test para verificar que pasa**

Run: `mvn -q -Dtest=PrintNumberingTest test`
Expected: PASS.

- [ ] **Step 5: Agregar `printNumber` a `LabelTableRow`**

En `LabelTableRow.java`, agregar la propiedad como **primer** campo/param:

```java
    private final StringProperty printNumber;
    private final StringProperty orderIds;
    // ... resto igual

    public LabelTableRow(String printNumber, String orderIds, String zone, String sku,
                         String productDescription, String details, int quantity) {
        this.printNumber = new SimpleStringProperty(printNumber);
        this.orderIds = new SimpleStringProperty(orderIds);
        // ... resto igual
    }

    public StringProperty printNumberProperty() { return printNumber; }
    public String getPrintNumber() { return printNumber.get(); }
```

- [ ] **Step 6: Agregar la columna "#" al FXML**

En `MainView.fxml`, dentro de `<columns>` de `labelTable`, como **primera** columna (antes de `labelOrderCol`):

```xml
                                    <TableColumn fx:id="printNumCol" maxWidth="70" minWidth="45" prefWidth="55" text="#" />
```

- [ ] **Step 7: Declarar la columna y configurarla en `MainController`**

Declaración del campo `@FXML` (junto a `labelOrderCol`, ~línea 106):

```java
    @FXML
    private TableColumn<LabelTableRow, String> printNumCol;
```

En `initialize()`, junto a la configuración de las otras columnas de `labelTable` (~línea 231):

```java
        printNumCol.setCellValueFactory(new PropertyValueFactory<>("printNumber"));
        printNumCol.setCellFactory(col -> centeredCell());
```

- [ ] **Step 8: Calcular la numeración en `displayResult`**

En `displayResult(SortResult result)` (~líneas 1980-1990), reemplazar el armado de filas para usar `PrintNumbering`:

```java
    private void displayResult(SortResult result) {
        ObservableList<LabelTableRow> rows = FXCollections.observableArrayList();
        List<Integer> groupSizes = result.groups().stream()
                .map(g -> g.labels().size())
                .toList();
        List<String> printNumbers = ar.com.leo.etiquetas.model.PrintNumbering.compute(groupSizes);
        int i = 0;
        for (SortedLabelGroup group : result.groups()) {
            rows.add(new LabelTableRow(
                    printNumbers.get(i++),
                    group.orderIds(),
                    group.zone(),
                    group.sku(),
                    group.productDescription(),
                    group.details(),
                    extractQuantityFromLabels(group.labels())));
        }
        filteredLabels = new FilteredList<>(rows, p -> true);
        // ... resto del método sin cambios
```

- [ ] **Step 9: Compilar toda la suite y correr tests**

Run: `mvn -q test`
Expected: BUILD SUCCESS, todos los tests pasan.

- [ ] **Step 10: Verificar en la app (manual)**

Run: `mvn javafx:run`
Cargar etiquetas por el sub-tab **Archivo Local** (un `.txt` ZPL de ejemplo) — **no** usar la descarga por API (marca órdenes como impresas). Verificar que la columna "#" numera 1..N sin huecos; grupos con varias etiquetas muestran rango (ej. `5–12`). Cerrar la app.

- [ ] **Step 11: Commit**

```bash
git add src/main/java/ar/com/leo/etiquetas/model/PrintNumbering.java src/test/java/ar/com/leo/etiquetas/model/PrintNumberingTest.java src/main/java/ar/com/leo/etiquetas/ui/LabelTableRow.java src/main/resources/ar/com/leo/ui/MainView.fxml src/main/java/ar/com/leo/ui/MainController.java
git commit -m "feat(etiquetas): columna # con numero fisico de impresion"
```

---

### Task 6: Enum `ShippingType` (clasificación + filtro)

**Files:**
- Create: `src/main/java/ar/com/leo/api/ml/model/ShippingType.java`
- Test: `src/test/java/ar/com/leo/api/ml/model/ShippingTypeTest.java`

**Interfaces:**
- Consumes: nada.
- Produces:
  - `enum ShippingType { FLEX, COLECTA, TURBO, OTRO }`.
  - `ShippingType.from(boolean turbo, String logisticType) -> ShippingType`.
  - `ShippingType.passes(ShippingType type, Set<ShippingType> checked) -> boolean` (checked vacío ⇒ true; si no ⇒ `checked.contains(type)`).

- [ ] **Step 1: Escribir el test que falla**

Create `src/test/java/ar/com/leo/api/ml/model/ShippingTypeTest.java`:

```java
package ar.com.leo.api.ml.model;

import org.junit.jupiter.api.Test;

import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ShippingTypeTest {

    @Test
    void fromClasificaSegunTurboYLogisticType() {
        assertEquals(ShippingType.TURBO, ShippingType.from(true, "self_service"));
        assertEquals(ShippingType.TURBO, ShippingType.from(true, "cross_docking"));
        assertEquals(ShippingType.FLEX, ShippingType.from(false, "self_service"));
        assertEquals(ShippingType.COLECTA, ShippingType.from(false, "cross_docking"));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, "xd_drop_off"));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, "drop_off"));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, "fulfillment"));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, ""));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, null));
    }

    @Test
    void passesSinFiltroDejaPasarTodo() {
        assertTrue(ShippingType.passes(ShippingType.OTRO, Set.of()));
        assertTrue(ShippingType.passes(ShippingType.FLEX, Set.of()));
    }

    @Test
    void passesConFiltroSoloDejaPasarLosTildados() {
        Set<ShippingType> checked = Set.of(ShippingType.FLEX, ShippingType.TURBO);
        assertTrue(ShippingType.passes(ShippingType.FLEX, checked));
        assertTrue(ShippingType.passes(ShippingType.TURBO, checked));
        assertFalse(ShippingType.passes(ShippingType.COLECTA, checked));
        assertFalse(ShippingType.passes(ShippingType.OTRO, checked));
    }
}
```

- [ ] **Step 2: Correr el test para verificar que falla**

Run: `mvn -q -Dtest=ShippingTypeTest test`
Expected: FAIL (compilación: `ShippingType` no existe).

- [ ] **Step 3: Implementar `ShippingType`**

Create `src/main/java/ar/com/leo/api/ml/model/ShippingType.java`:

```java
package ar.com.leo.api.ml.model;

import java.util.Set;

/** Tipo de envío de una orden ML, para clasificación y filtrado en la UI. */
public enum ShippingType {
    FLEX, COLECTA, TURBO, OTRO;

    /** Clasifica según el tag turbo y el logistic_type del shipment (mapeo estricto). */
    public static ShippingType from(boolean turbo, String logisticType) {
        if (turbo) return TURBO;
        if ("self_service".equals(logisticType)) return FLEX;
        if ("cross_docking".equals(logisticType)) return COLECTA;
        return OTRO;
    }

    /** Sin tipos tildados ⇒ pasa todo; con tildados ⇒ solo esos. */
    public static boolean passes(ShippingType type, Set<ShippingType> checked) {
        return checked.isEmpty() || checked.contains(type);
    }
}
```

- [ ] **Step 4: Correr el test para verificar que pasa**

Run: `mvn -q -Dtest=ShippingTypeTest test`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/main/java/ar/com/leo/api/ml/model/ShippingType.java src/test/java/ar/com/leo/api/ml/model/ShippingTypeTest.java
git commit -m "feat(etiquetas): enum ShippingType para clasificacion y filtro"
```

---

### Task 7: Agregar `logisticType` a `SlaInfo`

**Files:**
- Modify: `src/main/java/ar/com/leo/api/ml/MercadoLibreAPI.java:50` (record `SlaInfo`), `:1100` (construcción en `obtenerSla`)

**Interfaces:**
- Consumes: JSON del shipment ya leído en `obtenerSla` (variable `root`).
- Produces: `SlaInfo` con componente extra `String logisticType` (`sla.logisticType()`).

- [ ] **Step 1: Localizar todas las construcciones de `SlaInfo`**

Run: `grep -rn "new SlaInfo(" src/main/java`
Expected: una sola coincidencia, en `obtenerSla` (~línea 1100). Si hubiera más, actualizar todas en el Step 3.

- [ ] **Step 2: Extender el record**

En `MercadoLibreAPI.java` línea 50:

```java
    public record SlaInfo(String status, OffsetDateTime expectedDate, boolean turbo, String logisticType) {}
```

- [ ] **Step 3: Leer `logistic_type` y pasarlo al constructor**

En `obtenerSla`, dentro del bloque que parsea el shipment (`root`), después de leer los tags para turbo, leer el logistic_type; y actualizar el `return`:

```java
        boolean turbo = false;
        String logisticType = "";
        String shipUrl = "https://api.mercadolibre.com/shipments/" + shipmentId;
        // ... (request y envío igual)
        if (shipResponse != null && shipResponse.statusCode() == 200) {
            try {
                JsonNode root = mapper.readTree(shipResponse.body());
                JsonNode tags = root.path("tags");
                if (tags.isArray()) {
                    for (JsonNode tag : tags) {
                        if ("turbo".equals(tag.asString())) {
                            turbo = true;
                            break;
                        }
                    }
                }
                logisticType = root.path("logistic_type").asString("");
            } catch (Exception e) {
                AppLogger.warn("ML - Error al leer tags de shipment " + shipmentId + ": " + e.getMessage());
            }
        }

        return new SlaInfo(status, expectedDate, turbo, logisticType);
```

- [ ] **Step 4: Compilar toda la suite**

Run: `mvn -q test`
Expected: BUILD SUCCESS (los tests existentes siguen pasando).

- [ ] **Step 5: Commit**

```bash
git add src/main/java/ar/com/leo/api/ml/MercadoLibreAPI.java
git commit -m "feat(etiquetas): leer logistic_type del shipment en SlaInfo"
```

---

### Task 8: `OrderTableRow` lleva `ShippingType` y `displayOrders` lo asigna

**Files:**
- Modify: `src/main/java/ar/com/leo/etiquetas/ui/OrderTableRow.java`
- Modify: `src/main/java/ar/com/leo/ui/MainController.java` (`displayOrders`, armado de `OrderTableRow` ~1395-1425)

**Interfaces:**
- Consumes: `ShippingType.from(boolean, String)` de Task 6; `SlaInfo.turbo()`, `SlaInfo.logisticType()` de Task 7; `slaMap` (parámetro de `displayOrders`).
- Produces: `OrderTableRow.getShippingType() -> ShippingType`.

- [ ] **Step 1: Agregar `shippingType` a `OrderTableRow`**

En `OrderTableRow.java`, agregar campo plano (no property, solo para filtrar) como **último** parámetro del constructor:

```java
import ar.com.leo.api.ml.model.ShippingType;
// ...
    private final ShippingType shippingType;

    public OrderTableRow(boolean selected, String orderId, String zone, String sku, String productDescription,
                         String quantity, String status, String slaDate, List<OrdenML> ordenes,
                         ShippingType shippingType) {
        // ... asignaciones existentes ...
        this.shippingType = shippingType;
    }

    public ShippingType getShippingType() { return shippingType; }
```

- [ ] **Step 2: Calcular el tipo en `displayOrders` y pasarlo al constructor**

En `MainController.displayOrders`, en el bloque que ya detecta turbo (`esTurbo`, ~línea 1395), capturar el `logisticType` del primer shipment con SLA del grupo, y construir el `ShippingType`:

```java
            // Detectar turbo y logistic_type del envío
            boolean esTurbo = false;
            String logisticType = "";
            for (OrdenML o : group) {
                Long shipId = o.getShipmentId();
                if (shipId != null && slaMap.containsKey(shipId)) {
                    MercadoLibreAPI.SlaInfo sla = slaMap.get(shipId);
                    if (sla.turbo()) esTurbo = true;
                    if (logisticType.isEmpty() && sla.logisticType() != null && !sla.logisticType().isEmpty()) {
                        logisticType = sla.logisticType();
                    }
                }
            }
            ar.com.leo.api.ml.model.ShippingType shippingType =
                    ar.com.leo.api.ml.model.ShippingType.from(esTurbo, logisticType);
```

Y actualizar la llamada `rows.add(new OrderTableRow(...))` (~línea 1424) para pasar `shippingType` como último argumento:

```java
            rows.add(new OrderTableRow(true, orderIdStr, zone, skuJoiner.toString(),
                    descJoiner.toString(), qtyJoiner.toString(), status, slaDate, group, shippingType));
```

(Nota: reemplaza el bucle `esTurbo` anterior; no dupliques la detección de turbo.)

- [ ] **Step 3: Compilar**

Run: `mvn -q -DskipTests compile`
Expected: BUILD SUCCESS.

- [ ] **Step 4: Commit**

```bash
git add src/main/java/ar/com/leo/etiquetas/ui/OrderTableRow.java src/main/java/ar/com/leo/ui/MainController.java
git commit -m "feat(etiquetas): OrderTableRow con ShippingType calculado en displayOrders"
```

---

### Task 9: Checkboxes de filtro + `applyOrderFilters`

**Files:**
- Modify: `src/main/resources/ar/com/leo/ui/MainView.fxml:107-118` (sub-tab API)
- Modify: `src/main/java/ar/com/leo/ui/MainController.java` (declaraciones ~90-92, `initialize()` listener del `searchField` ~392-408, `setLoading` ~1296)

**Interfaces:**
- Consumes: `ShippingType.passes(ShippingType, Set<ShippingType>)` de Task 6; `OrderTableRow.getShippingType()` de Task 8; `filteredOrders`, `filteredLabels` (existentes).
- Produces: `applyOrderFilters()` que combina buscador + checkboxes de tipo de envío.

- [ ] **Step 1: Agregar los checkboxes al FXML**

En `MainView.fxml`, dentro del `VBox` del tab "API MercadoLibre", **después** del `HBox` que contiene `fetchOrdersBtn`/combos (después de la línea ~117), agregar:

```xml
                                    <HBox alignment="CENTER_LEFT" spacing="15">
                                        <Label text="Tipo de envío:" />
                                        <CheckBox fx:id="filterFlexCheck" text="📦 Flex" />
                                        <CheckBox fx:id="filterColectaCheck" text="🚚 Colecta" />
                                        <CheckBox fx:id="filterTurboCheck" text="⚡ Turbo" />
                                    </HBox>
```

- [ ] **Step 2: Declarar los checkboxes en `MainController`**

Junto a `estadoFilterCombo`/`despachoFilterCombo` (~líneas 90-92):

```java
    @FXML
    private CheckBox filterFlexCheck;
    @FXML
    private CheckBox filterColectaCheck;
    @FXML
    private CheckBox filterTurboCheck;
```

- [ ] **Step 3: Crear `applyOrderFilters()` y engancharlo**

Agregar el método (por ejemplo cerca del listener del `searchField`). Reemplaza el seteo directo del predicate del `searchField` sobre `filteredOrders`:

```java
    private void applyOrderFilters() {
        String filter = searchField.getText() == null ? "" : searchField.getText().trim().toLowerCase();

        java.util.Set<ar.com.leo.api.ml.model.ShippingType> checked = new java.util.HashSet<>();
        if (filterFlexCheck.isSelected()) checked.add(ar.com.leo.api.ml.model.ShippingType.FLEX);
        if (filterColectaCheck.isSelected()) checked.add(ar.com.leo.api.ml.model.ShippingType.COLECTA);
        if (filterTurboCheck.isSelected()) checked.add(ar.com.leo.api.ml.model.ShippingType.TURBO);

        if (filteredOrders != null) {
            filteredOrders.setPredicate(row -> {
                boolean matchesSearch = filter.isEmpty()
                        || row.getOrderId().toLowerCase().contains(filter)
                        || (row.getSku() != null && row.getSku().toLowerCase().contains(filter))
                        || (row.getZone() != null && row.getZone().toLowerCase().contains(filter))
                        || (row.getProductDescription() != null && row.getProductDescription().toLowerCase().contains(filter));
                boolean matchesType = ar.com.leo.api.ml.model.ShippingType.passes(row.getShippingType(), checked);
                return matchesSearch && matchesType;
            });
        }
    }
```

En el listener existente de `searchField.textProperty()` (~línea 392), reemplazar el bloque `if (filteredOrders != null) { filteredOrders.setPredicate(...); }` por una llamada a `applyOrderFilters();` (el bloque de `filteredLabels` para la tabla de etiquetas queda **sin cambios**):

```java
        searchField.textProperty().addListener((obs, oldVal, newVal) -> {
            String filter = newVal == null ? "" : newVal.trim().toLowerCase();
            applyOrderFilters();
            if (filteredLabels != null) {
                filteredLabels.setPredicate(row ->
                        filter.isEmpty() || (row.getOrderIds() != null && row.getOrderIds().toLowerCase().contains(filter))
                                || (row.getSku() != null && row.getSku().toLowerCase().contains(filter))
                                || (row.getZone() != null && row.getZone().toLowerCase().contains(filter))
                                || (row.getProductDescription() != null && row.getProductDescription().toLowerCase().contains(filter)));
            }
        });
```

Agregar listeners de los checkboxes en `initialize()` (por ejemplo junto a la config de los combos, ~línea 410):

```java
        filterFlexCheck.selectedProperty().addListener((o, a, b) -> applyOrderFilters());
        filterColectaCheck.selectedProperty().addListener((o, a, b) -> applyOrderFilters());
        filterTurboCheck.selectedProperty().addListener((o, a, b) -> applyOrderFilters());
```

- [ ] **Step 4: Deshabilitar los checkboxes durante loading**

En `setLoading` (~línea 1296, junto a `estadoFilterCombo.setDisable(loading)`):

```java
        filterFlexCheck.setDisable(loading);
        filterColectaCheck.setDisable(loading);
        filterTurboCheck.setDisable(loading);
```

- [ ] **Step 5: Compilar toda la suite**

Run: `mvn -q test`
Expected: BUILD SUCCESS, todos los tests pasan.

- [ ] **Step 6: Verificar en la app (manual)**

Run: `mvn javafx:run`
En el sub-tab API, tras "Obtener Órdenes": (a) sin tildar nada → aparecen todas las órdenes; (b) tildar solo Turbo → solo turbo; (c) tildar Flex+Colecta → ambos, ocultando turbo y otros; (d) el buscador combina con los checkboxes. Cerrar la app.

- [ ] **Step 7: Commit**

```bash
git add src/main/resources/ar/com/leo/ui/MainView.fxml src/main/java/ar/com/leo/ui/MainController.java
git commit -m "feat(etiquetas): filtros por tipo de envio (Flex/Colecta/Turbo)"
```

---

## Self-Review

**Spec coverage:**
- Cambio 0 (rename Despacho) → Task 1. ✅
- Cambio 1 (columna "#") → Task 5. ✅
- Cambio 2 (orden carros: tabla etiquetas + impresión + archivo) → Task 3 (LabelSorter cubre labelTable + impresión + archivo). Cambio 2 (tabla órdenes) → Task 4. ✅
- Cambio 3 (filtros): enum → Task 6; `logistic_type` en SlaInfo → Task 7; OrderTableRow + displayOrders → Task 8; checkboxes + filtro → Task 9. ✅
- Estrategia JUnit para lógica pura → Task 2 (setup) + tests en Tasks 2, 3, 5, 6. ✅

**Type consistency:** `CarrosOrdering.bucket(String,String)`, `PrintNumbering.compute(List<Integer>)`, `ShippingType.from(boolean,String)` / `passes(ShippingType,Set)`, `SlaInfo(...,String logisticType)`, `OrderTableRow(...,ShippingType)`, `LabelTableRow(String printNumber, ...)` — usados consistentes entre tasks. ✅

**Placeholders:** ninguno; todo paso tiene código real. ✅
