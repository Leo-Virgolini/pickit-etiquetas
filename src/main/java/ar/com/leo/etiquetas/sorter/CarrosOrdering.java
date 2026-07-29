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
