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
