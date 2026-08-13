package ar.com.leo.etiquetas.sorter;

import ar.com.leo.etiquetas.model.SortResult;
import ar.com.leo.etiquetas.model.ZplLabel;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * La tabla agrupa las etiquetas del mismo SKU en una fila, que es lo que le sirve al que va a
 * buscar los productos. La columna Orden tiene que listar todas las ventas de ese grupo: quedarse
 * con la primera deja al operario sin poder encontrar las otras.
 */
class LabelSorterOrdenesTest {

    private static ZplLabel etiqueta(String sku, String orden) {
        return new ZplLabel("^XA^XZ", sku, "Producto", "det", 1, false, orden);
    }

    @Test
    void unGrupoListaTodasSusOrdenes() {
        SortResult result = new LabelSorter().sort(List.of(
                etiqueta("1241214", "15258696888"),
                etiqueta("1241214", "15262795618"),
                etiqueta("1241214", "11698956321")), Map.of());

        assertEquals("15258696888\n15262795618\n11698956321",
                result.groups().getFirst().orderIds());
    }

    @Test
    void unaOrdenRepetidaSeListaUnaSolaVez() {
        // Dos unidades del mismo SKU en la misma venta salen en dos etiquetas.
        SortResult result = new LabelSorter().sort(List.of(
                etiqueta("1241214", "15258696888"),
                etiqueta("1241214", "15258696888")), Map.of());

        assertEquals("15258696888", result.groups().getFirst().orderIds());
    }

    @Test
    void lasEtiquetasSinOrdenNoDejanRenglonesVacios() {
        SortResult result = new LabelSorter().sort(List.of(
                etiqueta("1241214", ""),
                etiqueta("1241214", "15258696888")), Map.of());

        assertEquals("15258696888", result.groups().getFirst().orderIds());
    }

    @Test
    void sinNingunaOrdenLaColumnaQuedaVacia() {
        SortResult result = new LabelSorter().sort(List.of(etiqueta("1241214", "")), Map.of());

        assertEquals("", result.groups().getFirst().orderIds());
    }
}
