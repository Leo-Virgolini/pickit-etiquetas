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
