package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.ZplLabel;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ZplParserTest {

    private final ZplParser parser = new ZplParser();

    /**
     * Encabezado tal como lo imprime ML, con el rótulo y el número en campos separados y un
     * "20000" en el medio.
     */
    private static String etiqueta(String rotulo, String numero) {
        return "^XA\n"
                + "^CI28\n"
                + "^LH0,90\n"
                + "^FO30,40^A0N,28,28^FH^FD" + rotulo + ":^FS\n"
                + "^FO31,40^A0N,28,28^FH^FD" + rotulo + ":^FS\n"
                + "^FO134,39^A0N,25,25^FD20000^FS\n"
                + "^FO198,40^A0N,30,30^FD" + numero + "^FS\n"
                + "^FO10,130^A0N,70,70^FB160,1,0,C^FD1^FS\n"
                + "^FO200,181^A0N,24,24^FB570,3,-1^FH^FDColor: Gris  | SKU: 1241212^FS\n"
                + "^FO120,135^A0N,24,24^FD" + rotulo + ": 20000^FS\n"
                + "^FO272,132^A0N,27,27^FD" + numero + "^FS\n"
                + "^XZ\n";
    }

    @Test
    void tomaElNumeroDelPackId() {
        List<ZplLabel> labels = parser.parse(etiqueta("Pack ID", "13605621573"));

        assertEquals("13605621573", labels.getFirst().orderIds());
    }

    @Test
    void tomaElNumeroDeLaVentaCuandoNoHayPaquete() {
        // Las etiquetas que no agrupan varias órdenes traen "Venta ID:" en vez de "Pack ID:".
        List<ZplLabel> labels = parser.parse(etiqueta("Venta ID", "17019770092"));

        assertEquals("17019770092", labels.getFirst().orderIds());
    }

    @Test
    void noSeConfundeConElNumeroCortoQueVaEnElMedio() {
        // Entre el rótulo y el número, ML imprime un "20000" en su propio campo.
        List<ZplLabel> labels = parser.parse(etiqueta("Pack ID", "13605621573"));

        assertEquals("13605621573", labels.getFirst().orderIds());
    }

    @Test
    void sinRotuloConocidoLaOrdenQuedaVacia() {
        String sinId = "^XA\n"
                + "^FO200,181^A0N,24,24^FH^FDColor: Gris  | SKU: 1241212^FS\n"
                + "^XZ\n";

        assertEquals("", parser.parse(sinId).getFirst().orderIds());
    }

    @Test
    void elRestoDeLosDatosSigueSaliendoIgual() {
        List<ZplLabel> labels = parser.parse(etiqueta("Pack ID", "13605621573"));

        assertEquals(1, labels.size());
        assertEquals("1241212", labels.getFirst().sku());
        assertEquals(1, labels.getFirst().quantity());
    }
}
