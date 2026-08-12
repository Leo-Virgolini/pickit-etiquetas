package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class EmbalajeRendererTest {

    private static DatosEmbalaje datos(String nroBolsa, String nombreCaja, String nroCaja,
                                       String pluribol, String cantPluribol,
                                       String rollo, String cantPanos, String observaciones) {
        return new DatosEmbalaje(nroBolsa, nombreCaja, nroCaja, pluribol, cantPluribol,
                rollo, cantPanos, observaciones);
    }

    // -------------------------------------------------------------------------------------------
    // Caja y bolsa
    // -------------------------------------------------------------------------------------------

    @Test
    void cajaConNumeroYNombre() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "GRANDE", "3", "", "", "", "", ""));

        assertEquals(List.of("CAJA: 3 - GRANDE"), lineas);
    }

    @Test
    void cajaSoloConNumero() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "3", "", "", "", "", ""));

        assertEquals(List.of("CAJA: 3"), lineas);
    }

    @Test
    void cajaSoloConNombre() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "GRANDE", "", "", "", "", "", ""));

        assertEquals(List.of("CAJA: GRANDE"), lineas);
    }

    @Test
    void bolsaConNumero() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("5", "", "", "", "", "", "", ""));

        assertEquals(List.of("BOLSA: 5"), lineas);
    }

    // -------------------------------------------------------------------------------------------
    // Agregados
    // -------------------------------------------------------------------------------------------

    @Test
    void pluribolConVueltas() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "SI", "2", "", "", ""));

        assertEquals(List.of("PLURIBOL: SI - 2 vueltas"), lineas);
    }

    @Test
    void pluribolSinVueltas() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "SI", "", "", "", ""));

        assertEquals(List.of("PLURIBOL: SI"), lineas);
    }

    @Test
    void rolloConPanos() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "", "", "DIAMANTE", "3", ""));

        assertEquals(List.of("ROLLO: DIAMANTE - 3 paños"), lineas);
    }

    @Test
    void rolloSinPanos() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "", "", "CUADRADO", "", ""));

        assertEquals(List.of("ROLLO: CUADRADO"), lineas);
    }

    @Test
    void observaciones() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "", "", "", "", "Colchon + Tapa"));

        assertEquals(List.of("OBS: Colchon + Tapa"), lineas);
    }

    // -------------------------------------------------------------------------------------------
    // Combinaciones
    // -------------------------------------------------------------------------------------------

    @Test
    void elOrdenEsCajaPluribolRolloObservaciones() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "GRANDE", "3", "SI", "2", "DIAMANTE", "3", "Colchon + Tapa"));

        assertEquals(List.of(
                "CAJA: 3 - GRANDE",
                "PLURIBOL: SI - 2 vueltas",
                "ROLLO: DIAMANTE - 3 paños",
                "OBS: Colchon + Tapa"), lineas);
    }

    @Test
    void conCajaYBolsaCargadasGanaLaCaja() {
        // Puede pasar cuando cambia el embalaje y queda el número de bolsa viejo. Sin esto salen
        // 5 líneas y la última se imprime sobre el separador y el título del producto.
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("5", "GRANDE", "3", "SI", "2", "DIAMANTE", "3", "Colchon"));

        assertEquals(4, lineas.size(), lineas.toString());
        assertEquals("CAJA: 3 - GRANDE", lineas.get(0));
    }

    @Test
    void losSaltosDeLineaDelExcelSeColapsan() {
        // OBSERVACIONES cargado con Alt+Enter: el LF crudo dentro del ^FD pega las palabras.
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "", "", "", "", "Colchon\n+ Tapa"));

        assertEquals(List.of("OBS: Colchon + Tapa"), lineas);
    }

    @Test
    void sinDatosNoHayLineas() {
        assertEquals(List.of(), EmbalajeRenderer.lineas(DatosEmbalaje.VACIO));
    }

    @Test
    void losEspaciosSobrantesNoCuentanComoDato() {
        assertEquals(List.of(), EmbalajeRenderer.lineas(
                datos("  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ")));
    }

    // -------------------------------------------------------------------------------------------
    // Campo ZPL
    // -------------------------------------------------------------------------------------------

    @Test
    void elCampoZplApilaLasLineas() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("CAJA: 3", "BOLSA: 5"));

        assertEquals("^FO450,83^A0N,22,22^FB340,1,0,L^FDCAJA: 3^FS\n"
                        + "^FO450,107^A0N,22,22^FB340,1,0,L^FDBOLSA: 5^FS\n",
                zpl);
    }

    @Test
    void cuatroLineasNoLleganAlSeparadorDeLaZonaDePicking() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("a", "b", "c", "d"));

        // La última arranca en y=155 y con fuente 22 termina en 177: el separador está en 180.
        assertTrue(zpl.contains("^FO450,155^A0N,22,22"), zpl);
    }

    @Test
    void elCampoZplNeutralizaLosCaracteresDeControlDeZpl() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("CAJA: ^3~A"));

        assertTrue(zpl.contains("^FDCAJA:  3 A^FS"), zpl);
    }

    @Test
    void sinLineasNoHayCampo() {
        assertEquals("", EmbalajeRenderer.campoZpl(List.of()));
    }
}
