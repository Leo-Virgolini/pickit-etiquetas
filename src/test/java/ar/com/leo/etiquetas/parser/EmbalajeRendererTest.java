package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class EmbalajeRendererTest {

    private static DatosEmbalaje datos(String nroBolsa, String nombreCaja, String nroCaja,
                                       String rollo, String cantPanos, String observaciones) {
        return new DatosEmbalaje(nroBolsa, nombreCaja, nroCaja, rollo, cantPanos, observaciones);
    }

    // -------------------------------------------------------------------------------------------
    // Caja y bolsa
    // -------------------------------------------------------------------------------------------

    @Test
    void cajaConNumeroYNombre() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "GRANDE", "3", "", "", ""));

        assertEquals(List.of("CAJA: 3 - GRANDE"), lineas);
    }

    @Test
    void cajaSoloConNumero() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "3", "", "", ""));

        assertEquals(List.of("CAJA: 3"), lineas);
    }

    @Test
    void cajaSoloConNombre() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "GRANDE", "", "", "", ""));

        assertEquals(List.of("CAJA: GRANDE"), lineas);
    }

    @Test
    void bolsaConNumero() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("5", "", "", "", "", ""));

        assertEquals(List.of("BOLSA: 5"), lineas);
    }

    // -------------------------------------------------------------------------------------------
    // Agregados
    // -------------------------------------------------------------------------------------------

    @Test
    void rolloConPanos() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "DIAMANTE", "3", ""));

        assertEquals(List.of("NO ESTANDARIZADO", "ROLLO: DIAMANTE - 3 paños"), lineas);
    }

    @Test
    void rolloSinPanos() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "CUADRADO", "", ""));

        assertEquals(List.of("NO ESTANDARIZADO", "ROLLO: CUADRADO"), lineas);
    }

    @Test
    void observaciones() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "", "", "Colchon + Tapa"));

        assertEquals(List.of("NO ESTANDARIZADO", "OBS: Colchon + Tapa"), lineas);
    }

    // -------------------------------------------------------------------------------------------
    // Combinaciones
    // -------------------------------------------------------------------------------------------

    @Test
    void elOrdenEsCajaPluribolRolloObservaciones() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "GRANDE", "3", "DIAMANTE", "3", "Colchon + Tapa"));

        assertEquals(List.of(
                "CAJA: 3 - GRANDE",
                "ROLLO: DIAMANTE - 3 paños",
                "OBS: Colchon + Tapa"), lineas);
    }

    @Test
    void conCajaYBolsaCargadasGanaLaCaja() {
        // Puede pasar cuando cambia el embalaje y queda el número de bolsa viejo. Sin esto salen
        // 5 líneas y la última se imprime sobre el separador y el título del producto.
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("5", "GRANDE", "3", "DIAMANTE", "3", "Colchon"));

        assertEquals(3, lineas.size(), lineas.toString());
        assertEquals("CAJA: 3 - GRANDE", lineas.get(0));
    }

    @Test
    void losSaltosDeLineaDelExcelSeColapsan() {
        // OBSERVACIONES cargado con Alt+Enter: el LF crudo dentro del ^FD pega las palabras.
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "", "", "Colchon\n+ Tapa"));

        assertEquals(List.of("NO ESTANDARIZADO", "OBS: Colchon + Tapa"), lineas);
    }

    @Test
    void sinCajaNiBolsaSeAvisaEnLaEtiqueta() {
        assertEquals(List.of("NO ESTANDARIZADO"), EmbalajeRenderer.lineas(DatosEmbalaje.VACIO));
    }

    @Test
    void sinCajaNiBolsaLosAgregadosSeSiguenMostrando() {
        List<String> lineas = EmbalajeRenderer.lineas(
                datos("", "", "", "DIAMANTE", "3", "Colchon"));

        assertEquals(List.of(
                "NO ESTANDARIZADO",
                "ROLLO: DIAMANTE - 3 paños",
                "OBS: Colchon"), lineas);
    }

    @Test
    void losEspaciosSobrantesNoCuentanComoDato() {
        assertEquals(List.of("NO ESTANDARIZADO"), EmbalajeRenderer.lineas(
                datos("  ", "  ", "  ", "  ", "  ", "  ")));
    }

    // -------------------------------------------------------------------------------------------
    // Campo ZPL
    // -------------------------------------------------------------------------------------------

    @Test
    void elCampoZplApilaLasLineas() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("CAJA: 3", "BOLSA: 5"));

        // La última línea se queda con el alto libre restante: si el texto no entra en una línea
        // sigue abajo, en vez de que ZPL lo reimprima encima de la misma (^FB con maxLines=1
        // sobreescribe la última línea, no descarta el sobrante).
        assertEquals("^FO410,20^A0N,22,22^FB380,1,0,L^FDCAJA: 3^FS\n"
                        + "^FO410,44^A0N,22,22^FB380,5,0,L^FDBOLSA: 5^FS\n",
                zpl);
    }

    @Test
    void laUltimaLineaSeQuedaConElAltoRestante() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("CAJA: 3", "ROLLO: X", "OBS: larga"));

        assertTrue(zpl.contains("^FO410,68^A0N,22,22^FB380,4,0,L^FDOBS: larga^FS"), zpl);
    }

    @Test
    void elBloqueNoLlegaAlSeparadorDeLaZonaDePicking() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("a", "b", "c", "d", "e", "f"));

        // Seis líneas: la última arranca en y=140 y con fuente 22 termina en 162 (separador: 180).
        assertTrue(zpl.contains("^FO410,140^A0N,22,22^FB380,1,0,L^FDf^FS"), zpl);
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
