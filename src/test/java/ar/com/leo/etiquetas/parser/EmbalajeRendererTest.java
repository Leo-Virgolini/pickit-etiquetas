package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class EmbalajeRendererTest {

    private static DatosEmbalaje datos(String envase, String inscripcion, String rollo,
                                       String cantPanos, String observaciones) {
        return new DatosEmbalaje(envase, inscripcion, rollo, cantPanos, observaciones, true);
    }

    // -------------------------------------------------------------------------------------------
    // Envase
    // -------------------------------------------------------------------------------------------

    @Test
    void elEnvaseLlevaSuCodigoYSuInscripcion() {
        assertEquals(List.of("ENVASE: BOL-1 - 9Y"),
                EmbalajeRenderer.lineas(datos("BOL-1", "9Y", "", "", "")));
    }

    @Test
    void sinInscripcionConocidaVaSoloElCodigo() {
        assertEquals(List.of("ENVASE: CAJ-9"),
                EmbalajeRenderer.lineas(datos("CAJ-9", "", "", "", "")));
    }

    @Test
    void unaInscripcionQueEsSoloUnGuionSeIgnora() {
        // La fila "NO" de la hoja de estandarización tiene "-" en las columnas que no aplican.
        assertEquals(List.of("ENVASE: NO"),
                EmbalajeRenderer.lineas(datos("NO", "-", "", "", "")));
    }

    @Test
    void unSoloPanoVaEnSingular() {
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: DIAMANTES - 1 paño"),
                EmbalajeRenderer.lineas(datos("BOL-1", "", "DIAMANTES", "1", "")));
    }

    @Test
    void unProductoSinEnvaseLoDiceExplicitamente() {
        // "NO" es una decisión tomada, no un dato faltante: se imprime como cualquier otro valor.
        assertEquals(List.of("ENVASE: NO"),
                EmbalajeRenderer.lineas(datos("NO", "", "", "", "")));
    }

    // -------------------------------------------------------------------------------------------
    // Rollo y paños
    // -------------------------------------------------------------------------------------------

    @Test
    void elRolloConPanos() {
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: DIAMANTES - 2 paños"),
                EmbalajeRenderer.lineas(datos("BOL-1", "", "DIAMANTES", "2", "")));
    }

    @Test
    void losPanosEnCeroNoSeMuestran() {
        // La columna trae 0 en las filas sin rollo: "ROLLO: NO - 0 paños" sería ruido.
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: DIAMANTES"),
                EmbalajeRenderer.lineas(datos("BOL-1", "", "DIAMANTES", "0", "")));
    }

    @Test
    void unRolloEnNoSeImprimeIgual() {
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: NO"),
                EmbalajeRenderer.lineas(datos("BOL-1", "", "NO", "0", "")));
    }

    @Test
    void unaCantidadDePanosNoNumericaSeIgnora() {
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: DIAMANTES"),
                EmbalajeRenderer.lineas(datos("BOL-1", "", "DIAMANTES", "dos", "")));
    }

    // -------------------------------------------------------------------------------------------
    // Combinaciones
    // -------------------------------------------------------------------------------------------

    @Test
    void elOrdenEsEnvaseRolloObservaciones() {
        assertEquals(List.of(
                        "ENVASE: BOL-1 - 9Y",
                        "ROLLO: DIAMANTES - 2 paños",
                        "OBS: Colchon + Tapa"),
                EmbalajeRenderer.lineas(datos("BOL-1", "9Y", "DIAMANTES", "2", "Colchon + Tapa")));
    }

    @Test
    void losSaltosDeLineaDelExcelSeColapsan() {
        // OBSERVACIONES cargado con Alt+Enter: el LF crudo dentro del ^FD pega las palabras.
        assertEquals(List.of("ENVASE: BOL-1", "OBS: Colchon + Tapa"),
                EmbalajeRenderer.lineas(datos("BOL-1", "", "", "", "Colchon\n+ Tapa")));
    }

    // -------------------------------------------------------------------------------------------
    // Sin estandarizar
    // -------------------------------------------------------------------------------------------

    @Test
    void sinEstandarizarSoloSeAvisa() {
        DatosEmbalaje sinEstandarizar = new DatosEmbalaje("BOL-1", "9Y", "DIAMANTES", "2", "Obs", false);

        assertEquals(List.of("NO ESTANDARIZADO"), EmbalajeRenderer.lineas(sinEstandarizar));
    }

    @Test
    void datosVaciosNoEstanEstandarizados() {
        assertEquals(List.of("NO ESTANDARIZADO"), EmbalajeRenderer.lineas(DatosEmbalaje.VACIO));
    }

    // -------------------------------------------------------------------------------------------
    // Campo ZPL
    // -------------------------------------------------------------------------------------------

    @Test
    void elCampoZplApilaLasLineas() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("ENVASE: BOL-1", "OBS: algo"));

        assertEquals("^FO410,20^A0N,22,22^FB380,1,0,L^FDENVASE: BOL-1^FS\n"
                        + "^FO411,20^A0N,22,22^FB380,1,0,L^FDENVASE:^FS\n"
                        + "^FO410,44^A0N,22,22^FB380,5,0,L^FDOBS: algo^FS\n"
                        + "^FO411,44^A0N,22,22^FB380,5,0,L^FDOBS:^FS\n",
                zpl);
    }

    @Test
    void elAvisoDeNoEstandarizadoVaMasGrandeYEnNegrita() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("NO ESTANDARIZADO"));

        // Doble pasada con 1px de offset: es como se simula la negrita en ZPL.
        assertTrue(zpl.contains("^FO410,20^A0N,30,30"), zpl);
        assertTrue(zpl.contains("^FO411,20^A0N,30,30"), zpl);
    }

    @Test
    void unaLineaSinRotuloNoLlevaSegundaPasada() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("SIN DOS PUNTOS"));

        assertEquals(1, zpl.lines().count(), zpl);
    }

    @Test
    void elBloqueNoLlegaAlSeparadorDeLaZonaDePicking() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("a", "b", "c", "d", "e", "f"));

        // Seis líneas: la última arranca en y=140 y con fuente 22 termina en 162 (separador: 180).
        assertTrue(zpl.contains("^FO410,140^A0N,22,22^FB380,1,0,L^FDf^FS"), zpl);
    }

    @Test
    void elCampoZplNeutralizaLosCaracteresDeControlDeZpl() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("ENVASE: ^3~A"));

        assertTrue(zpl.contains("^FDENVASE:  3 A^FS"), zpl);
    }

    @Test
    void unaLineaQueNoEsLaUltimaSeTruncaEnVezDeSuperponerse() {
        String largo = "ENVASE: CAJ-1 - INSCRIPCION MUY LARGA QUE NO ENTRA";
        String zpl = EmbalajeRenderer.campoZpl(List.of(largo, "OBS: algo"));

        assertTrue(zpl.contains("^FDENVASE: CAJ-1 - INSCRIPCION...^FS"), zpl);
    }

    @Test
    void laUltimaLineaSeAcotaALoQueEntraEnSuAlto() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("ENVASE: X", "ROLLO: X", "OBS: " + "x".repeat(300)));

        String impreso = zpl.substring(zpl.indexOf("^FDOBS:") + 3, zpl.indexOf("^FS", zpl.indexOf("^FDOBS:")));
        assertTrue(impreso.replace(" ", "").length() <= 4 * 30, impreso);
        assertTrue(impreso.endsWith("..."), impreso);
    }

    @Test
    void unaPalabraMasLargaQueLaLineaSeParteParaQuePuedaEnvolver() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("ENVASE: X", "OBS: " + "a".repeat(50)));

        assertTrue(zpl.contains("^FDOBS: " + "a".repeat(25) + " " + "a".repeat(25) + "^FS"), zpl);
    }

    @Test
    void sinLineasNoHayCampo() {
        assertEquals("", EmbalajeRenderer.campoZpl(List.of()));
    }
}
