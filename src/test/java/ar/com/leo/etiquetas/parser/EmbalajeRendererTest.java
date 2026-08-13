package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class EmbalajeRendererTest {

    private static DatosEmbalaje datos(String envase, String inscripcion, String rollo,
                                       String cantPanos, String observaciones) {
        return new DatosEmbalaje(envase, inscripcion, rollo, cantPanos, observaciones, true, true);
    }

    /** Etiqueta de un producto suelto, que es el caso base. */
    private static List<String> lineas(DatosEmbalaje datos) {
        return EmbalajeRenderer.lineas(datos, 1);
    }

    // -------------------------------------------------------------------------------------------
    // Envase
    // -------------------------------------------------------------------------------------------

    @Test
    void elEnvaseLlevaSuCodigoYSuInscripcion() {
        assertEquals(List.of("ENVASE: BOL-1 \"9Y\""),
                lineas(datos("BOL-1", "9Y", "", "", "")));
    }

    @Test
    void sinInscripcionConocidaVaSoloElCodigo() {
        // Sin comillas vacías colgando del código.
        assertEquals(List.of("ENVASE: CAJ-9"),
                lineas(datos("CAJ-9", "", "", "", "")));
    }

    @Test
    void unaInscripcionQueEsSoloUnGuionSeIgnora() {
        // La fila "NO" de la hoja de estandarización tiene "-" en las columnas que no aplican.
        assertEquals(List.of("ENVASE: NO"),
                lineas(datos("NO", "-", "", "", "")));
    }

    @Test
    void unSoloPanoVaEnSingular() {
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: DIAMANTES - 1 paño"),
                lineas(datos("BOL-1", "", "DIAMANTES", "1", "")));
    }

    @Test
    void unProductoSinEnvaseLoDiceExplicitamente() {
        // "NO" es una decisión tomada, no un dato faltante: se imprime como cualquier otro valor.
        assertEquals(List.of("ENVASE: NO"),
                lineas(datos("NO", "", "", "", "")));
    }

    // -------------------------------------------------------------------------------------------
    // Rollo y paños
    // -------------------------------------------------------------------------------------------

    @Test
    void elRolloConPanos() {
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: DIAMANTES - 2 paños"),
                lineas(datos("BOL-1", "", "DIAMANTES", "2", "")));
    }

    @Test
    void losPanosEnCeroNoSeMuestran() {
        // La columna trae 0 en las filas sin rollo: "ROLLO: NO - 0 paños" sería ruido.
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: DIAMANTES"),
                lineas(datos("BOL-1", "", "DIAMANTES", "0", "")));
    }

    @Test
    void unRolloEnNoSeImprimeIgual() {
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: NO"),
                lineas(datos("BOL-1", "", "NO", "0", "")));
    }

    @Test
    void unaCantidadCombinadaSeImprimeTalCual() {
        // "2 Y 1" es el valor que aparece en el Excel real.
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: MIXTO - 2 Y 1 paños"),
                lineas(datos("BOL-1", "", "MIXTO", "2 Y 1", "")));
    }

    @Test
    void noSeRepiteLaPalabraPanosSiYaEstaEnElTexto() {
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: MIXTO - 2 PAÑOS GRANDES"),
                lineas(datos("BOL-1", "", "MIXTO", "2 PAÑOS GRANDES", "")));
    }

    @Test
    void unaCantidadDePanosNoNumericaSeImprimeTalCual() {
        // CANT PAÑOS es texto libre del usuario: si escribió "2-3", esconderlo sería peor que
        // mostrarlo, porque el dato existe y el operario lo necesita.
        assertEquals(List.of("ENVASE: BOL-1", "ROLLO: DIAMANTES - 2-3 paños"),
                lineas(datos("BOL-1", "", "DIAMANTES", "2-3", "")));
    }

    // -------------------------------------------------------------------------------------------
    // Combinaciones
    // -------------------------------------------------------------------------------------------

    @Test
    void elOrdenEsEnvaseRolloObservaciones() {
        assertEquals(List.of(
                        "ENVASE: BOL-1 \"9Y\"",
                        "ROLLO: DIAMANTES - 2 paños",
                        "OBS: Colchon + Tapa"),
                lineas(datos("BOL-1", "9Y", "DIAMANTES", "2", "Colchon + Tapa")));
    }

    @Test
    void losSaltosDeLineaDelExcelSeColapsan() {
        // OBSERVACIONES cargado con Alt+Enter: el LF crudo dentro del ^FD pega las palabras.
        assertEquals(List.of("ENVASE: BOL-1", "OBS: Colchon + Tapa"),
                lineas(datos("BOL-1", "", "", "", "Colchon\n+ Tapa")));
    }

    // -------------------------------------------------------------------------------------------
    // Sin estandarizar
    // -------------------------------------------------------------------------------------------

    @Test
    void sinEstandarizarSoloSeAvisa() {
        DatosEmbalaje sinEstandarizar = new DatosEmbalaje("BOL-1", "9Y", "DIAMANTES", "2", "Obs", false, true);

        assertEquals(List.of("NO ESTANDARIZADO"), lineas(sinEstandarizar));
    }

    @Test
    void unSkuQueNoEstaEnElExcelAvisaQueNoEstaEstandarizado() {
        // Recién se lo agrega al Excel en este mismo lote, así que nadie le cargó el envase todavía.
        assertEquals(List.of("NO ESTANDARIZADO"), lineas(DatosEmbalaje.SIN_CARGAR));
    }

    @Test
    void sinLasColumnasEnElExcelNoSeImprimeNada() {
        // Distinto de "no estandarizado": el archivo no tiene la columna, así que la función no
        // está en uso y la etiqueta no tiene nada que decir.
        assertEquals(List.of(), lineas(DatosEmbalaje.VACIO));
    }

    // -------------------------------------------------------------------------------------------
    // Campo ZPL
    // -------------------------------------------------------------------------------------------

    @Test
    void elCampoZplApilaLasLineas() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("ENVASE: BOL-1", "OBS: algo"));

        assertEquals("^FO400,20^A0N,26,26^FB390,1,0,L^FDENVASE: BOL-1^FS\n"
                        + "^FO401,20^A0N,26,26^FB390,1,0,L^FDENVASE:^FS\n"
                        + "^FO400,48^A0N,26,26^FB390,4,0,L^FDOBS: algo^FS\n"
                        + "^FO401,48^A0N,26,26^FB390,4,0,L^FDOBS:^FS\n",
                zpl);
    }

    @Test
    void elAvisoDeNoEstandarizadoSaleComoRecuadroNegro() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("NO ESTANDARIZADO"));

        // Mismo formato que tenía el banner MEDIR: recuadro relleno y texto en video inverso,
        // centrado, para que frene al operario.
        assertTrue(zpl.contains("^FO400,20^GB390,52,52^FS"), zpl);
        assertTrue(zpl.contains("^FO400,27^A0N,38,38^FB390,1,0,C^FR^FDNO ESTANDARIZADO^FS"), zpl);
    }

    @Test
    void unaLineaSinRotuloNoLlevaSegundaPasada() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("SIN DOS PUNTOS"));

        assertEquals(1, zpl.lines().count(), zpl);
    }

    @Test
    void elBloqueNoLlegaAlSeparadorDeLaZonaDePicking() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("a", "b", "c", "d", "e", "f"));

        // lineas() nunca arma más de cuatro, pero si alguna vez armara de más, lo que no entra
        // se descarta en lugar de imprimirse sobre la zona de picking.
        assertFalse(zpl.contains("^FDf^FS"), zpl);
        assertTrue(zpl.contains("^FO400,132^A0N,26,26^FB390,1,0,L^FDe^FS"), zpl);
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

        assertTrue(zpl.contains("^FDENVASE: CAJ-1 - INSCRIPC...^FS"), zpl);
    }

    @Test
    void laUltimaLineaSeAcotaALoQueEntraEnSuAlto() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("ENVASE: X", "ROLLO: X", "OBS: " + "x".repeat(300)));

        String impreso = zpl.substring(zpl.indexOf("^FDOBS:") + 3, zpl.indexOf("^FS", zpl.indexOf("^FDOBS:")));
        assertTrue(impreso.replace(" ", "").length() <= 3 * 27, impreso);
        assertTrue(impreso.endsWith("..."), impreso);
    }

    @Test
    void laUltimaLineaSeAcotaPorFilasYNoPorCaracteres() {
        // ^FB corta por palabras: estos 54 caracteres entran en dos filas por cuenta de caracteres
        // (2 x 27) pero necesitan tres, y lo que no entra ZPL lo reimprime encima de la última.
        // Es una observación real del Excel.
        String obs = "OBS: 3 COLCHON 1 TAPA. AJUSTAR PAÑO SEGÚN CANT VENDIDA";
        String zpl = EmbalajeRenderer.campoZpl(List.of("REFERENCIA", "ENVASE: X", "ROLLO: X", obs));

        assertTrue(zpl.contains("^FB390,2,0,L^FDOBS: 3 COLCHON 1 TAPA. AJUSTAR PAÑO SEGÚN CANT...^FS"), zpl);
    }

    @Test
    void unaPalabraComunNoSeParteAunqueNoEntreEnLoQueQuedaDeLaFila() {
        // "OBS: 3 COLCHON 1 TAPA." deja 4 caracteres libres y "AJUSTAR" son 7. Partirla en
        // "AJUS TAR" para no desperdiciar ese margen se lee mucho peor que bajarla entera.
        String zpl = EmbalajeRenderer.campoZpl(List.of("OBS: 3 COLCHON 1 TAPA. AJUSTAR PAÑO"));

        assertTrue(zpl.contains("^FDOBS: 3 COLCHON 1 TAPA. AJUSTAR PAÑO^FS"), zpl);
    }

    @Test
    void unaPalabraMasLargaQueLaLineaSeParteParaQuePuedaEnvolver() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("ENVASE: X", "OBS: " + "a".repeat(50)));

        // La primera pieza comparte fila con el rótulo, así que entra lo que sobra de los 27; las
        // siguientes se quedan con la fila entera.
        assertTrue(zpl.contains("^FDOBS: " + "a".repeat(22) + " " + "a".repeat(27) + " a^FS"), zpl);
    }

    @Test
    void sinLineasNoHayCampo() {
        assertEquals("", EmbalajeRenderer.campoZpl(List.of()));
    }

    // -------------------------------------------------------------------------------------------
    // Etiquetas de más de una unidad
    // -------------------------------------------------------------------------------------------

    @Test
    void conMasDeUnaUnidadLosDatosVanComoReferencia() {
        // El operario no está embalando un producto suelto, así que el envase es orientativo.
        assertEquals(List.of("REFERENCIA", "ENVASE: BOL-1 \"9Y\"", "ROLLO: DIAMANTES - 2 paños"),
                EmbalajeRenderer.lineas(datos("BOL-1", "9Y", "DIAMANTES", "2", ""), 2));
    }

    @Test
    void conMasDeUnaUnidadYSinEmbalajeCargadoNoSeImprimeNada() {
        DatosEmbalaje sinEstandarizar = new DatosEmbalaje("BOL-1", "9Y", "DIAMANTES", "2", "", false, true);

        assertEquals(List.of(), EmbalajeRenderer.lineas(sinEstandarizar, 2));
        assertEquals(List.of(), EmbalajeRenderer.lineas(DatosEmbalaje.SIN_CARGAR, 2));
    }

    @Test
    void elEncabezadoDeReferenciaNoSaleSolo() {
        // Las columnas de envase, rollo y observaciones son del usuario y se ubican por encabezado:
        // puede tener ESTANDARIZADO y ninguna de las otras tres. Sin datos abajo, "REFERENCIA" no
        // dice nada.
        DatosEmbalaje soloLaFormula = new DatosEmbalaje("", "", "", "", "", true, true);

        assertEquals(List.of(), EmbalajeRenderer.lineas(soloLaFormula, 2));
    }

    @Test
    void laLineaDeReferenciaVaEnOtraFuenteSubrayadaYEnNegrita() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("REFERENCIA", "ENVASE: BOL-1"));

        // Fuente bitmap B, de trazo más cuadrado que el ^A0 del resto del bloque.
        assertTrue(zpl.contains("^FO400,20^ABN,22,14^FDREFERENCIA^FS"), zpl);
        // Negrita: segunda pasada corrida 1px, como en las demás líneas.
        assertTrue(zpl.contains("^FO401,20^ABN,22,14^FDREFERENCIA^FS"), zpl);
        // Subrayado: ZPL no lo tiene como atributo, así que es una línea dibujada. La fuente es
        // monoespaciada, así que el ancho es exacto: 10 glifos de 14 más los 9 espacios de 4 que
        // la fuente B deja entre caracteres.
        assertTrue(zpl.contains("^FO400,42^GB176,2,2^FS"), zpl);
    }

    @Test
    void laReferenciaNoCorreLasLineasDeAbajo() {
        String zpl = EmbalajeRenderer.campoZpl(List.of("REFERENCIA", "ENVASE: BOL-1"));

        // El encabezado ocupa un alto de línea normal: la primera línea de datos sigue en y=48.
        assertTrue(zpl.contains("^FO400,48^A0N,26,26^FB390,4,0,L^FDENVASE: BOL-1^FS"), zpl);
    }

    // -------------------------------------------------------------------------------------------
    // Qué SKU se lista en el aviso final
    // -------------------------------------------------------------------------------------------

    @Test
    void soloSeAvisaPorLoQueSalioImpreso() {
        assertTrue(EmbalajeRenderer.avisaSinEstandarizar(lineas(DatosEmbalaje.SIN_CARGAR)));
        // Con más de una unidad no se imprime nada, así que tampoco hay nada que listar.
        assertFalse(EmbalajeRenderer.avisaSinEstandarizar(
                EmbalajeRenderer.lineas(DatosEmbalaje.SIN_CARGAR, 2)));
        assertFalse(EmbalajeRenderer.avisaSinEstandarizar(lineas(datos("BOL-1", "", "", "", ""))));
    }
}
