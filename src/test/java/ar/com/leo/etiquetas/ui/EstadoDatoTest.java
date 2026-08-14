package ar.com.leo.etiquetas.ui;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import ar.com.leo.etiquetas.model.MedidaSku;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

class EstadoDatoTest {

    private static MedidaSku medida(Double ancho, DatosEmbalaje embalaje) {
        return new MedidaSku("1241212", "Producto A",
                ancho, ancho, ancho, ancho,
                null, null, null, null,
                false, "", embalaje);
    }

    private static DatosEmbalaje estandarizado() {
        return new DatosEmbalaje("BOL-1", "9Y", "DIAMANTES", "2", "", true, true);
    }

    private static DatosEmbalaje sinEstandarizar() {
        // Con datos cargados pero la fórmula del Excel diciendo que falta algo.
        return new DatosEmbalaje("BOL-1", "9Y", "DIAMANTES", "2", "", false, true);
    }

    private static DatosEmbalaje soloLaFormula() {
        // La fórmula dice que sí, pero ninguna de las tres columnas está cargada.
        return new DatosEmbalaje("", "", "", "", "", true, true);
    }

    // -------------------------------------------------------------------------------------------
    // Texto
    // -------------------------------------------------------------------------------------------

    @Test
    void alCopiarUnaCeldaSePegaElTextoVisible() {
        // Los handlers de copia de la tabla serializan con toString(): sin esto se pegaría
        // "NO_APLICA" en vez de lo que el usuario ve.
        assertEquals("✓ SI", EstadoDato.SI.toString());
        assertEquals("—", EstadoDato.NO_APLICA.toString());
    }

    @Test
    void cadaEstadoTieneSuTexto() {
        assertEquals("✓ SI", EstadoDato.SI.texto());
        assertEquals("✘ NO", EstadoDato.NO.texto());
        assertEquals("⚠ SIN DATOS", EstadoDato.SIN_DATOS.texto());
        assertEquals("—", EstadoDato.NO_APLICA.texto());
    }

    // -------------------------------------------------------------------------------------------
    // Embalaje
    // -------------------------------------------------------------------------------------------

    @Test
    void conLaFormulaEnSiElEstadoEsSi() {
        assertEquals(EstadoDato.SI, EstadoDato.estandarizadoDe(estandarizado()));
    }

    @Test
    void conLaFormulaEnNoElEstadoEsNo() {
        // Aunque tenga envase y rollo cargados: la fórmula del Excel es la fuente de verdad.
        assertEquals(EstadoDato.NO, EstadoDato.estandarizadoDe(sinEstandarizar()));
    }

    @Test
    void sinLaColumnaEnElExcelNoAplica() {
        assertEquals(EstadoDato.NO_APLICA, EstadoDato.estandarizadoDe(DatosEmbalaje.VACIO));
    }

    @Test
    void conLaFormulaEnSiYSinNingunaColumnaElEstadoLoDice() {
        // La etiqueta de ese SKU imprime "SIN DATOS DE EMBALAJE": si la tabla dijera "SI" no habría
        // forma de saber cuál de las dos tiene razón.
        assertEquals(EstadoDato.SIN_DATOS, EstadoDato.estandarizadoDe(soloLaFormula()));
    }

    @Test
    void alcanzaConUnaDeLasTresColumnasParaQueElEstadoSeaSi() {
        DatosEmbalaje soloObservaciones = new DatosEmbalaje("", "", "", "", "FRAGIL", true, true);

        assertEquals(EstadoDato.SI, EstadoDato.estandarizadoDe(soloObservaciones));
    }

    @Test
    void elEstadoNoDependeDeLaCantidadDeLaEtiqueta() {
        // Es un dato del SKU: sus etiquetas de una unidad y de varias comparten la fila del Excel,
        // y desde que los avisos salen siempre, también comparten lo que se imprime.
        assertEquals(EstadoDato.NO, EstadoDato.estandarizadoDe(DatosEmbalaje.SIN_CARGAR));
    }

    // -------------------------------------------------------------------------------------------
    // Elegibilidad
    // -------------------------------------------------------------------------------------------

    @Test
    void losCarrosNoLlevanNingunoDeLosDosDatos() {
        assertEquals(false, EstadoDato.esSkuElegible("CARROS", "1241212"));
    }

    @Test
    void unSkuNoNumericoNoEsElegible() {
        // "SKU INVALIDO: ..." es el sentinela del parser: nunca llega al Excel de medidas.
        assertEquals(false, EstadoDato.esSkuElegible("J1-D", "SKU INVALIDO: ABC"));
    }

    @Test
    void unGrupoConVariosSkusNoEsElegible() {
        assertEquals(false, EstadoDato.esSkuElegible("J1-D", "1241212\n1241255"));
    }

    @Test
    void unSkuNumericoDeZonaComunEsElegible() {
        assertEquals(true, EstadoDato.esSkuElegible("J1-D", "1241212"));
    }

    @Test
    void unSkuVacioNoEsElegible() {
        assertEquals(false, EstadoDato.esSkuElegible("J1-D", "  "));
        assertEquals(false, EstadoDato.esSkuElegible("J1-D", null));
    }

    // -------------------------------------------------------------------------------------------
    // SKU que no figura en el Excel
    // -------------------------------------------------------------------------------------------

    @Test
    void unSkuAusenteDelExcelCuentaComoSinEstandarizar() {
        // Todavía no tiene fila, así que nadie le cargó el envase: la etiqueta tiene que avisarlo.
        assertEquals(DatosEmbalaje.SIN_CARGAR, EstadoDato.embalajeDe(null, true));
        assertEquals(EstadoDato.NO, EstadoDato.estandarizadoDe(EstadoDato.embalajeDe(null, true)));
    }

    @Test
    void unSkuAusenteConElModuloApagadoNoInformaNada() {
        // Sin la columna en el Excel la función no está en uso: reclamar embalaje sería ruido.
        assertEquals(DatosEmbalaje.VACIO, EstadoDato.embalajeDe(null, false));
        assertEquals(EstadoDato.NO_APLICA, EstadoDato.estandarizadoDe(EstadoDato.embalajeDe(null, false)));
    }

    @Test
    void unSkuPresenteUsaElEmbalajeDeSuFila() {
        MedidaSku medida = medida(null, estandarizado());

        assertEquals(estandarizado(), EstadoDato.embalajeDe(medida, true));
    }
}
