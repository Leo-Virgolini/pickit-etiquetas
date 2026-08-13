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
        return new DatosEmbalaje("BOL-1", "9Y", "DIAMANTES", "2", "", true);
    }

    private static DatosEmbalaje sinEstandarizar() {
        // Con datos cargados pero la fórmula del Excel diciendo que falta algo.
        return new DatosEmbalaje("BOL-1", "9Y", "DIAMANTES", "2", "", false);
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
        assertEquals("—", EstadoDato.NO_APLICA.texto());
    }

    // -------------------------------------------------------------------------------------------
    // Medidas
    // -------------------------------------------------------------------------------------------

    @Test
    void unSkuMedidoEstaEnSi() {
        assertEquals(EstadoDato.SI, EstadoDato.medidasDe(medida(10.0, DatosEmbalaje.VACIO)));
    }

    @Test
    void unSkuSinMedidasEstaEnNo() {
        assertEquals(EstadoDato.NO, EstadoDato.medidasDe(medida(null, DatosEmbalaje.VACIO)));
    }

    @Test
    void unSkuQueNoFiguraEnElExcelEstaEnNo() {
        assertEquals(EstadoDato.NO, EstadoDato.medidasDe(null));
    }

    // -------------------------------------------------------------------------------------------
    // Embalaje
    // -------------------------------------------------------------------------------------------

    @Test
    void conLaFormulaEnSiElEstadoEsSi() {
        assertEquals(EstadoDato.SI, EstadoDato.estandarizadoDe(medida(null, estandarizado())));
    }

    @Test
    void conLaFormulaEnNoElEstadoEsNo() {
        // Aunque tenga envase y rollo cargados: la fórmula del Excel es la fuente de verdad.
        assertEquals(EstadoDato.NO, EstadoDato.estandarizadoDe(medida(null, sinEstandarizar())));
    }

    @Test
    void sinNingunDatoElEstadoEsNo() {
        assertEquals(EstadoDato.NO, EstadoDato.estandarizadoDe(medida(null, DatosEmbalaje.VACIO)));
    }

    @Test
    void unSkuQueNoFiguraEnElExcelNoEstaEstandarizado() {
        assertEquals(EstadoDato.NO, EstadoDato.estandarizadoDe(null));
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
}
