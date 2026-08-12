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

    private static DatosEmbalaje conCaja() {
        return new DatosEmbalaje("", "GRANDE", "3", "", "", "", "", "");
    }

    private static DatosEmbalaje conBolsa() {
        return new DatosEmbalaje("5", "", "", "", "", "", "", "");
    }

    private static DatosEmbalaje soloAgregados() {
        return new DatosEmbalaje("", "", "", "SI", "2", "DIAMANTE", "3", "Colchon");
    }

    // -------------------------------------------------------------------------------------------
    // Texto
    // -------------------------------------------------------------------------------------------

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
    void conCajaElEmbalajeEstaEnSi() {
        assertEquals(EstadoDato.SI, EstadoDato.embalajeDe(medida(null, conCaja())));
    }

    @Test
    void conBolsaElEmbalajeEstaEnSi() {
        assertEquals(EstadoDato.SI, EstadoDato.embalajeDe(medida(null, conBolsa())));
    }

    @Test
    void soloConAgregadosElEmbalajeEstaEnNo() {
        // Pluribol, rollo y observaciones no alcanzan: el criterio es el mismo que en la etiqueta.
        assertEquals(EstadoDato.NO, EstadoDato.embalajeDe(medida(null, soloAgregados())));
    }

    @Test
    void sinNingunDatoElEmbalajeEstaEnNo() {
        assertEquals(EstadoDato.NO, EstadoDato.embalajeDe(medida(null, DatosEmbalaje.VACIO)));
    }

    @Test
    void unSkuQueNoFiguraEnElExcelNoTieneEmbalaje() {
        assertEquals(EstadoDato.NO, EstadoDato.embalajeDe(null));
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
