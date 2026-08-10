package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.Embalaje;
import ar.com.leo.etiquetas.parser.EmbalajeResolver.Estado;
import ar.com.leo.etiquetas.parser.EmbalajeResolver.ResultadoEmbalaje;
import org.junit.jupiter.api.Test;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class EmbalajeResolverTest {

    private static final Map<String, Embalaje> CATALOGO = EmbalajeResolver.indexar(java.util.List.of(
            new Embalaje("CAJA 3", "CAJA", 30.0, 20.0, 15.0),
            new Embalaje("Bolsa Chica", "BOLSA", 25.0, 2.0, 18.0)
    ));

    @Test
    void codigoDelCatalogoResuelveOk() {
        ResultadoEmbalaje r = EmbalajeResolver.resolver("CAJA 3", CATALOGO);
        assertEquals("CAJA 3", r.textoEtiqueta());
        assertEquals(Estado.OK, r.estado());
    }

    @Test
    void celdaVaciaQuedaSinAsignar() {
        ResultadoEmbalaje r = EmbalajeResolver.resolver("   ", CATALOGO);
        assertEquals("-", r.textoEtiqueta());
        assertEquals(Estado.SIN_ASIGNAR, r.estado());
    }

    @Test
    void codigoNuloQuedaSinAsignar() {
        ResultadoEmbalaje r = EmbalajeResolver.resolver(null, CATALOGO);
        assertEquals("-", r.textoEtiqueta());
        assertEquals(Estado.SIN_ASIGNAR, r.estado());
    }

    @Test
    void codigoQueNoEstaEnElCatalogoEsInvalido() {
        ResultadoEmbalaje r = EmbalajeResolver.resolver("CAJA3", CATALOGO);
        assertEquals("-", r.textoEtiqueta());
        assertEquals(Estado.CODIGO_INVALIDO, r.estado());
    }

    @Test
    void laComparacionIgnoraMayusculasYEspaciosSobrantes() {
        ResultadoEmbalaje r = EmbalajeResolver.resolver("  caja   3 ", CATALOGO);
        assertEquals(Estado.OK, r.estado());
    }

    @Test
    void seImprimeElCodigoDelCatalogoNoElEscritoPorElUsuario() {
        ResultadoEmbalaje r = EmbalajeResolver.resolver("bolsa chica", CATALOGO);
        assertEquals("Bolsa Chica", r.textoEtiqueta());
    }

    @Test
    void conCatalogoVacioCualquierCodigoEsInvalido() {
        ResultadoEmbalaje r = EmbalajeResolver.resolver("CAJA 3", Map.of());
        assertEquals(Estado.CODIGO_INVALIDO, r.estado());
    }

    // ---------------------------------------------------------------------------------------
    // Campo ZPL
    // ---------------------------------------------------------------------------------------

    @Test
    void elCampoZplLlevaDosPasadasYAncholimitado() {
        String zpl = EmbalajeResolver.campoZpl("CAJA 3");
        assertEquals("^FO45,85^A0N,30,30^FB735,1,0,L^FDEMBALAJE: CAJA 3^FS\n"
                        + "^FO46,85^A0N,30,30^FB735,1,0,L^FDEMBALAJE: CAJA 3^FS\n",
                zpl);
    }

    @Test
    void elCampoZplNeutralizaLosCaracteresDeControlDeZpl() {
        // ^ y ~ cortarían el campo y el resto del ZPL se interpretaría como comandos.
        String zpl = EmbalajeResolver.campoZpl("CAJA ^3~A");
        assertTrue(zpl.contains("^FDEMBALAJE: CAJA  3 A^FS"), zpl);
    }

    @Test
    void sinTextoNoHayCampo() {
        assertEquals("", EmbalajeResolver.campoZpl(null));
        assertEquals("", EmbalajeResolver.campoZpl("  "));
    }
}
