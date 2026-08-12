package ar.com.leo.etiquetas.ui;

import ar.com.leo.etiquetas.model.MedidaSku;

/**
 * Estado de un dato del Excel de medidas para mostrarlo en la tabla de etiquetas.
 *
 * Los criterios son los mismos que usan la etiqueta y el aviso post-proceso, y viven acá para que
 * no puedan divergir: si la tabla dijera que un SKU tiene embalaje y la etiqueta imprimiera
 * "NO ESTANDARIZADO", no habría forma de saber cuál de las dos tiene razón.
 */
public enum EstadoDato {

    SI("✓ SI"),
    NO("✘ NO"),
    /** Ese SKU nunca lleva el dato: es un carro, un SKU no numérico, o el módulo está apagado. */
    NO_APLICA("—");

    private final String texto;

    EstadoDato(String texto) {
        this.texto = texto;
    }

    public String texto() {
        return texto;
    }

    /** Las cuatro columnas base cm/kg cargadas. Un SKU que no figura en el Excel cuenta como NO. */
    public static EstadoDato medidasDe(MedidaSku medida) {
        return medida != null && medida.estaMedido() ? SI : NO;
    }

    /**
     * Caja o bolsa cargada. Los agregados (rollo, observaciones) no cuentan, igual que
     * en la línea "NO ESTANDARIZADO" de la etiqueta.
     */
    public static EstadoDato embalajeDe(MedidaSku medida) {
        return medida != null && medida.embalaje().tieneCajaOBolsa() ? SI : NO;
    }

    /**
     * Un SKU lleva medidas y embalaje si tiene número propio y su etiqueta corresponde a un solo
     * producto. Los carros listan varios; los no numéricos son sentinelas del parser
     * ("SKU INVALIDO: ...") que nunca llegan al Excel de medidas.
     */
    public static boolean esSkuElegible(String zona, String sku) {
        return !"CARROS".equals(zona)
                && sku != null && !sku.isBlank()
                && !sku.contains("\n")
                && sku.matches("\\d+");
    }
}
