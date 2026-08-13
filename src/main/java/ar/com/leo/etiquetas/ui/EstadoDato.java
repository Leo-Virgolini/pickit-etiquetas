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

    /**
     * Los handlers de copia de la tabla serializan cada celda con toString(): sin esto, copiar una
     * fila pegaría el nombre de la constante ("NO_APLICA") en vez de lo que el usuario ve.
     */
    @Override
    public String toString() {
        return texto;
    }

    /** Las cuatro columnas base cm/kg cargadas. Un SKU que no figura en el Excel cuenta como NO. */
    public static EstadoDato medidasDe(MedidaSku medida) {
        return medida != null && medida.estaMedido() ? SI : NO;
    }

    /**
     * Sale de la columna ESTANDARIZADO del Excel, que es una fórmula del usuario: resume si
     * completó envase, tipo de rollo y cantidad de paños. La app no lo deduce por su cuenta, así
     * que la tabla y la etiqueta siempre dicen lo mismo.
     */
    public static EstadoDato estandarizadoDe(MedidaSku medida) {
        if (medida == null) return NO;
        // Sin la columna en el Excel no hay nada que informar: la función no está en uso.
        if (!medida.embalaje().aplica()) return NO_APLICA;
        return medida.embalaje().estandarizado() ? SI : NO;
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
