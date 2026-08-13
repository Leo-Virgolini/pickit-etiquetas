package ar.com.leo.etiquetas.ui;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import ar.com.leo.etiquetas.model.MedidaSku;

/**
 * Estado de la columna ESTANDARIZADO de un SKU, para mostrarlo en la tabla de etiquetas.
 *
 * Los criterios son los mismos que usa la etiqueta y viven acá para que no puedan divergir: si la
 * tabla dijera que un SKU tiene el embalaje cargado y la etiqueta imprimiera "NO ESTANDARIZADO",
 * no habría forma de saber cuál de las dos tiene razón.
 */
public enum EstadoDato {

    SI("✓ SI"),
    NO("✘ NO"),
    /** Ese SKU nunca lleva el dato: es un carro, un SKU no numérico, o la función está apagada. */
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

    /**
     * Sale de la columna ESTANDARIZADO del Excel, que es una fórmula del usuario: resume si
     * completó envase, tipo de rollo y cantidad de paños. La app no lo deduce por su cuenta, así
     * que la tabla y la etiqueta siempre dicen lo mismo.
     */
    public static EstadoDato estandarizadoDe(DatosEmbalaje datos) {
        if (datos == null) return NO;
        // Sin la columna en el Excel no hay nada que informar: la función no está en uso.
        if (!datos.aplica()) return NO_APLICA;
        return datos.estandarizado() ? SI : NO;
    }

    /**
     * Datos de embalaje a usar para un SKU. La etiqueta y la tabla parten de este mismo objeto, así
     * que no pueden discrepar.
     *
     * Un SKU que no figura en el Excel cuenta como sin estandarizar: se lo agrega en este mismo
     * lote, pero recién ahí alguien va a poder cargarle el envase, y mientras tanto el operario
     * tiene que enterarse. Salvo que el módulo esté apagado, donde no hay nada que reclamar.
     */
    public static DatosEmbalaje embalajeDe(MedidaSku medida, boolean moduloActivo) {
        if (medida != null) return medida.embalaje();
        return moduloActivo ? DatosEmbalaje.SIN_CARGAR : DatosEmbalaje.VACIO;
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
