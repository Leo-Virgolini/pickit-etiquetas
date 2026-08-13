package ar.com.leo.etiquetas.ui;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import ar.com.leo.etiquetas.model.MedidaSku;

import java.util.Map;

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
     * Si el Excel tiene la columna ESTANDARIZADO, mirando las filas que sí se leyeron. Hace falta
     * para los SKU ausentes, que no traen ninguna fila de donde deducirlo.
     */
    public static boolean moduloEmbalajeActivo(Map<String, MedidaSku> medidas) {
        return medidas != null && medidas.values().stream().anyMatch(m -> m.embalaje().aplica());
    }

    /**
     * Si esa etiqueta lleva la sección de embalaje. Además de ser un SKU elegible, tiene que ser de
     * una sola unidad: con dos o más el operario no está embalando un producto suelto, así que la
     * instrucción de envase no le sirve.
     */
    public static boolean llevaEmbalaje(String zona, String sku, int cantidad) {
        return cantidad == 1 && esSkuElegible(zona, sku);
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
