package ar.com.leo.etiquetas.model;

/**
 * Datos de embalaje de un SKU, leídos del Excel de medidas.
 * Los campos de texto se imprimen tal cual en la etiqueta y nunca son null: una columna ausente o
 * vacía llega como cadena vacía.
 *
 * @param envase        código del envase (`BOL-1`, `CAJ-1`), o `NO` si el producto no lleva
 * @param inscripcion   texto escrito en el envase físico, resuelto desde la hoja ESTANDARIZACION
 * @param estandarizado lo calcula una fórmula del usuario y resume si el embalaje está cargado.
 *                      Es la única fuente de verdad al respecto: la app no lo deduce de los
 *                      demás campos.
 * @param aplica        el Excel tiene la columna ESTANDARIZADO. Distinguirlo de "la columna dice
 *                      NO" evita que un archivo anterior a esta función marque todos los SKU como
 *                      sin estandarizar, cuando en realidad no está en uso.
 */
public record DatosEmbalaje(
        String envase,
        String inscripcion,
        String rollo,
        String cantPanos,
        String observaciones,
        boolean estandarizado,
        boolean aplica
) {

    /** El Excel no tiene las columnas de embalaje: la función no está en uso. */
    public static final DatosEmbalaje VACIO = new DatosEmbalaje("", "", "", "", "", false, false);
}
