package ar.com.leo.etiquetas.model;

/**
 * Datos de embalaje que el usuario carga a mano en el Excel de medidas, uno por SKU.
 * Todos los campos son texto: se imprimen tal cual en la etiqueta y la app no opera con ellos.
 * Nunca son null — una columna ausente o vacía llega como cadena vacía.
 */
public record DatosEmbalaje(
        String nroBolsa,
        String nombreCaja,
        String nroCaja,
        String pluribol,
        String cantPluribol,
        String rollo,
        String cantPanos,
        String observaciones
) {

    public static final DatosEmbalaje VACIO = new DatosEmbalaje("", "", "", "", "", "", "", "");

    /**
     * Un SKU está embalado cuando tiene caja o bolsa. El pluribol, el rollo y las observaciones
     * son agregados: no alcanzan por sí solos para considerarlo resuelto.
     */
    public boolean tieneCajaOBolsa() {
        return cargado(nroBolsa) || cargado(nombreCaja) || cargado(nroCaja);
    }

    private static boolean cargado(String valor) {
        return valor != null && !valor.isBlank();
    }
}
