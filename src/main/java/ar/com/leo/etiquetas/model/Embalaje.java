package ar.com.leo.etiquetas.model;

/**
 * Embalaje predeterminado del catálogo (hoja EMBALAJES del Excel madre de medidas).
 * El {@code codigo} es lo que se imprime en la etiqueta; tipo y medidas son documentación
 * del catálogo y no se usan para calcular nada ni se suben a ML.
 */
public record Embalaje(
        String codigo,
        String tipo,
        Double anchoCm,
        Double altoCm,
        Double profundidadCm
) {
}
