package ar.com.leo.etiquetas.model;

public record MedidaSku(
        String sku,
        String producto,
        Double anchoCm,
        Double altoCm,
        Double profundidadCm,
        Double pesoKg,
        Double anchoMasCm,
        Double altoMasCm,
        Double profundidadMasCm,
        Double pesoMasKg,
        boolean subido,
        String error,
        String embalaje
) {

    public boolean estaMedido() {
        return anchoCm != null && anchoCm > 0
                && altoCm != null && altoCm > 0
                && profundidadCm != null && profundidadCm > 0
                && pesoKg != null && pesoKg > 0;
    }

    public boolean tieneMedidasParaSubir() {
        return anchoMasCm != null && anchoMasCm > 0
                && altoMasCm != null && altoMasCm > 0
                && profundidadMasCm != null && profundidadMasCm > 0
                && pesoMasKg != null && pesoMasKg > 0;
    }
}
