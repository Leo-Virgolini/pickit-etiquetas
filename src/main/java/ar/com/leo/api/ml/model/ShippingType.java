package ar.com.leo.api.ml.model;

import java.util.Set;

/** Tipo de envío de una orden ML, para clasificación y filtrado en la UI. */
public enum ShippingType {
    FLEX, COLECTA, TURBO, OTRO;

    /** Clasifica según el tag turbo y el logistic_type del shipment (mapeo estricto). */
    public static ShippingType from(boolean turbo, String logisticType) {
        if (turbo) return TURBO;
        if ("self_service".equals(logisticType)) return FLEX;
        if ("cross_docking".equals(logisticType)) return COLECTA;
        return OTRO;
    }

    /** Sin tipos tildados ⇒ pasa todo; con tildados ⇒ solo esos. */
    public static boolean passes(ShippingType type, Set<ShippingType> checked) {
        return checked.isEmpty() || checked.contains(type);
    }
}
