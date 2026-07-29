package ar.com.leo.etiquetas.model;

import java.util.List;

public record SortedLabelGroup(String zone, String sku, String productDescription, String details, String orderIds, List<ZplLabel> labels) {

    public SortedLabelGroup(String zone, String sku, String productDescription, String details, List<ZplLabel> labels) {
        this(zone, sku, productDescription, details, "", labels);
    }

    public int count() {
        return labels.size();
    }
}
