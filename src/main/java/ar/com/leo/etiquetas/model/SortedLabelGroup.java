package ar.com.leo.etiquetas.model;

import java.util.List;

public record SortedLabelGroup(String zone, String sku, String productDescription, String details,
                               List<ZplLabel> labels) {
}
