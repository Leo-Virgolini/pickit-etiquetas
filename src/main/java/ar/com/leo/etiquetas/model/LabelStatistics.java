package ar.com.leo.etiquetas.model;

import java.util.Map;

public record LabelStatistics(int totalLabels, Map<String, Integer> countByZone, int uniqueSkus, int unmappedLabels) {
}
