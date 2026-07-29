package ar.com.leo.etiquetas.sorter;

import ar.com.leo.etiquetas.model.LabelStatistics;
import ar.com.leo.etiquetas.model.SortResult;
import ar.com.leo.etiquetas.model.SortedLabelGroup;
import ar.com.leo.etiquetas.model.ZplLabel;

import java.util.*;
import java.util.stream.Collectors;

public class LabelSorter {

    private static final String UNKNOWN = "???";

    private int zoneGroupPriority(String zone) {
        String z = zone.toUpperCase();
        if (z.startsWith("J")) return 0;
        if (z.startsWith("TURBOS")) return 4;
        if (z.startsWith("T")) return 1;
        if (z.startsWith("COMBOS")) return 2;
        if (z.startsWith("CARROS")) return 3;
        if (z.startsWith("RETIROS")) return 5;
        return Integer.MAX_VALUE;
    }

    public SortResult sort(List<ZplLabel> labels, Map<String, String> skuToZone) {
        Map<String, List<ZplLabel>> grouped = labels.stream()
                .collect(Collectors.groupingBy(
                        l -> resolveZoneForLabel(l, skuToZone) + "|" + (l.sku() != null ? l.sku() : ""),
                        LinkedHashMap::new,
                        Collectors.toList()));

        List<SortedLabelGroup> groups = grouped.entrySet().stream()
                .map(entry -> {
                    String[] parts = entry.getKey().split("\\|", 2);
                    String zone = parts[0];
                    String sku = parts.length > 1 ? parts[1] : "";
                    String desc = entry.getValue().stream()
                            .map(ZplLabel::productDescription)
                            .filter(Objects::nonNull)
                            .findFirst()
                            .orElse("");
                    String details = entry.getValue().stream()
                            .map(ZplLabel::details)
                            .filter(Objects::nonNull)
                            .findFirst()
                            .orElse("");
                    String orderIds = entry.getValue().stream()
                            .map(ZplLabel::orderIds)
                            .filter(o -> o != null && !o.isEmpty())
                            .findFirst()
                            .orElse("");
                    return new SortedLabelGroup(zone, sku, desc, details, orderIds, entry.getValue());
                })
                .sorted(Comparator
                        .<SortedLabelGroup>comparingInt(g -> zoneGroupPriority(g.zone()))
                        .thenComparing(g -> g.zone().toUpperCase())
                        .thenComparing(g -> {
                            try {
                                return Long.parseLong(g.sku());
                            } catch (NumberFormatException e) {
                                return Long.MAX_VALUE;
                            }
                        }))
                .toList();

        LabelStatistics stats = buildStatistics(labels, skuToZone);
        return new SortResult(groups, stats);
    }

    private String resolveZoneForLabel(ZplLabel label, Map<String, String> skuToZone) {
        if (label.turbo()) return "TURBOS";
        return resolveZone(label.sku(), skuToZone);
    }

    private String resolveZone(String sku, Map<String, String> skuToZone) {
        if (sku == null || sku.isEmpty()) {
            return UNKNOWN;
        }
        if (sku.contains("\n")) {
            // CARROS solo si hay 2+ SKUs distintos
            long distinct = Arrays.stream(sku.split("\n"))
                    .filter(s -> !s.isBlank())
                    .distinct()
                    .count();
            if (distinct > 1) return "CARROS";
            // Si todos son el mismo SKU, resolver como SKU individual
            String singleSku = sku.split("\n")[0].trim();
            return skuToZone.getOrDefault(singleSku, UNKNOWN);
        }
        return skuToZone.getOrDefault(sku, UNKNOWN);
    }

    private LabelStatistics buildStatistics(List<ZplLabel> labels, Map<String, String> skuToZone) {
        int total = labels.size();
        Map<String, Integer> countByZone = new LinkedHashMap<>();
        Set<String> uniqueSkus = new HashSet<>();
        int unmapped = 0;

        for (ZplLabel label : labels) {
            String zone = resolveZoneForLabel(label, skuToZone);
            countByZone.merge(zone, 1, Integer::sum);
            if (label.sku() != null) {
                for (String s : label.sku().split("\n")) {
                    String trimmed = s.trim();
                    if (!trimmed.isEmpty()) uniqueSkus.add(trimmed);
                }
            }
            if (zone.equals(UNKNOWN)) {
                unmapped++;
            }
        }

        return new LabelStatistics(total, countByZone, uniqueSkus.size(), unmapped);
    }
}
