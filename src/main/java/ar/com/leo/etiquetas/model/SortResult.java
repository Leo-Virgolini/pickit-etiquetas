package ar.com.leo.etiquetas.model;

import java.util.List;

public record SortResult(List<SortedLabelGroup> groups, LabelStatistics statistics) {

    public List<ZplLabel> sortedFlatList() {
        return groups.stream()
                .flatMap(g -> g.labels().stream())
                .toList();
    }
}
