package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.Embalaje;

import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Resuelve el código de embalaje cargado en la fila del SKU contra el catálogo de la hoja EMBALAJES.
 * Devuelve el texto a imprimir en la etiqueta y en qué estado quedó la asignación, para que la app
 * pueda avisar tanto los SKU sin embalaje como los que tienen un código con typo.
 */
public final class EmbalajeResolver {

    /** Texto que se imprime cuando no hay un embalaje válido para el SKU. */
    public static final String SIN_DATO = "-";

    public enum Estado {
        /** El SKU tiene un código que existe en el catálogo. */
        OK,
        /** La celda EMBALAJE está vacía, o el SKU no figura en el Excel de medidas. */
        SIN_ASIGNAR,
        /** La celda tiene un código que no figura en el catálogo (típicamente un typo). */
        CODIGO_INVALIDO
    }

    public record ResultadoEmbalaje(String textoEtiqueta, Estado estado) {
    }

    private EmbalajeResolver() {
    }

    /** Indexa el catálogo por código normalizado, preservando el código original dentro del record. */
    public static Map<String, Embalaje> indexar(Collection<Embalaje> embalajes) {
        Map<String, Embalaje> index = new LinkedHashMap<>();
        for (Embalaje e : embalajes) {
            if (e == null || e.codigo() == null || e.codigo().isBlank()) continue;
            index.putIfAbsent(normalizar(e.codigo()), e);
        }
        return index;
    }

    public static ResultadoEmbalaje resolver(String codigoCrudo, Map<String, Embalaje> catalogo) {
        if (codigoCrudo == null || codigoCrudo.isBlank()) {
            return new ResultadoEmbalaje(SIN_DATO, Estado.SIN_ASIGNAR);
        }
        Embalaje embalaje = catalogo == null ? null : catalogo.get(normalizar(codigoCrudo));
        if (embalaje == null) {
            return new ResultadoEmbalaje(SIN_DATO, Estado.CODIGO_INVALIDO);
        }
        // Se imprime el código tal como figura en el catálogo (no como lo escribió el usuario en la
        // fila del SKU) para que todas las etiquetas muestren la misma forma del mismo embalaje.
        return new ResultadoEmbalaje(embalaje.codigo(), Estado.OK);
    }

    /** Mayúsculas, sin espacios en los extremos y con los espacios internos colapsados. */
    public static String normalizar(String raw) {
        if (raw == null) return "";
        return raw.replace('\n', ' ')
                .replaceAll("\\s+", " ")
                .trim()
                .toUpperCase();
    }
}
