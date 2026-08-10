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

    /**
     * Fragmento ZPL con la línea "EMBALAJE: <código>", listo para inyectar en el bloque de
     * coordenadas absolutas (^LH0,0) del inicio de la etiqueta. Cadena vacía si no hay texto.
     *
     * Va en y=85: debajo del #N (que termina en y=65) y del banner MEDIR (que llega a y=80), y por
     * encima del "Pack ID:" del formato de ML (y=130). El ^FB acota el ancho para que un código
     * largo no se derrame fuera del área imprimible, y la doble pasada con 1px de offset simula
     * negrita igual que ZONA y COD.EXT.
     */
    public static String campoZpl(String codigo) {
        if (codigo == null || codigo.isBlank()) return "";
        String texto = "EMBALAJE: " + sanitizar(codigo);
        return "^FO45,85^A0N,30,30^FB735,1,0,L^FD" + texto + "^FS\n"
                + "^FO46,85^A0N,30,30^FB735,1,0,L^FD" + texto + "^FS\n";
    }

    /**
     * ^ y ~ son los prefijos de comando de ZPL: dentro de un ^FD cortarían el campo y el resto de
     * la etiqueta se interpretaría como comandos. El código lo tipea el usuario en el Excel, así
     * que se neutralizan antes de inyectarlo.
     */
    private static String sanitizar(String codigo) {
        return codigo.replace('^', ' ').replace('~', ' ');
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
