package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;

import java.util.ArrayList;
import java.util.List;

/**
 * Arma las líneas de embalaje que van en la etiqueta a partir de lo cargado en el Excel, y su
 * fragmento ZPL. Solo aparecen las líneas cuyos datos estén cargados: un SKU lleva entre una y tres
 * (caja o bolsa, rollo y observaciones), o una sola con el aviso si no tiene embalaje asignado.
 */
public final class EmbalajeRenderer {

    // Bloque en el margen superior derecho, entre el borde de la etiqueta y el separador de la zona
    // de picking (y=180).
    //
    // x=410 es el límite por la izquierda: entre y=129 y y=160 está el bloque "Pack ID: ..." de ML,
    // cuyo número llega hasta x≈400 con su fuente 30. Arriba no hay nada de ML —el texto "Recortá
    // esta parte..." se elimina— y el banner MEDIR se ubica a la izquierda, debajo del #N.
    private static final int X = 410;
    private static final int Y_INICIAL = 20;
    /** Hasta dónde puede llegar el bloque: el separador de la zona de picking está en y=180. */
    private static final int Y_LIMITE = 178;
    private static final int ALTO_LINEA = 24;
    private static final int FUENTE = 22;
    // El aviso de embalaje sin cargar va más grande y en negrita: es lo que tiene que frenar al
    // operario, no un dato más de la lista. Siempre es la única línea del bloque.
    private static final int FUENTE_AVISO = 30;
    private static final int ANCHO = 380;
    /**
     * Caracteres que entran en una línea. Con la fuente 0 de ZPL, proporcional, un carácter ocupa
     * poco más de la mitad del alto nominal: 380 / (22 · 0,55) ≈ 31, redondeado para abajo.
     */
    private static final int MAX_CARACTERES = 30;
    /**
     * Largo máximo de una palabra sin cortar. Deja lugar para el rótulo más largo ("OBS: ", 5
     * caracteres) en la misma línea que la primera pieza.
     */
    private static final int MAX_PALABRA = MAX_CARACTERES - 5;
    /** Marca de texto cortado. Tres puntos y no "…": la fuente residual puede no traer ese glifo. */
    private static final String ELIPSIS = "...";

    /** Reemplaza a la línea de caja o bolsa cuando el SKU no tiene ninguna de las dos cargada. */
    private static final String SIN_ESTANDARIZAR = "NO ESTANDARIZADO";

    private EmbalajeRenderer() {
    }

    public static List<String> lineas(DatosEmbalaje datos) {
        List<String> lineas = new ArrayList<>();
        if (datos == null) return lineas;

        // El aviso sale de la columna ESTANDARIZADO y no de si hay envase cargado: es una fórmula
        // del usuario que resume si completó envase, tipo de rollo y cantidad de paños. Cuando dice
        // que no, la etiqueta solo tiene que frenar al operario, sin ruido alrededor.
        if (!datos.estandarizado()) {
            lineas.add(SIN_ESTANDARIZAR);
            return lineas;
        }

        // "NO" en el envase es una decisión tomada —ese producto no lleva caja ni bolsa— así que se
        // imprime como cualquier otro valor.
        if (cargado(datos.envase())) {
            lineas.add("ENVASE: " + unir(valor(datos.envase()), inscripcion(datos.inscripcion())));
        }

        if (cargado(datos.rollo())) {
            String linea = "ROLLO: " + valor(datos.rollo());
            // Los paños solo si son más de cero: la columna trae 0 en las filas sin rollo.
            int panos = entero(datos.cantPanos());
            if (panos > 0) linea += " - " + panos + (panos == 1 ? " paño" : " paños");
            lineas.add(linea);
        }

        if (cargado(datos.observaciones())) lineas.add("OBS: " + valor(datos.observaciones()));

        return lineas;
    }

    /**
     * Inscripción a imprimir. Las filas de la hoja de estandarización que no llevan inscripción
     * —las bolsas, y la fila "NO"— traen un guion, que no aporta nada en la etiqueta.
     */
    private static String inscripcion(String raw) {
        String v = valor(raw);
        return v.equals("-") ? "" : v;
    }

    /** Cantidad de paños, o 0 si la celda está vacía o no es un número. */
    private static int entero(String raw) {
        try {
            return (int) Double.parseDouble(valor(raw).replace(',', '.'));
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    /**
     * Fragmento ZPL con las líneas apiladas, listo para inyectar en el bloque de coordenadas
     * absolutas (^LH0,0) del inicio de la etiqueta. Cadena vacía si no hay ninguna línea.
     *
     * El ^FB acota cada línea al ancho disponible: un valor largo se recorta en vez de derramarse
     * fuera del área imprimible.
     */
    public static String campoZpl(List<String> lineas) {
        if (lineas == null || lineas.isEmpty()) return "";

        StringBuilder zpl = new StringBuilder();
        int y = Y_INICIAL;
        for (int i = 0; i < lineas.size(); i++) {
            String linea = lineas.get(i);
            boolean aviso = SIN_ESTANDARIZAR.equals(linea);
            boolean ultima = i == lineas.size() - 1;

            int fuente = aviso ? FUENTE_AVISO : FUENTE;
            // La última línea se queda con todo el alto libre que sobra, así que puede repartirse
            // en varias. Las anteriores no pueden crecer sin correr las de abajo, así que entran
            // en una sola — el nombre de caja también es texto libre del usuario.
            int maxLineas = ultima ? Math.max(1, (Y_LIMITE - y) / ALTO_LINEA) : 1;
            // En los dos casos se acota el largo: con ^FB, ZPL no descarta lo que no entra, lo
            // reimprime encima de la última línea y queda una mancha ilegible.
            String texto = truncar(partirPalabrasLargas(linea), maxLineas * MAX_CARACTERES);

            zpl.append(campo(X, y, fuente, maxLineas, texto));
            // Negrita simulada con una segunda pasada corrida 1px, igual que ZONA y COD.EXT. El
            // aviso va entero; en las demás líneas se repasa solo el rótulo, que se superpone
            // exactamente sobre el de la primera pasada. Así no hay que calcular su ancho, que con
            // una fuente proporcional no se puede saber de antemano.
            String negrita = aviso ? texto : rotulo(texto);
            if (!negrita.isEmpty()) zpl.append(campo(X + 1, y, fuente, maxLineas, negrita));

            y += ALTO_LINEA;
        }
        return zpl.toString();
    }

    private static String campo(int x, int y, int fuente, int maxLineas, String texto) {
        return "^FO" + x + ',' + y
                + "^A0N," + fuente + ',' + fuente
                + "^FB" + ANCHO + ',' + maxLineas + ",0,L"
                + "^FD" + sanitizar(texto) + "^FS\n";
    }

    /** El rótulo de la línea, con sus dos puntos, o vacío si no tiene. */
    private static String rotulo(String texto) {
        int corte = texto.indexOf(':');
        return corte == -1 ? "" : texto.substring(0, corte + 1);
    }

    /**
     * ^ y ~ son los prefijos de comando de ZPL: dentro de un ^FD cortarían el campo y el resto de
     * la etiqueta se interpretaría como comandos. Los valores los tipea el usuario en el Excel,
     * así que se neutralizan antes de inyectarlos.
     */
    private static String sanitizar(String texto) {
        return texto.replace('^', ' ').replace('~', ' ');
    }

    /**
     * Inserta cortes en las palabras que no entran en una línea. ^FB corta por palabras: una
     * palabra más larga que el ancho se baja entera a la línea siguiente, dejando el rótulo solo
     * arriba ("OBS:" en una línea y el texto en la de abajo). Pasa con códigos y URLs, que no
     * tienen espacios donde cortar.
     */
    private static String partirPalabrasLargas(String texto) {
        StringBuilder salida = new StringBuilder(texto.length() + 8);
        for (String palabra : texto.split(" ")) {
            if (!salida.isEmpty()) salida.append(' ');
            for (int i = 0; i < palabra.length(); i += MAX_PALABRA) {
                if (i > 0) salida.append(' ');
                salida.append(palabra, i, Math.min(i + MAX_PALABRA, palabra.length()));
            }
        }
        return salida.toString();
    }

    private static String truncar(String texto, int maxCaracteres) {
        if (texto.length() <= maxCaracteres) return texto;
        return texto.substring(0, maxCaracteres - ELIPSIS.length()) + ELIPSIS;
    }

    /** Une número y nombre con " - ", omitiendo el separador si falta alguno. */
    private static String unir(String primero, String segundo) {
        if (primero.isEmpty()) return segundo;
        if (segundo.isEmpty()) return primero;
        return primero + " - " + segundo;
    }

    private static boolean cargado(String valor) {
        return valor != null && !valor.isBlank();
    }

    /**
     * Texto de una celda listo para imprimir: sin saltos de línea ni espacios de más. El usuario
     * puede haber usado Alt+Enter en el Excel, y un LF crudo dentro de un ^FD pega las palabras
     * porque la impresora lo ignora.
     */
    private static String valor(String raw) {
        if (raw == null) return "";
        return raw.replaceAll("\\s+", " ").trim();
    }
}
