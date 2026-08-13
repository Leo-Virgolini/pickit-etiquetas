package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;

import java.util.ArrayList;
import java.util.List;

/**
 * Arma las líneas de embalaje que van en la etiqueta a partir de lo cargado en el Excel, y su
 * fragmento ZPL. Un SKU lleva entre una y tres líneas —envase, rollo y observaciones—, o una sola
 * con el aviso si su columna ESTANDARIZADO dice que el embalaje todavía no está cargado.
 */
public final class EmbalajeRenderer {

    // Bloque en el margen superior derecho, entre el borde de la etiqueta y el separador de la zona
    // de picking (y=180).
    //
    // x=400 es el límite por la izquierda: entre y=129 y y=160 está el bloque "Pack ID: ..." de ML
    // —o "Venta ID: ...", según el tipo de envío—, cuyo número llega hasta x≈380 con sus 11 dígitos
    // y su fuente 30. Los 20 de diferencia son el margen por si alguna vez viene de 12. Arriba no
    // hay nada de ML: el texto "Recortá esta parte..." se elimina.
    private static final int X = 400;
    private static final int Y_INICIAL = 20;
    /** Hasta dónde puede llegar el bloque: el separador de la zona de picking está en y=180. */
    private static final int Y_LIMITE = 178;
    private static final int ALTO_LINEA = 26;
    private static final int FUENTE = 24;
    private static final int ANCHO = 390;
    /**
     * Caracteres que entran en una línea. Con la fuente 0 de ZPL, proporcional, un carácter ocupa
     * poco más de la mitad del alto nominal: 390 / (24 · 0,55) ≈ 29.
     */
    private static final int MAX_CARACTERES = 29;
    /**
     * Largo máximo de una palabra sin cortar. Deja lugar para el rótulo más largo ("OBS: ", 5
     * caracteres) en la misma línea que la primera pieza.
     */
    private static final int MAX_PALABRA = MAX_CARACTERES - 5;
    /** Marca de texto cortado. Tres puntos y no "…": la fuente residual puede no traer ese glifo. */
    private static final String ELIPSIS = "...";

    // El aviso de embalaje sin cargar no es un dato más de la lista: tiene que frenar al operario.
    // Va en un recuadro relleno con el texto en video inverso y centrado, y siempre es la única
    // línea del bloque. Con fuente 38 sus 16 caracteres en mayúscula ocupan ~365 de los 390 de
    // ancho; más grande arañaría el borde, y ^FB no descarta lo que no entra sino que lo reimprime
    // encima.
    private static final int FUENTE_AVISO = 38;
    private static final int ALTO_AVISO = 52;

    /** Reemplaza a todo el bloque cuando la columna ESTANDARIZADO no dice que sí. */
    private static final String SIN_ESTANDARIZAR = "NO ESTANDARIZADO";
    /** Encabeza el bloque cuando la etiqueta es de más de una unidad. */
    private static final String REFERENCIA = "REFERENCIA";

    private EmbalajeRenderer() {
    }

    /**
     * Líneas a imprimir para un SKU en una etiqueta de {@code cantidad} unidades.
     *
     * Con más de una unidad el operario no está embalando un producto suelto, así que el envase
     * cargado es orientativo: las líneas salen igual pero encabezadas por "REFERENCIA", y si no hay
     * nada cargado no sale nada, en vez del aviso que frena.
     */
    public static List<String> lineas(DatosEmbalaje datos, int cantidad) {
        List<String> lineas = new ArrayList<>();
        // Sin la columna ESTANDARIZADO en el Excel la función no está en uso: la etiqueta sale como
        // antes de existir, en vez de reclamar algo que el usuario todavía no puede cargar.
        if (datos == null || !datos.aplica()) return lineas;

        boolean referencia = cantidad > 1;

        // El aviso sale de la columna ESTANDARIZADO y no de si hay envase cargado: es una fórmula
        // del usuario que resume si completó envase, tipo de rollo y cantidad de paños. Cuando dice
        // que no, la etiqueta solo tiene que frenar al operario, sin ruido alrededor.
        if (!datos.estandarizado()) {
            if (referencia) return lineas;
            lineas.add(SIN_ESTANDARIZAR);
            return lineas;
        }

        if (referencia) lineas.add(REFERENCIA);

        // "NO" en el envase es una decisión tomada —ese producto no lleva caja ni bolsa— así que se
        // imprime como cualquier otro valor.
        if (cargado(datos.envase())) {
            lineas.add("ENVASE: " + unir(valor(datos.envase()), inscripcion(datos.inscripcion())));
        }

        if (cargado(datos.rollo())) {
            String linea = "ROLLO: " + valor(datos.rollo());
            linea += panos(datos.cantPanos());
            lineas.add(linea);
        }

        if (cargado(datos.observaciones())) lineas.add("OBS: " + valor(datos.observaciones()));

        // La fórmula del usuario puede decir que sí sin que haya ninguna de las tres columnas de
        // datos: son suyas y se ubican por encabezado, así que pueden faltar. Un "REFERENCIA" solo
        // no dice nada.
        if (lineas.size() == 1 && REFERENCIA.equals(lineas.get(0))) lineas.clear();

        return lineas;
    }

    /**
     * Si esas líneas son el aviso de embalaje sin cargar. El diálogo del final del lote lista
     * exactamente los SKU que salieron avisados en papel, así que se pregunta por lo que se imprime
     * y no por lo que dice el Excel.
     */
    public static boolean avisaSinEstandarizar(List<String> lineas) {
        return lineas != null && lineas.size() == 1 && SIN_ESTANDARIZAR.equals(lineas.get(0));
    }

    /**
     * Inscripción a imprimir. Las filas de la hoja de estandarización que no llevan inscripción
     * —las bolsas, y la fila "NO"— traen un guion, que no aporta nada en la etiqueta.
     */
    private static String inscripcion(String raw) {
        String v = valor(raw);
        return v.equals("-") ? "" : v;
    }

    /**
     * Sufijo con la cantidad de paños, o vacío si no hay ninguno.
     *
     * Se omite cuando es cero —la columna trae 0 en las filas sin rollo y "0 paños" sería ruido—
     * pero un valor no numérico se imprime tal cual: es una columna de texto libre y si el usuario
     * escribió "2-3", esconderlo sería peor que mostrarlo.
     */
    private static String panos(String raw) {
        String v = valor(raw);
        if (v.isEmpty()) return "";
        try {
            int cantidad = (int) Double.parseDouble(v.replace(',', '.'));
            if (cantidad <= 0) return "";
            return " - " + cantidad + (cantidad == 1 ? " paño" : " paños");
        } catch (NumberFormatException e) {
            // El valor real del Excel es del tipo "2 Y 1", que se lee bien con el sufijo. Si el
            // usuario ya escribió la palabra, no se repite.
            String enMayusculas = v.toUpperCase();
            boolean yaLoDice = enMayusculas.contains("PAÑO") || enMayusculas.contains("PANO");
            return " - " + v + (yaLoDice ? "" : " paños");
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

            if (SIN_ESTANDARIZAR.equals(linea)) {
                zpl.append(aviso(y));
                y += ALTO_AVISO;
                continue;
            }

            boolean ultima = i == lineas.size() - 1;
            // La última línea se queda con todo el alto libre que sobra, así que puede repartirse
            // en varias. Las anteriores no pueden crecer sin correr las de abajo, así que entran
            // en una sola — la inscripción del envase también es texto libre del usuario.
            int maxLineas = ultima ? Math.max(1, (Y_LIMITE - y) / ALTO_LINEA) : 1;
            // En los dos casos se acota el largo: con ^FB, ZPL no descarta lo que no entra, lo
            // reimprime encima de la última línea y queda una mancha ilegible.
            String texto = truncar(partirPalabrasLargas(linea), maxLineas * MAX_CARACTERES);

            zpl.append(campo(X, y, maxLineas, texto));
            // Negrita simulada con una segunda pasada corrida 1px, igual que ZONA y COD.EXT. El
            // encabezado de referencia va entero porque no tiene rótulo del que salir; en las demás
            // líneas se repasa solo el rótulo, que se superpone exactamente sobre el de la primera
            // pasada. Así no hay que calcular su ancho, que con una fuente proporcional no se puede
            // saber de antemano.
            String negrita = REFERENCIA.equals(linea) ? texto : rotulo(texto);
            if (!negrita.isEmpty()) zpl.append(campo(X + 1, y, maxLineas, negrita));

            y += ALTO_LINEA;
        }
        return zpl.toString();
    }

    /** Recuadro relleno con el aviso en video inverso, centrado vertical y horizontalmente. */
    private static String aviso(int y) {
        return "^FO" + X + ',' + y + "^GB" + ANCHO + ',' + ALTO_AVISO + ',' + ALTO_AVISO + "^FS\n"
                + "^FO" + X + ',' + (y + (ALTO_AVISO - FUENTE_AVISO) / 2)
                + "^A0N," + FUENTE_AVISO + ',' + FUENTE_AVISO
                + "^FB" + ANCHO + ",1,0,C^FR^FD" + SIN_ESTANDARIZAR + "^FS\n";
    }

    private static String campo(int x, int y, int maxLineas, String texto) {
        return "^FO" + x + ',' + y
                + "^A0N," + FUENTE + ',' + FUENTE
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
