package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;

import java.util.ArrayList;
import java.util.List;

/**
 * Arma las líneas de embalaje que van en la etiqueta a partir de lo cargado en el Excel, y su
 * fragmento ZPL. Solo aparecen las líneas cuyos datos estén cargados: un SKU puede llevar desde
 * ninguna hasta tres (caja o bolsa, rollo y observaciones).
 */
public final class EmbalajeRenderer {

    // Bloque en el margen superior derecho, entre el borde de la etiqueta y el separador de la zona
    // de picking (y=180). Seis líneas llegan hasta y≈162.
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
    // operario, no un dato más de la lista.
    private static final int ALTO_LINEA_AVISO = 34;
    private static final int FUENTE_AVISO = 30;
    private static final int ANCHO = 380;
    /**
     * Caracteres que entran en una línea. Con la fuente 0 de ZPL, proporcional, un carácter ocupa
     * poco más de la mitad del alto nominal: 380 / (22 · 0,55) ≈ 31, redondeado para abajo.
     */
    private static final int MAX_CARACTERES = 36;
    /**
     * Largo máximo de una palabra sin cortar. Es menor que MAX_CARACTERES para que la primera pieza
     * entre en la misma línea que el rótulo ("OBS: " son 5 caracteres).
     */
    private static final int MAX_PALABRA = 30;

    /** Reemplaza a la línea de caja o bolsa cuando el SKU no tiene ninguna de las dos cargada. */
    private static final String SIN_ESTANDARIZAR = "NO ESTANDARIZADO";

    private EmbalajeRenderer() {
    }

    public static List<String> lineas(DatosEmbalaje datos) {
        List<String> lineas = new ArrayList<>();
        if (datos == null) return lineas;

        // Caja y bolsa son excluyentes: si están las dos cargadas gana la caja. Puede pasar cuando
        // cambia el embalaje y queda el número de bolsa viejo; con las dos salen cinco líneas y la
        // última se imprimiría sobre el separador de la zona de picking.
        //
        // Sin ninguna de las dos la línea igual se imprime, avisando: el operario tiene que poder
        // distinguir "a este SKU todavía no le cargaron el embalaje" de una etiqueta generada sin
        // esta función, en vez de embalar a criterio propio.
        String caja = unir(valor(datos.nroCaja()), valor(datos.nombreCaja()));
        if (!caja.isEmpty()) lineas.add("CAJA: " + caja);
        else if (cargado(datos.nroBolsa())) lineas.add("BOLSA: " + valor(datos.nroBolsa()));
        else lineas.add(SIN_ESTANDARIZAR);

        if (cargado(datos.rollo())) {
            String linea = "ROLLO: " + valor(datos.rollo());
            if (cargado(datos.cantPanos())) linea += " - " + valor(datos.cantPanos()) + " paños";
            lineas.add(linea);
        }

        if (cargado(datos.observaciones())) lineas.add("OBS: " + valor(datos.observaciones()));

        return lineas;
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
            // en varias: con ^FB de una sola línea ZPL no descarta el sobrante, lo reimprime encima
            // de la misma y queda ilegible. Las anteriores no pueden crecer sin correr las de
            // abajo, así que se cortan — el nombre de caja también es texto libre del usuario.
            int maxLineas = ultima ? Math.max(1, (Y_LIMITE - y) / ALTO_LINEA) : 1;
            String texto = ultima ? partirPalabrasLargas(linea) : truncar(linea);

            zpl.append(campo(X, y, fuente, maxLineas, texto));
            // El aviso va en negrita, simulada con una segunda pasada corrida 1px, igual que ZONA
            // y COD.EXT.
            if (aviso) zpl.append(campo(X + 1, y, fuente, maxLineas, texto));

            y += aviso ? ALTO_LINEA_AVISO : ALTO_LINEA;
        }
        return zpl.toString();
    }

    private static String campo(int x, int y, int fuente, int maxLineas, String texto) {
        return "^FO" + x + ',' + y
                + "^A0N," + fuente + ',' + fuente
                + "^FB" + ANCHO + ',' + maxLineas + ",0,L"
                + "^FD" + sanitizar(texto) + "^FS\n";
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

    private static String truncar(String texto) {
        if (texto.length() <= MAX_CARACTERES) return texto;
        return texto.substring(0, MAX_CARACTERES - 1) + "…";
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
