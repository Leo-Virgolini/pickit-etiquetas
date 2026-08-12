package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;

import java.util.ArrayList;
import java.util.List;

/**
 * Arma las líneas de embalaje que van en la etiqueta a partir de lo cargado en el Excel, y su
 * fragmento ZPL. Solo aparecen las líneas cuyos datos estén cargados: un SKU puede llevar desde
 * ninguna hasta cuatro (caja o bolsa, pluribol, rollo y observaciones).
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
    /** Cuántas líneas entran antes del separador de la zona de picking. */
    private static final int MAX_LINEAS = 6;
    private static final int ALTO_LINEA = 24;
    private static final int FUENTE = 22;
    private static final int ANCHO = 380;

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

        if (cargado(datos.pluribol())) {
            String linea = "PLURIBOL: " + valor(datos.pluribol());
            if (cargado(datos.cantPluribol())) linea += " - " + valor(datos.cantPluribol()) + " vueltas";
            lineas.add(linea);
        }

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
            // La última línea se queda con todo el alto libre que sobra. Es la de observaciones,
            // el único texto que puede no entrar en una línea: con ^FB de una sola línea ZPL no
            // descarta el sobrante, lo reimprime encima de la misma y queda ilegible.
            boolean ultima = i == lineas.size() - 1;
            int maxLineas = ultima ? Math.max(1, MAX_LINEAS - i) : 1;

            zpl.append("^FO").append(X).append(',').append(y)
                    .append("^A0N,").append(FUENTE).append(',').append(FUENTE)
                    .append("^FB").append(ANCHO).append(',').append(maxLineas).append(",0,L")
                    .append("^FD").append(sanitizar(lineas.get(i))).append("^FS\n");
            y += ALTO_LINEA;
        }
        return zpl.toString();
    }

    /**
     * ^ y ~ son los prefijos de comando de ZPL: dentro de un ^FD cortarían el campo y el resto de
     * la etiqueta se interpretaría como comandos. Los valores los tipea el usuario en el Excel,
     * así que se neutralizan antes de inyectarlos.
     */
    private static String sanitizar(String texto) {
        return texto.replace('^', ' ').replace('~', ' ');
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
