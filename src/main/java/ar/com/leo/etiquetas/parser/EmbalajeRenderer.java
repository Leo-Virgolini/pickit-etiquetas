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

    // Bloque en el margen superior derecho: debajo del banner MEDIR (que termina en y=80) y por
    // encima del separador de la zona de picking (y=180). Cuatro líneas llegan hasta y≈169.
    private static final int X = 450;
    private static final int Y_INICIAL = 85;
    private static final int ALTO_LINEA = 22;
    private static final int FUENTE = 18;
    private static final int ANCHO = 340;

    private EmbalajeRenderer() {
    }

    public static List<String> lineas(DatosEmbalaje datos) {
        List<String> lineas = new ArrayList<>();
        if (datos == null) return lineas;

        // Caja y bolsa son excluyentes: si están las dos cargadas gana la caja. Puede pasar cuando
        // cambia el embalaje y queda el número de bolsa viejo; con las dos salen cinco líneas y la
        // última se imprimiría sobre el separador de la zona de picking.
        String caja = unir(valor(datos.nroCaja()), valor(datos.nombreCaja()));
        if (!caja.isEmpty()) lineas.add("CAJA " + caja);
        else if (cargado(datos.nroBolsa())) lineas.add("BOLSA " + valor(datos.nroBolsa()));

        if (cargado(datos.pluribol())) {
            String linea = "PLURIBOL: " + valor(datos.pluribol());
            if (cargado(datos.cantPluribol())) linea += " - " + valor(datos.cantPluribol()) + " vueltas";
            lineas.add(linea);
        }

        if (cargado(datos.rollo())) {
            String linea = "ROLLO " + valor(datos.rollo());
            if (cargado(datos.cantPanos())) linea += " - " + valor(datos.cantPanos()) + " paños";
            lineas.add(linea);
        }

        if (cargado(datos.observaciones())) lineas.add("Obs: " + valor(datos.observaciones()));

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
        for (String linea : lineas) {
            zpl.append("^FO").append(X).append(',').append(y)
                    .append("^A0N,").append(FUENTE).append(',').append(FUENTE)
                    .append("^FB").append(ANCHO).append(",1,0,L")
                    .append("^FD").append(sanitizar(linea)).append("^FS\n");
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
