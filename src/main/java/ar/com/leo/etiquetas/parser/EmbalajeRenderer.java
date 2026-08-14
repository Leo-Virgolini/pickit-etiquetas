package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;

import java.util.ArrayList;
import java.util.List;

/**
 * Arma las líneas de embalaje que van en la etiqueta a partir de lo cargado en el Excel, y su
 * fragmento ZPL. Un SKU lleva entre una y tres líneas —envase, rollo y observaciones—, o una sola
 * con un aviso en recuadro: cuando su columna ESTANDARIZADO dice que el embalaje todavía no está
 * resuelto, o cuando dice que sí pero no hay ninguna de las tres columnas cargada.
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
    /**
     * Dónde arranca la primera línea. No es el punto más alto que pinta el bloque: el encabezado
     * de referencia sube {@link #SUBIDA_REFERENCIA} por encima, y ese es el techo real
     * ({@link #Y_TECHO}).
     */
    private static final int Y_INICIAL = 20;
    /**
     * Lo más arriba que llega cualquier cosa del bloque.
     *
     * Vale la pena tenerlo escrito porque el margen de arriba no es infinito: el "#N" de la
     * etiqueta usa y=30 justamente para no ser cortado por el borde superior. Está cubierto por un
     * test, así que mover Y_INICIAL o SUBIDA_REFERENCIA sin pensar en el borde falla.
     */
    private static final int Y_TECHO = 14;
    /** Hasta dónde puede llegar el bloque: el separador de la zona de picking está en y=180. */
    private static final int Y_LIMITE = 178;
    private static final int FUENTE = 26;
    /**
     * Paso entre una línea y la siguiente. Es el alto de la fuente y no uno mayor: dentro de un
     * ^FB las filas de una misma línea ya se separan así, que es el ritmo con el que se lee el
     * bloque. Dejar aire de más entre líneas se lo saca a la última, que es la que se reparte el
     * alto que sobra: con 28 las observaciones de una etiqueta de varias unidades entraban en dos
     * filas y la tercera no llegaba por dos puntos.
     */
    private static final int ALTO_LINEA = FUENTE;
    private static final int ANCHO = 390;
    /**
     * Caracteres que entran en una fila.
     *
     * Es un número plano para una fuente proporcional, así que es una aproximación: medido sobre
     * una impresión real, un dígito ocupa ~13,4 puntos —entran 29 en los 390— y una mayúscula ~15,
     * con lo que de esas entrarían 25. Se eligió el valor de los dígitos, que es el que aprovecha
     * el ancho. El costo es que una fila de puras mayúsculas puede pasarse, y ^FB no descarta lo
     * que sobra sino que lo reimprime encima de la fila.
     *
     * Las observaciones del Excel son en mayúscula, así que si alguna vez aparece una fila
     * ilegible, este es el número a bajar.
     */
    private static final int MAX_CARACTERES = 29;
    /**
     * Pedazo más chico en el que vale la pena partir una palabra.
     *
     * Marca el equilibrio entre dos formas de quedar mal. Si en la fila entra menos que esto, la
     * palabra arranca abajo: partir "AJUSTAR" en "AJUS TAR" por cuatro caracteres se lee peor que
     * dejar ese margen. Si entra más, se parte: bajarla entera dejaría media fila vacía, y en una
     * línea de una sola fila —envase, rollo— el recorte se quedaría con el rótulo solo y el valor
     * se perdería entero.
     */
    private static final int MIN_PIEZA = MAX_CARACTERES / 2;
    /** Marca de texto cortado. Tres puntos y no "…": la fuente residual puede no traer ese glifo. */
    private static final String ELIPSIS = "...";

    // El aviso de embalaje sin cargar no es un dato más de la lista: tiene que frenar al operario.
    // Va en un recuadro relleno con el texto en video inverso y centrado, y siempre es la única
    // línea del bloque. Con fuente 38 sus 16 caracteres en mayúscula ocupan ~365 de los 390 de
    // ancho; más grande arañaría el borde, y ^FB no descarta lo que no entra sino que lo reimprime
    // encima.
    private static final int FUENTE_AVISO = 38;
    private static final int ALTO_AVISO = 52;
    /**
     * Ancho de una mayúscula en la fuente escalable, en milésimos de su alto. Medido en impresión:
     * a cuerpo 26 ocupa unos 15 puntos.
     */
    private static final int ANCHO_MAYUSCULA = 577;

    /** Reemplaza a todo el bloque cuando la columna ESTANDARIZADO no dice que sí. */
    private static final String SIN_ESTANDARIZAR = "NO ESTANDARIZADO";
    /**
     * Reemplaza a todo el bloque cuando la fórmula dice que sí pero no hay ninguna de las tres
     * columnas de datos cargada. Es un problema distinto —ahí lo que falta es la carga, no la
     * decisión— pero para el operario da lo mismo: no tiene ninguna indicación de embalaje.
     */
    private static final String SIN_DATOS = "SIN DATOS DE EMBALAJE";
    /** Encabeza el bloque cuando la etiqueta es de más de una unidad. */
    private static final String REFERENCIA = "REFERENCIA";
    // El encabezado va en la fuente residual B, de trazo más cuadrado que el ^A0 del resto, para
    // que se lea como un rótulo y no como un dato más. Las fuentes bitmap de Zebra escalan solo en
    // múltiplos enteros de su base (11x7), así que 22x14 es el tamaño más cercano al cuerpo: dos
    // píxeles más bajo, con lo que entra en el mismo alto de línea sin correr nada.
    private static final char FUENTE_REFERENCIA = 'B';
    private static final int ALTO_REFERENCIA = 22;
    private static final int ANCHO_REFERENCIA = 14;
    /**
     * Espacio que la fuente B deja entre caracteres. La matriz 11x7 es el glifo solo: el hueco son
     * 2 dots aparte, escalados por la misma magnificación que el resto.
     *
     * Confirmado en impresión: sin contarlo, el subrayado llegaba hasta la octava letra, que es
     * justo donde termina el glifo número ocho sin sus huecos (8 · 14 + 7 · 4 = 140).
     */
    private static final int GAP_REFERENCIA = 4;
    /**
     * Cuánto sube el encabezado dentro de su propia línea.
     *
     * Mide 24 con el subrayado incluido, contra los 26 del paso del bloque, así que apoyado en el
     * arranque de su línea deja el subrayado a dos puntos de la primera línea de datos y se lee
     * como si fueran una sola cosa. Sube él y no baja el resto: el alto que se libere abajo es el
     * que se reparte la última línea.
     *
     * Lo que sube sale de {@link #Y_TECHO}: si alguna vez se imprime rozado por el borde de
     * arriba, ese es el número a subir.
     */
    private static final int SUBIDA_REFERENCIA = Y_INICIAL - Y_TECHO;

    private EmbalajeRenderer() {
    }

    /**
     * Líneas a imprimir para un SKU en una etiqueta de {@code cantidad} unidades.
     *
     * Con más de una unidad el operario no está embalando un producto suelto, así que el envase
     * cargado es orientativo: las líneas salen igual, pero encabezadas por "REFERENCIA". Es la
     * única diferencia. Los avisos no dependen de la cantidad —el embalaje está sin resolver
     * igual— y salen sin encabezado, porque no hay ningún dato que rotular.
     */
    public static List<String> lineas(DatosEmbalaje datos, int cantidad) {
        List<String> lineas = new ArrayList<>();
        // Sin la columna ESTANDARIZADO en el Excel la función no está en uso: la etiqueta sale como
        // antes de existir, en vez de reclamar algo que el usuario todavía no puede cargar.
        if (datos == null || !datos.aplica()) return lineas;

        // El aviso sale de la columna ESTANDARIZADO y no de si hay envase cargado: es una fórmula
        // del usuario que resume si completó envase, tipo de rollo y cantidad de paños. Cuando dice
        // que no, la etiqueta solo tiene que frenar al operario, sin ruido alrededor: ni siquiera
        // el encabezado de referencia, que rotula datos de embalaje y acá no hay ninguno.
        if (!datos.estandarizado()) {
            lineas.add(SIN_ESTANDARIZAR);
            return lineas;
        }

        boolean referencia = cantidad > 1;
        if (referencia) lineas.add(REFERENCIA);

        // "NO" en el envase es una decisión tomada —ese producto no lleva caja ni bolsa— así que se
        // imprime como cualquier otro valor.
        if (cargado(datos.envase())) {
            lineas.add("ENVASE: " + envase(datos));
        }

        if (cargado(datos.rollo())) {
            String linea = "ROLLO: " + valor(datos.rollo());
            linea += panos(datos.cantPanos());
            lineas.add(linea);
        }

        if (cargado(datos.observaciones())) lineas.add("OBS: " + valor(datos.observaciones()));

        // La fórmula del usuario puede decir que sí sin que haya ninguna de las tres columnas de
        // datos: son suyas y se ubican por encabezado, así que pueden faltar. Dejar el bloque en
        // blanco escondería que al Excel le falta la carga, y con más de una unidad además dejaría
        // "REFERENCIA" solo, que no dice nada.
        boolean sinDatos = lineas.isEmpty()
                || (lineas.size() == 1 && REFERENCIA.equals(lineas.getFirst()));
        if (sinDatos) return new ArrayList<>(List.of(SIN_DATOS));

        return lineas;
    }

    /**
     * Si esas líneas son el aviso de embalaje sin cargar. El diálogo del final del lote lista
     * exactamente los SKU que salieron avisados en papel, así que se pregunta por lo que se imprime
     * y no por lo que dice el Excel.
     */
    public static boolean avisaSinEstandarizar(List<String> lineas) {
        return esAviso(lineas, SIN_ESTANDARIZAR);
    }

    /**
     * Si esas líneas son el aviso de que faltan los datos de embalaje. Se lista aparte del anterior
     * porque lo que hay que arreglar en el Excel es otra cosa: ahí la fórmula ya dice que sí y lo
     * que falta es completar las columnas.
     */
    public static boolean avisaSinDatos(List<String> lineas) {
        return esAviso(lineas, SIN_DATOS);
    }

    private static boolean esAviso(List<String> lineas, String aviso) {
        return lineas != null && lineas.size() == 1 && aviso.equals(lineas.getFirst());
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

            if (SIN_ESTANDARIZAR.equals(linea) || SIN_DATOS.equals(linea)) {
                zpl.append(recuadro(y, linea));
                y += ALTO_AVISO;
                continue;
            }

            if (REFERENCIA.equals(linea)) {
                zpl.append(referencia(y));
                y += ALTO_LINEA;
                continue;
            }

            // Con este alto de línea entran seis, y lineas() nunca arma más de cuatro. El corte
            // es por si eso cambiara: perder una línea es preferible a imprimir sobre la zona de
            // picking, que empieza en y=180.
            if (y + FUENTE > Y_LIMITE) break;

            boolean ultima = i == lineas.size() - 1;
            // La última línea se queda con todo el alto libre que sobra, así que puede repartirse
            // en varias. Las anteriores no pueden crecer sin correr las de abajo, así que entran
            // en una sola — la inscripción del envase también es texto libre del usuario.
            //
            // La cuenta sale exacta porque ALTO_LINEA es el interlineado que ^FB usa entre las
            // filas: la última fila termina en y + maxLineas · ALTO_LINEA.
            int maxLineas = ultima ? Math.max(1, (Y_LIMITE - y) / ALTO_LINEA) : 1;
            // En los dos casos se acota: con ^FB, ZPL no descarta lo que no entra, lo reimprime
            // encima de la última fila y queda una mancha ilegible.
            String texto = acotar(linea, maxLineas);

            zpl.append(campo(X, y, maxLineas, texto));
            // Negrita simulada con una segunda pasada corrida 1px, igual que ZONA y COD.EXT. Se
            // repasa solo el rótulo, que se superpone exactamente sobre el de la primera pasada.
            // Así no hay que calcular su ancho, que con una fuente proporcional no se puede saber
            // de antemano.
            String negrita = rotulo(texto);
            if (!negrita.isEmpty()) zpl.append(campo(X + 1, y, maxLineas, negrita));

            y += ALTO_LINEA;
        }
        return zpl.toString();
    }

    /**
     * Encabezado de referencia: fuente bitmap, en negrita y subrayado, y unos puntos por encima
     * del arranque de su línea (ver {@link #SUBIDA_REFERENCIA}).
     *
     * ZPL no tiene subrayado como atributo, así que se dibuja. Con la fuente proporcional del resto
     * del bloque habría que estimar el ancho de la palabra; la bitmap es monoespaciada, así que sale
     * exacto contando glifos y huecos.
     */
    private static String referencia(int y) {
        int yTexto = y - SUBIDA_REFERENCIA;
        String campo = "^A" + FUENTE_REFERENCIA + "N," + ALTO_REFERENCIA + ',' + ANCHO_REFERENCIA
                + "^FD" + REFERENCIA + "^FS\n";
        // Los glifos más los huecos que quedan entre ellos, sin el que sobraría al final.
        int ancho = REFERENCIA.length() * ANCHO_REFERENCIA
                + (REFERENCIA.length() - 1) * GAP_REFERENCIA;
        return "^FO" + X + ',' + yTexto + campo
                + "^FO" + (X + 1) + ',' + yTexto + campo
                + "^FO" + X + ',' + (yTexto + ALTO_REFERENCIA) + "^GB" + ancho + ",2,2^FS\n";
    }

    /** Recuadro relleno con el aviso en video inverso, centrado vertical y horizontalmente. */
    private static String recuadro(int y, String texto) {
        int fuente = fuenteAviso(texto);
        return "^FO" + X + ',' + y + "^GB" + ANCHO + ',' + ALTO_AVISO + ',' + ALTO_AVISO + "^FS\n"
                + "^FO" + X + ',' + (y + (ALTO_AVISO - fuente) / 2)
                + "^A0N," + fuente + ',' + fuente
                + "^FB" + ANCHO + ",1,0,C^FR^FD" + sanitizar(texto) + "^FS\n";
    }

    /**
     * Cuerpo con el que el aviso entra entero en una fila, hasta el tamaño de referencia.
     *
     * El aviso es lo único del bloque que tiene que leerse de lejos, así que va lo más grande
     * posible; pero pasarse no es una opción, porque ^FB no descarta lo que no entra sino que lo
     * reimprime encima de la fila y queda una mancha. La cuenta toma todos los caracteres como
     * mayúsculas —los espacios son más angostos— así que se queda corta, que es el lado seguro.
     */
    private static int fuenteAviso(String texto) {
        int entero = ANCHO * 1000 / (texto.length() * ANCHO_MAYUSCULA);
        return Math.min(FUENTE_AVISO, entero);
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
     * Reparte el texto en las filas que el ^FB tiene permitido imprimir, y recorta lo que sobra.
     *
     * No alcanza con contar caracteres. ^FB corta por palabras, así que un texto de
     * {@code maxFilas · MAX_CARACTERES} caracteres puede necesitar una fila más: "OBS: 3 COLCHON 1
     * TAPA. AJUSTAR PAÑO SEGÚN CANT VENDIDA" mide 54 y entra en 2 x 27 por cuenta de caracteres,
     * pero al cortarse por palabras ocupa tres filas.
     *
     * Las filas se devuelven unidas por espacios, que es donde ^FB va a cortar. Reproduce el mismo
     * reparto: una fila solo se corta cuando la palabra que sigue no entraba, así que la impresora
     * tampoco la va a poder subir.
     */
    private static String acotar(String texto, int maxFilas) {
        List<String> filas = envolver(texto);

        if (filas.size() > maxFilas) {
            String ultima = filas.get(maxFilas - 1);
            if (ultima.length() + ELIPSIS.length() > MAX_CARACTERES) {
                ultima = ultima.substring(0, MAX_CARACTERES - ELIPSIS.length()).stripTrailing();
            }
            filas = new ArrayList<>(filas.subList(0, maxFilas - 1));
            filas.add(ultima + ELIPSIS);
        }

        return String.join(" ", filas);
    }

    /**
     * Cómo queda repartido el texto en filas del ancho disponible.
     *
     * Una palabra que no entra en lo que queda de la fila se parte ahí mismo y sigue abajo, salvo
     * que quede muy poco lugar (ver {@link #MIN_PIEZA}). Bajarla entera siempre desperdiciaría el
     * resto de la fila, y en las líneas de una sola fila dejaría el rótulo solo.
     */
    private static List<String> envolver(String texto) {
        List<String> filas = new ArrayList<>();
        StringBuilder fila = new StringBuilder();

        for (String palabra : texto.split(" ")) {
            while (!palabra.isEmpty()) {
                int libre = MAX_CARACTERES - fila.length() - (fila.isEmpty() ? 0 : 1);

                if (palabra.length() <= libre) {
                    if (!fila.isEmpty()) fila.append(' ');
                    fila.append(palabra);
                    palabra = "";
                } else if (libre >= MIN_PIEZA) {
                    if (!fila.isEmpty()) fila.append(' ');
                    fila.append(palabra, 0, libre);
                    palabra = palabra.substring(libre);
                    filas.add(fila.toString());
                    fila = new StringBuilder();
                } else {
                    // Queda tan poco que partirla no vale la pena: arranca abajo.
                    filas.add(fila.toString());
                    fila = new StringBuilder();
                }
            }
        }

        if (!fila.isEmpty()) filas.add(fila.toString());
        return filas;
    }

    /**
     * Código del envase con su inscripción entre comillas, del tipo {@code CAJ-1 "9Y"}. Las comillas
     * marcan que eso es lo que está escrito en el envase físico, que es como lo busca el operario.
     * Sin inscripción va el código solo, sin comillas vacías colgando.
     */
    private static String envase(DatosEmbalaje datos) {
        String codigo = valor(datos.envase());
        String inscripcion = inscripcion(datos.inscripcion());
        if (inscripcion.isEmpty()) return codigo;
        if (codigo.isEmpty()) return "\"" + inscripcion + "\"";
        return codigo + " \"" + inscripcion + "\"";
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
