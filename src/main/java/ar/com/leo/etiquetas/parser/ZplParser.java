package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.ZplLabel;
import ar.com.leo.util.ZplHexDecoder;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ZplParser {

    private static final Pattern LABEL_BLOCK = Pattern.compile("(\\^XA.*?\\^XZ)", Pattern.DOTALL);
    private static final Pattern SEPARATOR = Pattern.compile("^\\^XA\\s*\\^MCY\\s*\\^XZ$", Pattern.DOTALL);
    private static final Pattern FD_FIELD = Pattern.compile("\\^FD(.*?)\\^FS", Pattern.DOTALL);
    private static final Pattern SKU_PATTERN = Pattern.compile("SKU:\\s*(\\S+)");
    private static final Pattern QUANTITY_PATTERN = Pattern.compile(
            "\\^A0N,70,70\\^FB160,1,0,C(?:\\^FR)?\\^FD(\\d+)\\^FS");
    /**
     * El rótulo con el que ML anuncia el número de la orden: "Pack ID:", o "Venta ID:" cuando la
     * etiqueta no agrupa varias.
     */
    private static final Pattern ROTULO_ORDEN = Pattern.compile("(?:Pack|Venta)\\s*ID:");
    /** El número en sí. Los pack y order id de ML son de 11 dígitos. */
    private static final Pattern NUMERO_ORDEN = Pattern.compile("\\d{8,}");
    /**
     * Cuántos campos después del rótulo se busca el número.
     *
     * Acotarlo es lo que evita quedarse con un número equivocado. En la etiqueta real el rótulo
     * viene duplicado —ML simula la negrita— y entre él y el número hay un campo con un "20000",
     * así que el número queda a dos o tres campos. Sin tope, la búsqueda seguiría hasta el
     * contenido del código de barras, que también es una tirada larga de dígitos: la columna
     * mostraría un número plausible pero equivocado, y el operario la usa para buscar la venta.
     */
    private static final int MAX_CAMPOS_ORDEN = 3;
    /**
     * El bloque de tipo de envío que imprime ML: "Envío Turbo" frente a "Envío Flex".
     *
     * En la descarga por API el dato sale de los tags del shipment, pero por archivo no hay de
     * dónde sacarlo salvo la etiqueta misma, y sin él una etiqueta turbo no se agrupa en TURBOS.
     * Se ancla en el "Envío" para no confundirlo con el "ZONA: TURBOS" que la app inyecta y que un
     * archivo ya procesado trae adentro.
     */
    private static final Pattern ENVIO_TURBO = Pattern.compile("Env\\S*o\\s+Turbo", Pattern.CASE_INSENSITIVE);

    private static final Pattern NON_DIGIT_START = Pattern.compile("^\\D+");
    private static final Pattern NON_DIGIT_END = Pattern.compile("\\D+$");

    /**
     * Normaliza un SKU:
     * 1. Trim de espacios
     * 2. Toma el texto antes del primer espacio
     * 3. Quita caracteres no numéricos al inicio/final
     * 4. Valida que sea numérico, si no marca "SKU INVALIDO: ..."
     */
    public static String normalizeSku(String raw) {
        if (raw == null) return null;

        // 1. Trim
        String sku = raw.trim();
        if (sku.isEmpty()) return null;

        // 2. Tomar antes del primer espacio
        int spaceIdx = sku.indexOf(' ');
        if (spaceIdx > 0) {
            sku = sku.substring(0, spaceIdx);
        }

        // 3. Quitar caracteres no numéricos al inicio y final
        sku = NON_DIGIT_START.matcher(sku).replaceFirst("");
        sku = NON_DIGIT_END.matcher(sku).replaceFirst("");

        if (sku.isEmpty()) {
            return "SKU INVALIDO: " + raw.trim();
        }

        // 4. Validar que sea numérico
        if (!sku.matches("\\d+")) {
            return "SKU INVALIDO: " + raw.trim();
        }

        return sku;
    }

    /**
     * La tirada larga de dígitos de un campo, o vacío si no la tiene.
     *
     * El campo tiene que ser el número y nada más: los códigos de barras llevan su contenido en un
     * ^FD y también traen tiradas largas de dígitos, pero mezcladas con otros caracteres.
     */
    private static String numeroDeOrden(String fieldContent) {
        String texto = fieldContent.trim();
        return NUMERO_ORDEN.matcher(texto).matches() ? texto : "";
    }

    public List<ZplLabel> parseFile(Path filePath) throws IOException {
        String content = Files.readString(filePath, StandardCharsets.UTF_8);
        return parse(content);
    }

    public List<ZplLabel> parse(String zplContent) {
        List<ZplLabel> labels = new ArrayList<>();
        Matcher blockMatcher = LABEL_BLOCK.matcher(zplContent);

        while (blockMatcher.find()) {
            String block = blockMatcher.group(1);

            if (SEPARATOR.matcher(block.trim()).matches()) {
                continue;
            }

            // Quitar ^MCY standalone dentro del bloque (algunos labels de ML lo traen,
            // otros no — normalizamos para evitar comportamiento inconsistente en la impresora)
            block = block.replaceAll("(?m)^\\^MCY[ \\t]*\\r?\\n", "");

            String decoded = ZplHexDecoder.decode(block);
            List<String> skus = new ArrayList<>();
            List<String> descriptions = new ArrayList<>();
            List<String> detailsList = new ArrayList<>();
            String previousField = null;

            String orderIds = "";
            int camposDesdeRotulo = -1;

            Matcher fdMatcher = FD_FIELD.matcher(decoded);
            while (fdMatcher.find()) {
                String fieldContent = fdMatcher.group(1);

                if (orderIds.isEmpty()) {
                    Matcher rotulo = ROTULO_ORDEN.matcher(fieldContent);
                    if (rotulo.find()) {
                        // El número puede venir pegado al rótulo, en el mismo campo.
                        camposDesdeRotulo = 0;
                        orderIds = numeroDeOrden(fieldContent.substring(rotulo.end()));
                    } else if (camposDesdeRotulo >= 0 && camposDesdeRotulo < MAX_CAMPOS_ORDEN) {
                        camposDesdeRotulo++;
                        orderIds = numeroDeOrden(fieldContent);
                    }
                }

                Matcher skuMatcher = SKU_PATTERN.matcher(fieldContent);
                if (skuMatcher.find()) {
                    skus.add(normalizeSku(skuMatcher.group(1)));
                    if (previousField != null && !previousField.isEmpty()) {
                        descriptions.add(previousField);
                    }
                    String beforeSku = fieldContent.substring(0, skuMatcher.start()).trim();
                    if (beforeSku.endsWith("|")) {
                        beforeSku = beforeSku.substring(0, beforeSku.length() - 1).trim();
                    }
                    if (!beforeSku.isEmpty()) {
                        detailsList.add(beforeSku);
                    }
                }
                previousField = fieldContent.trim();
            }

            String sku = skus.isEmpty() ? null : String.join("\n", skus);
            String description = descriptions.isEmpty() ? null : String.join("\n", descriptions);
            String details = detailsList.isEmpty() ? null : String.join("\n", detailsList);

            int quantity = 1;
            Matcher qtyMatcher = QUANTITY_PATTERN.matcher(decoded);
            if (qtyMatcher.find()) {
                quantity = Integer.parseInt(qtyMatcher.group(1));
            }

            boolean turbo = ENVIO_TURBO.matcher(decoded).find();

            labels.add(new ZplLabel(block, sku, description, details, quantity, turbo, orderIds));
        }

        return labels;
    }
}
