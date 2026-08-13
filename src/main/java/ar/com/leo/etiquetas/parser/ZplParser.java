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
     * El número que ML imprime junto a "Pack ID:" —o "Venta ID:" cuando la etiqueta no agrupa
     * varias órdenes—. Es el mismo dato que la descarga por API pone en la columna Orden, así que
     * la tabla dice lo mismo venga la etiqueta de la API o de un archivo.
     *
     * Entre el rótulo y el número hay otro campo con un "20000", así que no alcanza con tomar el
     * campo siguiente: se busca la primera tirada larga de dígitos.
     */
    private static final Pattern ORDEN_PATTERN = Pattern.compile(
            "(?:Pack|Venta)\\s*ID:.*?\\^FD\\s*(\\d{8,})\\s*\\^FS", Pattern.DOTALL);

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

            Matcher fdMatcher = FD_FIELD.matcher(decoded);
            while (fdMatcher.find()) {
                String fieldContent = fdMatcher.group(1);
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

            String orderIds = "";
            Matcher ordenMatcher = ORDEN_PATTERN.matcher(decoded);
            if (ordenMatcher.find()) {
                orderIds = ordenMatcher.group(1);
            }

            // turbo queda en false: es un dato del envío que solo conoce la API.
            labels.add(new ZplLabel(block, sku, description, details, quantity, false, orderIds));
        }

        return labels;
    }
}
