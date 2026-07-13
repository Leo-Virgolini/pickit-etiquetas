package ar.com.leo.api.dux;

import ar.com.leo.AppLogger;
import ar.com.leo.api.HttpRetryHandler;
import ar.com.leo.api.dux.model.DuxResponse;
import ar.com.leo.api.dux.model.Item;
import ar.com.leo.api.dux.model.Stock;
import ar.com.leo.api.dux.model.TokensDux;
import tools.jackson.databind.ObjectMapper;
import tools.jackson.databind.json.JsonMapper;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Supplier;

import static ar.com.leo.api.HttpRetryHandler.BASE_SECRET_DIR;

public class DuxApi {

    private static final ObjectMapper mapper = JsonMapper.shared();
    private static final HttpClient httpClient = HttpClient.newHttpClient();
    private static final HttpRetryHandler retryHandler = new HttpRetryHandler(httpClient, 7000L, 0.143);
    private static final Path TOKEN_FILE = BASE_SECRET_DIR.resolve("dux_tokens.json");
    private static TokensDux tokens;

    public static List<Item> obtenerProductos(String fecha) throws IOException {
        List<Item> allItems = new ArrayList<>();
        int offset = 0;
        int total = Integer.MAX_VALUE;
        final int limit = 50;
        int intentosVacios = 0;
        final int MAX_INTENTOS_VACIOS = 3;

        final boolean conFiltroFecha = fecha != null && !fecha.isBlank();
        final String fechaParam = conFiltroFecha
                ? "&fecha=" + URLEncoder.encode(fecha, StandardCharsets.UTF_8)
                : "";
        final String tipoProductos = conFiltroFecha ? "productos modificados (por stock o precio)" : "productos";

        while (offset < total) {
            final int finalOffset = offset;
            Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                    .uri(URI.create("https://erp.duxsoftware.com.ar/WSERP/rest/services/items?offset=" + finalOffset
                            + "&limit=" + limit + fechaParam))
                    .GET()
                    .header("accept", "application/json")
                    .header("authorization", tokens.token)
                    .build();

            HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

            if (response == null || response.statusCode() != 200) {
                String body = response != null ? response.body() : "sin respuesta";
                throw new IOException("Error al obtener productos: " + body);
            }

            String body = response.body().trim();
            DuxResponse duxResponse = mapper.readValue(body, DuxResponse.class);

            if (duxResponse.getPaging() != null) {
                int nuevoTotal = duxResponse.getPaging().getTotal();
                if (total == Integer.MAX_VALUE || nuevoTotal != total) {
                    total = nuevoTotal;
                    AppLogger.info("DUX - Total de " + tipoProductos + " en DUX: " + total);
                }
            }

            if (duxResponse.getResults() == null || duxResponse.getResults().isEmpty()) {
                if (offset >= total) {
                    break;
                }
                intentosVacios++;
                AppLogger.warn("DUX - Respuesta vacía en offset " + offset + " (intento " + intentosVacios + "/" + MAX_INTENTOS_VACIOS + ")");
                if (intentosVacios >= MAX_INTENTOS_VACIOS) {
                    AppLogger.warn("DUX - Terminando después de " + MAX_INTENTOS_VACIOS + " intentos vacíos. Productos obtenidos: " + allItems.size() + " / " + total);
                    break;
                }
                offset += limit;
                continue;
            }

            intentosVacios = 0;
            allItems.addAll(duxResponse.getResults());
            AppLogger.info(String.format("DUX - Obtenidos: %d / %d", allItems.size(), total));
            offset += limit;

            if (allItems.size() >= total) {
                break;
            }
        }

        AppLogger.info("DUX - Descarga completa: " + allItems.size() + " " + tipoProductos);
        return allItems;
    }

    public static Item obtenerProductoPorCodigo(String codigoItem) throws IOException {
        String codEncoded = URLEncoder.encode(codigoItem, StandardCharsets.UTF_8);

        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://erp.duxsoftware.com.ar/WSERP/rest/services/items?codigoItem=" + codEncoded + "&limit=1"))
                .GET()
                .header("accept", "application/json")
                .header("authorization", tokens.token)
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response == null || response.statusCode() != 200) {
            String body = response != null ? response.body() : "sin respuesta";
            AppLogger.warn("DUX - Error al obtener producto " + codigoItem + ": " + body);
            return null;
        }

        DuxResponse duxResponse = mapper.readValue(response.body().trim(), DuxResponse.class);

        if (duxResponse.getResults() != null && !duxResponse.getResults().isEmpty()) {
            return duxResponse.getResults().getFirst();
        }
        return null;
    }

    public static int obtenerStockPorCodigo(String codigoItem) {
        try {
            Item item = obtenerProductoPorCodigo(codigoItem);
            if (item == null || item.getStock() == null || item.getStock().isEmpty()) {
                return -1;
            }
            int totalStock = 0;
            for (Stock stock : item.getStock()) {
                String stockDisp = stock.getStockDisponible();
                if (stockDisp != null && !stockDisp.isBlank()) {
                    try {
                        totalStock += (int) Double.parseDouble(stockDisp);
                    } catch (NumberFormatException ignored) {
                    }
                }
            }
            return totalStock;
        } catch (Exception e) {
            AppLogger.warn("DUX - Error al obtener stock de " + codigoItem + ": " + e.getMessage());
            return -1;
        }
    }

    public static boolean inicializar() {
        tokens = cargarTokens();
        return tokens != null;
    }

    /**
     * Prueba manual: trae un producto de DUX e imprime el JSON crudo completo
     * para inspeccionar TODOS los campos que devuelve la API (no solo los que
     * mapea el modelo Item).
     *
     * <p>Uso:
     * <pre>
     *   (sin args)          → trae el primer producto del catálogo (offset=0, limit=1)
     *   &lt;codigoItem&gt;  → trae ese producto puntual por código
     * </pre>
     */
    public static void main(String[] args) throws Exception {
        if (!inicializar()) {
            System.err.println("No se pudieron cargar los tokens de DUX desde " + TOKEN_FILE);
            return;
        }

        String url;
        String sku = "1011021";

        String cod = URLEncoder.encode(sku.trim(), StandardCharsets.UTF_8);
        url = "https://erp.duxsoftware.com.ar/WSERP/rest/services/items?codigoItem=" + cod + "&limit=1";
        System.out.println(">> Trayendo producto por código: " + sku.trim());

        System.out.println(">> URL: " + url);

        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .GET()
                .header("accept", "application/json")
                .header("authorization", tokens.token)
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response == null) {
            System.err.println("Sin respuesta del servidor.");
            return;
        }
        System.out.println(">> HTTP status: " + response.statusCode());

        String body = response.body();
        // Intentar prettificar; si falla, imprimir crudo.
        try {
            Object json = mapper.readValue(body, Object.class);
            System.out.println(mapper.writerWithDefaultPrettyPrinter().writeValueAsString(json));
        } catch (Exception e) {
            System.out.println(body);
        }
    }

    private static TokensDux cargarTokens() {
        try {
            File f = TOKEN_FILE.toFile();
            if (!f.exists()) return null;
            return mapper.readValue(f, TokensDux.class);
        } catch (Exception e) {
            AppLogger.warn("DUX - Error al cargar tokens: " + e.getMessage());
            return null;
        }
    }
}
