package ar.com.leo.api.ml;

import ar.com.leo.AppLogger;
import ar.com.leo.api.HttpRetryHandler;
import ar.com.leo.api.ml.model.MLCredentials;
import ar.com.leo.api.ml.model.OrdenML;
import ar.com.leo.api.ml.model.TokensML;
import ar.com.leo.api.ml.model.Venta;
import ar.com.leo.etiquetas.model.ZplLabel;
import ar.com.leo.etiquetas.parser.ZplParser;
import ar.com.leo.pedidos.model.PedidoML;
import javafx.application.Platform;
import javafx.scene.control.TextInputDialog;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.ObjectMapper;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.ObjectNode;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static ar.com.leo.api.HttpRetryHandler.BASE_SECRET_DIR;
import static ar.com.leo.etiquetas.parser.ZplParser.normalizeSku;

public class MercadoLibreAPI {

    public record MLOrderResult(List<Venta> ventas, List<OrdenML> ordenes) {
    }

    public record SlaInfo(String status, OffsetDateTime expectedDate, boolean turbo) {
    }

    public record ShipmentInfo(String substatus, OffsetDateTime slaDate) {
    }

    private static final int MAX_SHIPMENTS_PER_REQUEST = 50;
    private static final Path MERCADOLIBRE_FILE = BASE_SECRET_DIR.resolve("ml_credentials.json");
    private static final Path TOKEN_FILE = BASE_SECRET_DIR.resolve("ml_tokens.json");
    private static final Object TOKEN_LOCK = new Object();
    private static final ObjectMapper mapper = JsonMapper.shared();
    private static final HttpClient httpClient = HttpClient.newHttpClient();
    // Rate limit de ML: 18000 req/hora → 300 req/min → 5 req/seg
    // Al ser un programa que usa pocas requests se usa un valor más alto (25 req/seg)
    private static final HttpRetryHandler retryHandler = new HttpRetryHandler(httpClient, 30000L, 25, MercadoLibreAPI::verificarTokens);

    public static HttpRetryHandler getRetryHandler() {
        return retryHandler;
    }

    private static final ExecutorService executor = Executors.newFixedThreadPool(25);
    private static final ZplParser zplParser = new ZplParser();

    // Patrones para identificar a qué envío pertenece cada etiqueta mirando su contenido ZPL,
    // en vez de asumir el orden de la respuesta de ML (que no está garantizado).
    private static final Pattern ZPL_ENVIO_PATTERN = Pattern.compile("Envio:\\s*(\\d+)");
    private static final Pattern ZPL_QR_ID_PATTERN = Pattern.compile("\"id\"\\s*:\\s*\"(\\d+)\"");
    private static MLCredentials mlCredentials;
    private static volatile TokensML tokens;

    public static String getAccessToken() {
        verificarTokens();
        return tokens != null ? tokens.accessToken : null;
    }

    public static void shutdownExecutors() {
        executor.shutdown();
        try {
            if (!executor.awaitTermination(5, TimeUnit.SECONDS)) {
                executor.shutdownNow();
                if (!executor.awaitTermination(2, TimeUnit.SECONDS)) {
                    AppLogger.warn("El executor no terminó correctamente después de shutdownNow.");
                }
            }
        } catch (InterruptedException e) {
            executor.shutdownNow();
            Thread.currentThread().interrupt();
        }
    }

    public static String getUserId() throws IOException {
        MercadoLibreAPI.verificarTokens();
        final String url = "https://api.mercadolibre.com/users/me";

        final Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response.statusCode() != 200) {
            throw new IOException("Error al obtener el user ID de ML: " + response.body());
        }

        return mapper.readTree(response.body()).get("id").asString();
    }

    private static final String SUBSTATUS_PENDIENTES = "ready_to_print";
    private static final String SUBSTATUS_IMPRESAS = "printed,ready_for_dropoff,ready_for_pickup";

    /**
     * Obtiene las ventas de ML con etiqueta lista para imprimir.
     * Hace búsquedas separadas por grupo de substatus para poder asignar el substatus
     * a cada orden sin necesidad de consultar /shipments/{id} individualmente.
     *
     * @param incluirImpresas si es true, incluye también las que ya fueron impresas/despachadas por el vendedor
     */
    public static MLOrderResult obtenerVentasReadyToPrint(String userId, boolean incluirImpresas) {
        verificarTokens();

        List<Venta> ventas = new ArrayList<>();
        List<OrdenML> ordenes = new ArrayList<>();
        Set<Long> orderIdsSeen = new HashSet<>();

        // Siempre buscar pendientes
        searchAndCollect(userId, SUBSTATUS_PENDIENTES, "ready_to_print", orderIdsSeen, ventas, ordenes);

        // Si incluir impresas, hacer segunda búsqueda
        if (incluirImpresas) {
            searchAndCollect(userId, SUBSTATUS_IMPRESAS, "printed", orderIdsSeen, ventas, ordenes);
        }
        return new MLOrderResult(ventas, ordenes);
    }

    /**
     * Busca órdenes con un substatus dado, pagina automáticamente, y agrega los resultados.
     * A cada OrdenML le asigna el substatusTag como shippingSubstatus.
     */
    private static void searchAndCollect(String userId, String substatus, String substatusTag,
                                         Set<Long> orderIdsSeen, List<Venta> ventas, List<OrdenML> ordenes) {
        final int limit = 50;

        String firstUrl = buildOrderSearchUrl(userId, substatus, 0, limit);
        String firstBody = fetchOrderSearchPage(firstUrl);
        if (firstBody == null) return;

        JsonNode firstRoot = mapper.readTree(firstBody);
        int total = firstRoot.path("paging").path("total").asInt(0);

        List<String> allBodies = new ArrayList<>();
        allBodies.add(firstBody);

        if (total > limit) {
            List<CompletableFuture<String>> futures = new ArrayList<>();
            for (int off = limit; off < total; off += limit) {
                String pageUrl = buildOrderSearchUrl(userId, substatus, off, limit);
                futures.add(CompletableFuture.supplyAsync(() -> fetchOrderSearchPage(pageUrl), executor));
            }
            for (var future : futures) {
                String body = future.join();
                if (body != null) allBodies.add(body);
            }
        }

        for (String body : allBodies) {
            JsonNode root = mapper.readTree(body);
            JsonNode results = root.path("results");
            if (!results.isArray()) continue;

            for (JsonNode order : results) {
                OrdenML ordenML = parseOrder(order, orderIdsSeen, ventas);
                if (ordenML != null) {
                    ordenML.setShippingSubstatus(substatusTag);
                    ordenes.add(ordenML);
                }
            }
        }
    }

    private static String buildOrderSearchUrl(String userId, String substatus, int offset, int limit) {
        return String.format(
                "https://api.mercadolibre.com/orders/search?seller=%s&shipping.status=ready_to_ship&shipping.substatus=%s&sort=date_asc&offset=%d&limit=%d",
                userId, substatus, offset, limit);
    }


    private static String fetchOrderSearchPage(String url) {
        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response == null || response.statusCode() != 200) {
            String body = response != null ? response.body() : "sin respuesta";
            AppLogger.warn("ML - Error al obtener órdenes: " + body);
            return null;
        }
        return response.body();
    }

    private static OrdenML parseOrder(JsonNode order, Set<Long> orderIdsSeen, List<Venta> ventas) {
        long orderId = order.path("id").asLong();
        if (!orderIdsSeen.add(orderId)) return null;

        // Excluir órdenes con tag "delivered"
        JsonNode tagsNode = order.path("tags");
        if (tagsNode.isArray()) {
            for (JsonNode tag : tagsNode) {
                if ("delivered".equals(tag.asString())) return null;
            }
        }

        String dateCreated = order.path("date_created").asString("");
        OffsetDateTime fecha = null;
        if (!dateCreated.isBlank()) {
            try {
                fecha = OffsetDateTime.parse(dateCreated);
            } catch (Exception e) {
                AppLogger.warn("ML - Error al parsear fecha de orden " + orderId + ": " + dateCreated);
            }
        }
        JsonNode packNode = order.path("pack_id");
        Long packId = packNode.isNull() || packNode.isMissingNode() ? null : packNode.asLong();
        JsonNode shippingObj = order.path("shipping");
        JsonNode shippingIdNode = shippingObj.path("id");
        Long shipmentId = shippingIdNode.isNull() || shippingIdNode.isMissingNode() ? null : shippingIdNode.asLong();
        OrdenML ordenML = new OrdenML(orderId, packId, shipmentId, fecha, "");

        JsonNode orderItems = order.path("order_items");
        if (!orderItems.isArray()) return null;

        for (JsonNode orderItem : orderItems) {
            JsonNode item = orderItem.path("item");
            String rawSku = item.path("seller_sku").asString("");
            if (rawSku.isBlank()) {
                rawSku = item.path("seller_custom_field").asString("");
            }
            String sku = rawSku.isBlank() ? "" : normalizeSku(rawSku);
            if (sku == null) sku = "";
            String itemTitle = item.path("title").asString("");
            double quantity = orderItem.path("quantity").asDouble(0);

            if (quantity <= 0) {
                AppLogger.warn("ML - Producto con cantidad inválida en orden " + orderId + ": " + sku);
                String errorSku = sku.isBlank() ? itemTitle : sku;
                Venta venta = new Venta("CANT INVALIDA: " + errorSku, quantity, "ML", itemTitle);
                ventas.add(venta);
                ordenML.getItems().add(venta);
                continue;
            }
            if (sku.isBlank()) {
                AppLogger.warn("ML - Producto sin SKU en orden " + orderId + ": " + itemTitle);
                Venta venta = new Venta("SIN SKU: " + itemTitle, quantity, "ML", itemTitle);
                ventas.add(venta);
                ordenML.getItems().add(venta);
                continue;
            }
            Venta venta = new Venta(sku, quantity, "ML", itemTitle);
            ventas.add(venta);
            ordenML.getItems().add(venta);
        }

        return ordenML.getItems().isEmpty() ? null : ordenML;
    }

    // -----------------------------------------------------------------------------------------------------------------
    // DESCARGA DE ETIQUETAS ZPL
    // -----------------------------------------------------------------------------------------------------------------

    /**
     * Descarga etiquetas ZPL para las órdenes ready_to_ship y las devuelve como List&lt;ZplLabel&gt;
     * enriquecidas con SKU y descripción de las órdenes.
     * <p><b>IMPORTANTE:</b> La API de MercadoLibre cambia automáticamente el substatus de los envíos
     * de {@code ready_to_print} a {@code printed} al descargar las etiquetas vía
     * {@code GET /shipment_labels}. Este es un efecto colateral del endpoint de ML,
     * no una acción explícita de esta aplicación.</p>
     *
     * @param soloSlaHoy si es true, filtra ready_to_print con SLA &lt;= hoy y printed con SLA &gt;= hoy
     */
    public static List<ZplLabel> descargarEtiquetasZpl(String userId, boolean incluirImpresas, boolean soloSlaHoy) {
        MLOrderResult result = obtenerVentasReadyToPrint(userId, incluirImpresas);
        List<OrdenML> ordenes = result.ordenes();

        if (ordenes.isEmpty()) {
            return List.of();
        }

        // Construir mapa shipmentId → (sku, descripción) desde las órdenes
        Map<Long, SkuInfo> shipmentSkuMap = new LinkedHashMap<>();
        List<Long> shipmentIds = new ArrayList<>();

        for (OrdenML orden : ordenes) {
            Long shipId = orden.getShipmentId();
            if (shipId == null || shipId <= 0) continue;

            String sku = "";
            String desc = "";
            int totalQty = 0;
            if (!orden.getItems().isEmpty()) {
                Venta primerItem = orden.getItems().getFirst();
                sku = primerItem.getSku();
                desc = primerItem.getTitulo();
            }
            for (Venta v : orden.getItems()) {
                totalQty += (int) v.getCantidad();
            }

            if (!shipmentSkuMap.containsKey(shipId)) {
                shipmentSkuMap.put(shipId, new SkuInfo(sku, desc, Math.max(totalQty, 1), String.valueOf(orden.getOrderId())));
                shipmentIds.add(shipId);
            }
        }

        // Filtrar por SLA si corresponde
        if (soloSlaHoy) {
            Map<Long, SlaInfo> slaMap = obtenerSlasParalelo(shipmentIds);

            OffsetDateTime hoyFin = java.time.LocalDate.now()
                    .atTime(23, 59, 59).atZone(java.time.ZoneId.systemDefault()).toOffsetDateTime();

            List<Long> filtrados = new ArrayList<>();
            for (Long shipId : shipmentIds) {
                SlaInfo sla = slaMap.get(shipId);
                if (sla == null || sla.expectedDate() == null) {
                    filtrados.add(shipId); // sin SLA, incluir por las dudas
                    continue;
                }
                OffsetDateTime expected = sla.expectedDate();
                if (expected.isBefore(hoyFin) || expected.isEqual(hoyFin)) {
                    filtrados.add(shipId); // SLA hoy o antes
                }
            }

            shipmentIds = filtrados;
        }

        if (shipmentIds.isEmpty()) {
            return List.of();
        }

        // Descargar en batches de 50
        List<ZplLabel> allLabels = new ArrayList<>();
        for (int i = 0; i < shipmentIds.size(); i += MAX_SHIPMENTS_PER_REQUEST) {
            List<Long> batch = shipmentIds.subList(i, Math.min(i + MAX_SHIPMENTS_PER_REQUEST, shipmentIds.size()));
            List<ZplLabel> batchLabels = descargarBatchZpl(batch, shipmentSkuMap, Set.of());
            allLabels.addAll(batchLabels);
        }

        return allLabels;
    }

    /**
     * Descarga etiquetas ZPL solo para las órdenes dadas (usadas en el paso 2 del flujo).
     */
    public static List<ZplLabel> descargarEtiquetasZplParaOrdenes(List<OrdenML> ordenes, Set<Long> turboShipmentIds) {
        // Agrupar órdenes por shipmentId para detectar carros (múltiples items/órdenes por envío)
        Map<Long, List<OrdenML>> ordenesPorShipment = new LinkedHashMap<>();
        for (OrdenML orden : ordenes) {
            Long shipId = orden.getShipmentId();
            if (shipId == null || shipId <= 0) continue;
            ordenesPorShipment.computeIfAbsent(shipId, k -> new ArrayList<>()).add(orden);
        }

        Map<Long, SkuInfo> shipmentSkuMap = new LinkedHashMap<>();
        List<Long> shipmentIds = new ArrayList<>();

        for (var entry : ordenesPorShipment.entrySet()) {
            Long shipId = entry.getKey();
            List<OrdenML> group = entry.getValue();

            // Recolectar todos los SKUs y títulos de todas las órdenes del envío
            StringJoiner skuJoiner = new StringJoiner("\n");
            StringJoiner titleJoiner = new StringJoiner("\n");
            int totalQty = 0;

            for (OrdenML o : group) {
                for (Venta v : o.getItems()) {
                    String s = v.getSku() != null ? v.getSku() : "";
                    if (!s.isEmpty()) skuJoiner.add(s);
                    titleJoiner.add(v.getTitulo() != null && !v.getTitulo().isEmpty() ? v.getTitulo() : s);
                    totalQty += (int) v.getCantidad();
                }
            }

            String sku = skuJoiner.toString();
            String title = titleJoiner.toString();
            OrdenML first = group.getFirst();
            String orderIdsStr = first.getPackId() != null
                    ? String.valueOf(first.getPackId())
                    : String.valueOf(first.getOrderId());

            shipmentSkuMap.put(shipId, new SkuInfo(sku, title, Math.max(totalQty, 1), orderIdsStr));
            shipmentIds.add(shipId);
        }

        if (shipmentIds.isEmpty()) {
            AppLogger.info("ML - No hay envíos para descargar etiquetas.");
            return List.of();
        }

        AppLogger.info("ML - Descargando etiquetas ZPL para " + shipmentIds.size() + " envíos seleccionados...");

        List<ZplLabel> allLabels = new ArrayList<>();
        for (int i = 0; i < shipmentIds.size(); i += MAX_SHIPMENTS_PER_REQUEST) {
            List<Long> batch = shipmentIds.subList(i, Math.min(i + MAX_SHIPMENTS_PER_REQUEST, shipmentIds.size()));
            List<ZplLabel> batchLabels = descargarBatchZpl(batch, shipmentSkuMap, turboShipmentIds);
            allLabels.addAll(batchLabels);
            AppLogger.info(String.format("ML - Descargadas %d/%d etiquetas ZPL",
                    Math.min(i + MAX_SHIPMENTS_PER_REQUEST, shipmentIds.size()), shipmentIds.size()));
        }

        return allLabels;
    }

    /**
     * Descarga un batch de etiquetas ZPL (máximo 50 shipment IDs).
     * <p><b>IMPORTANTE:</b> ML cambia el substatus de {@code ready_to_print} a {@code printed}
     * automáticamente al llamar a {@code GET /shipment_labels}.</p>
     */
    private static List<ZplLabel> descargarBatchZpl(List<Long> shipmentIds, Map<Long, SkuInfo> skuMap, Set<Long> turboShipmentIds) {
        verificarTokens();

        StringJoiner sj = new StringJoiner(",");
        shipmentIds.forEach(id -> sj.add(String.valueOf(id)));

        String url = "https://api.mercadolibre.com/shipment_labels?shipment_ids=" + sj + "&response_type=zpl2";

        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        // Usar sendWithRetry pero necesitamos bytes, así que hacemos la request manual con retry logic
        HttpResponse<byte[]> response;
        try {
            HttpRequest request = requestBuilder.get();
            response = httpClient.send(request, HttpResponse.BodyHandlers.ofByteArray());
        } catch (IOException | InterruptedException e) {
            AppLogger.warn("ML - Error al descargar etiquetas ZPL: " + e.getMessage());
            return List.of();
        }

        if (response.statusCode() != 200) {
            AppLogger.warn("ML - Error descargando etiquetas (HTTP " + response.statusCode() + "): "
                    + new String(response.body(), StandardCharsets.UTF_8));
            return List.of();
        }

        // La respuesta puede ser un ZIP o texto plano ZPL
        String contentType = response.headers().firstValue("content-type").orElse("");
        String zplContent;

        if (contentType.contains("zip") || contentType.contains("octet-stream")) {
            zplContent = extractZplFromZip(response.body());
        } else {
            zplContent = new String(response.body(), StandardCharsets.UTF_8);
        }

        // Parsear las etiquetas
        List<ZplLabel> parsed = zplParser.parse(zplContent);

        // Enriquecer con datos de las órdenes (SKU más confiable desde API).
        // IMPORTANTE: alinear cada etiqueta con su envío por CONTENIDO (el id dentro del ZPL),
        // NO por posición. ML no garantiza que el orden de las etiquetas devueltas coincida con
        // el de shipment_ids ni que haya exactamente una por envío; emparejar por posición
        // desalinea SKU/zona/COD.EXT cuando se imprimen varias etiquetas.
        List<ZplLabel> enriched = new ArrayList<>();
        for (int i = 0; i < parsed.size(); i++) {
            ZplLabel label = parsed.get(i);

            Long shipId = identificarShipment(label.rawZpl(), shipmentIds);
            if (shipId == null) {
                // Fallback al comportamiento anterior (por posición) + aviso: puede cruzar datos.
                shipId = i < shipmentIds.size() ? shipmentIds.get(i) : null;
                AppLogger.warn("ML - No se pudo identificar el envío de una etiqueta por su contenido; "
                        + "alineando por posición (shipId=" + shipId + "). Si las etiquetas salen con "
                        + "SKU/COD.EXT cruzados, revisá el formato de etiqueta de ML.");
            }

            String sku = label.sku();
            String desc = label.productDescription();
            int qty = 1;
            boolean turbo = false;
            String orderIds = "";
            if (shipId != null) {
                SkuInfo info = skuMap.get(shipId);
                if (info != null) {
                    if (info.sku != null && !info.sku.isBlank()) {
                        sku = info.sku;
                    }
                    if (info.title != null && !info.title.isBlank()) {
                        desc = info.title;
                    }
                    qty = info.quantity;
                    orderIds = info.orderIds;
                }
                turbo = turboShipmentIds.contains(shipId);
            }

            enriched.add(new ZplLabel(label.rawZpl(), sku, desc, label.details(), qty, turbo, orderIds));
        }

        return enriched;
    }

    /**
     * Identifica a qué shipmentId pertenece una etiqueta ZPL mirando su contenido,
     * para no depender del orden en que ML devuelve las etiquetas.
     * <p>Estrategia: 1) campo "Envio: &lt;id&gt;" (coincidencia exacta); 2) el {@code id} del
     * QR, que es el shipmentId seguido de dígitos de seguridad (coincidencia por prefijo).
     *
     * @param candidatos shipmentIds esperados en este batch.
     * @return el shipmentId identificado, o {@code null} si no se puede determinar.
     */
    private static Long identificarShipment(String rawZpl, List<Long> candidatos) {
        // 1) Campo "Envio: <id>"
        Matcher env = ZPL_ENVIO_PATTERN.matcher(rawZpl);
        if (env.find()) {
            try {
                long id = Long.parseLong(env.group(1));
                if (candidatos.contains(id)) return id;
            } catch (NumberFormatException ignored) {
                // id no parseable: probar con el QR
            }
        }
        // 2) id del QR = shipmentId + dígitos de seguridad → match por prefijo.
        // Si varios candidatos son prefijo (caso raro con ids de distinta longitud),
        // gana el más específico (el más largo) para evitar falsos positivos.
        Matcher qr = ZPL_QR_ID_PATTERN.matcher(rawZpl);
        if (qr.find()) {
            String qrId = qr.group(1);
            Long mejor = null;
            for (Long c : candidatos) {
                String s = String.valueOf(c);
                if (qrId.startsWith(s)
                        && (mejor == null || s.length() > String.valueOf(mejor).length())) {
                    mejor = c;
                }
            }
            return mejor;
        }
        return null;
    }

    private static String extractZplFromZip(byte[] zipData) {
        StringBuilder sb = new StringBuilder();
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(zipData))) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (!entry.isDirectory()) {
                    byte[] content = zis.readAllBytes();
                    sb.append(new String(content, StandardCharsets.UTF_8));
                    sb.append("\n");
                }
                zis.closeEntry();
            }
        } catch (IOException e) {
            AppLogger.warn("ML - Error al extraer ZPL del ZIP: " + e.getMessage());
        }
        return sb.toString();
    }

    private record SkuInfo(String sku, String title, int quantity, String orderIds) {
    }

    // -----------------------------------------------------------------------------------------------------------------
    // VENTAS SELLER AGREEMENT (sin envío)
    // -----------------------------------------------------------------------------------------------------------------

    /**
     * Obtiene las ventas de ML sin envío (acuerdo con el vendedor) que NO tengan la nota "impreso".
     */
    public static MLOrderResult obtenerVentasSellerAgreement(String userId) {
        verificarTokens();

        List<Venta> ventas = new ArrayList<>();
        List<OrdenML> ordenes = new ArrayList<>();
        Set<Long> orderIdsSeen = new HashSet<>();
        int offset = 0;
        final int limit = 50;
        boolean hasMore = true;
        int omitidas = 0;

        while (hasMore) {
            final int currentOffset = offset;
            String fechaDesde = OffsetDateTime.now()
                    .minusDays(7)
                    .format(java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:00:00.000XXX"));

            String url = String.format(
                    "https://api.mercadolibre.com/orders/search?seller=%s&tags=no_shipping&order.status=paid&order.date_created.from=%s&sort=date_asc&offset=%d&limit=%d",
                    userId, URLEncoder.encode(fechaDesde, StandardCharsets.UTF_8), currentOffset, limit);

            Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .header("Authorization", "Bearer " + tokens.accessToken)
                    .GET()
                    .build();

            HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

            if (response == null || response.statusCode() != 200) {
                String body = response != null ? response.body() : "sin respuesta";
                AppLogger.warn("ML - Error al obtener órdenes seller_agreement (offset " + currentOffset + "): " + body);
                break;
            }

            JsonNode root = mapper.readTree(response.body());
            JsonNode results = root.path("results");

            if (!results.isArray() || results.isEmpty()) {
                break;
            }

            for (JsonNode order : results) {
                long orderId = order.path("id").asLong();
                if (!orderIdsSeen.add(orderId)) continue;

                JsonNode tagsNode = order.path("tags");
                if (tagsNode.isArray()) {
                    boolean esEntregada = false;
                    for (JsonNode tag : tagsNode) {
                        if ("delivered".equals(tag.asString())) {
                            esEntregada = true;
                            break;
                        }
                    }
                    if (esEntregada) continue;
                }

                if (order.path("fulfilled").asBoolean(false)) continue;

                if (tieneNota(orderId)) {
                    omitidas++;
                    continue;
                }

                String dateCreated = order.path("date_created").asString("");
                OffsetDateTime fecha = null;
                if (!dateCreated.isBlank()) {
                    try {
                        fecha = OffsetDateTime.parse(dateCreated);
                    } catch (Exception e) {
                        AppLogger.warn("ML - Error al parsear fecha de orden " + orderId + ": " + dateCreated);
                    }
                }
                JsonNode packNode = order.path("pack_id");
                Long packId = packNode.isNull() || packNode.isMissingNode() ? null : packNode.asLong();
                OrdenML ordenML = new OrdenML(orderId, packId, null, fecha, "");

                JsonNode orderItems = order.path("order_items");
                if (!orderItems.isArray()) continue;

                for (JsonNode orderItem : orderItems) {
                    JsonNode item = orderItem.path("item");
                    String rawSku = item.path("seller_sku").asString("");
                    if (rawSku.isBlank()) {
                        rawSku = item.path("seller_custom_field").asString("");
                    }
                    String sku = rawSku.isBlank() ? "" : normalizeSku(rawSku);
                    if (sku == null) sku = "";
                    String itemTitle = item.path("title").asString("");
                    double quantity = orderItem.path("quantity").asDouble(0);

                    if (quantity <= 0) {
                        AppLogger.warn("ML Acuerdo - Producto con cantidad inválida en orden " + orderId + ": " + sku);
                        String errorSku = sku.isBlank() ? itemTitle : sku;
                        Venta venta = new Venta("CANT INVALIDA: " + errorSku, quantity, "ML Acuerdo", itemTitle);
                        ventas.add(venta);
                        ordenML.getItems().add(venta);
                        continue;
                    }
                    if (sku.isBlank()) {
                        AppLogger.warn("ML Acuerdo - Producto sin SKU en orden " + orderId + ": " + itemTitle);
                        Venta venta = new Venta("SIN SKU: " + itemTitle, quantity, "ML Acuerdo", itemTitle);
                        ventas.add(venta);
                        ordenML.getItems().add(venta);
                        continue;
                    }
                    Venta venta = new Venta(sku, quantity, "ML Acuerdo", itemTitle);
                    ventas.add(venta);
                    ordenML.getItems().add(venta);
                }

                if (!ordenML.getItems().isEmpty()) {
                    ordenes.add(ordenML);
                }
            }

            JsonNode paging = root.path("paging");
            int total = paging.path("total").asInt(0);
            offset += limit;
            hasMore = offset < total;

            AppLogger.info(String.format("ML - Obtenidas %d/%d órdenes seller_agreement", Math.min(offset, total), total));
        }

        AppLogger.info("ML - Ventas seller_agreement: " + ventas.size() + " (omitidas con nota: " + omitidas + ")");
        return new MLOrderResult(ventas, ordenes);
    }

    // -----------------------------------------------------------------------------------------------------------------
    // STOCK
    // -----------------------------------------------------------------------------------------------------------------

    public static int obtenerStockPorSku(String userId, String sku) {
        verificarTokens();

        String encodedSku = URLEncoder.encode(sku, StandardCharsets.UTF_8);

        String itemId = buscarItemPorSkuParam(userId, encodedSku, "seller_sku");

        if (itemId == null) {
            itemId = buscarItemPorSkuParam(userId, encodedSku, "sku");
        }

        if (itemId == null) {
            return -1;
        }

        return obtenerStockDeItem(itemId);
    }

    private static String buscarItemPorSkuParam(String userId, String encodedSku, String paramName) {
        String url = String.format(
                "https://api.mercadolibre.com/users/%s/items/search?%s=%s",
                userId, paramName, encodedSku);

        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response == null || response.statusCode() != 200) {
            return null;
        }

        try {
            JsonNode root = mapper.readTree(response.body());
            JsonNode results = root.path("results");

            if (!results.isArray() || results.isEmpty()) {
                return null;
            }

            return results.get(0).asString();
        } catch (Exception e) {
            return null;
        }
    }

    private static int obtenerStockDeItem(String itemId) {
        Supplier<HttpRequest> itemRequest = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/items/" + itemId))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> itemResponse = retryHandler.sendWithRetry(itemRequest);

        if (itemResponse == null || itemResponse.statusCode() != 200) {
            AppLogger.warn("ML - Error al obtener item " + itemId + ": " +
                    (itemResponse != null ? itemResponse.body() : "sin respuesta"));
            return -1;
        }

        String userProductId;
        try {
            JsonNode itemRoot = mapper.readTree(itemResponse.body());
            userProductId = itemRoot.path("user_product_id").asString("");
            if (userProductId.isBlank()) {
                return itemRoot.path("available_quantity").asInt(0);
            }
        } catch (Exception e) {
            AppLogger.warn("ML - Error al leer user_product_id de item " + itemId + ": " + e.getMessage());
            return -1;
        }

        Supplier<HttpRequest> stockRequest = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/user-products/" + userProductId + "/stock"))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> stockResponse = retryHandler.sendWithRetry(stockRequest);

        if (stockResponse == null || stockResponse.statusCode() != 200) {
            AppLogger.warn("ML - Error al obtener stock de user_product " + userProductId + ": " +
                    (stockResponse != null ? stockResponse.body() : "sin respuesta"));
            return -1;
        }

        try {
            JsonNode stockRoot = mapper.readTree(stockResponse.body());
            JsonNode locations = stockRoot.path("locations");

            if (!locations.isArray() || locations.isEmpty()) {
                return 0;
            }

            int totalStock = 0;
            for (JsonNode location : locations) {
                totalStock += location.path("quantity").asInt(0);
            }
            return totalStock;
        } catch (Exception e) {
            AppLogger.warn("ML - Error al leer stock de user_product " + userProductId + ": " + e.getMessage());
            return -1;
        }
    }

    public static Map<String, Integer> obtenerStockPorSkus(String userId, List<String> skus) {
        Map<String, Integer> stockMap = new LinkedHashMap<>();

        List<CompletableFuture<Void>> futures = new ArrayList<>();

        for (String sku : skus) {
            CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
                int stock = obtenerStockPorSku(userId, sku);
                synchronized (stockMap) {
                    stockMap.put(sku, stock);
                }
            }, executor);
            futures.add(future);
        }

        CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();

        return stockMap;
    }

    // -----------------------------------------------------------------------------------------------------------------
    // DIMENSIONES DE PAQUETE (ME2)
    // -----------------------------------------------------------------------------------------------------------------

    /**
     * Resuelve SKU → item_id consultando por seller_sku y sku como fallback.
     * Devuelve null si ningún item matchea.
     */
    public static String buscarItemIdPorSku(String userId, String sku) {
        verificarTokens();
        String encoded = URLEncoder.encode(sku, StandardCharsets.UTF_8);
        String itemId = buscarItemPorSkuParam(userId, encoded, "seller_sku");
        if (itemId == null) {
            itemId = buscarItemPorSkuParam(userId, encoded, "sku");
        }
        return itemId;
    }

    public record UploadResult(boolean ok, String itemId, String mensaje) {
    }

    /**
     * Actualiza los atributos SELLER_PACKAGE_* del item. ML exige enteros, cm para dimensiones, g para peso.
     * Entrada: cm y kg (se convierten internamente).
     *
     * Según la doc de ML, los SELLER_PACKAGE_* se declaran SIEMPRE a nivel ítem (no por variación).
     * El ejemplo de la doc los muestra a nivel top sin wrapper variations, y esos atributos no están
     * marcados como variation_attribute en ninguna categoría. Por eso hacemos siempre PUT a nivel ítem.
     * Consecuencia: si el MLA tiene variaciones con distintos embalajes, todas comparten estas medidas.
     */
    public static UploadResult actualizarDimensionesPaquete(String userId, String sku,
                                                            double anchoCm, double altoCm,
                                                            double profundidadCm, double pesoKg) {
        verificarTokens();

        String itemId = buscarItemIdPorSku(userId, sku);
        if (itemId == null) {
            return new UploadResult(false, null, "No se encontró item_id para SKU " + sku);
        }

        int anchoEntero = (int) Math.round(anchoCm);
        int altoEntero = (int) Math.round(altoCm);
        int largoEntero = (int) Math.round(profundidadCm);
        int pesoGramos = (int) Math.round(pesoKg * 1000.0);

        if (anchoEntero <= 0 || altoEntero <= 0 || largoEntero <= 0 || pesoGramos <= 0) {
            return new UploadResult(false, itemId, "Medidas inválidas (enteros > 0 requeridos): "
                    + anchoEntero + "x" + altoEntero + "x" + largoEntero + " cm, " + pesoGramos + " g");
        }

        String body = bodyAtributosNivelItem(anchoEntero, altoEntero, largoEntero, pesoGramos);
        HttpResponse<String> response = ejecutarPut(itemId, body);
        if (response == null) {
            return new UploadResult(false, itemId, "Sin respuesta de ML");
        }
        int status = response.statusCode();
        if (status == 200 || status == 201) {
            return new UploadResult(true, itemId, "OK");
        }
        return new UploadResult(false, itemId, extraerMensajeError(status, response.body()));
    }

    /**
     * GET /items/{id}?include_attributes=all — devuelve el JSON crudo del ítem incluyendo
     * attributes[] de cada variación. Útil para debug/inspección.
     */
    public static String obtenerItemJson(String itemId) {
        verificarTokens();
        Supplier<HttpRequest> req = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/items/" + itemId + "?include_attributes=all"))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();
        HttpResponse<String> response = retryHandler.sendWithRetry(req);
        if (response == null) return null;
        if (response.statusCode() != 200) {
            return "HTTP " + response.statusCode() + " — " + response.body();
        }
        return response.body();
    }

    /**
     * GET /user-products/{up_id} — devuelve el JSON del User Product (modelo nuevo).
     */
    public static String obtenerUserProductJson(String userProductId) {
        verificarTokens();
        Supplier<HttpRequest> req = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/user-products/" + userProductId))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();
        HttpResponse<String> response = retryHandler.sendWithRetry(req);
        if (response == null) return null;
        if (response.statusCode() != 200) {
            return "HTTP " + response.statusCode() + " — " + response.body();
        }
        return response.body();
    }

    /**
     * GET /sites/{site}/user-products-families/{family_id} — devuelve la familia completa de UPs.
     */
    public static String obtenerFamilyJson(String siteId, String familyId) {
        verificarTokens();
        Supplier<HttpRequest> req = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/sites/" + siteId + "/user-products-families/" + familyId))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();
        HttpResponse<String> response = retryHandler.sendWithRetry(req);
        if (response == null) return null;
        if (response.statusCode() != 200) {
            return "HTTP " + response.statusCode() + " — " + response.body();
        }
        return response.body();
    }

    /**
     * GET /users/{user_id} — útil para ver tags como "user_product_seller".
     */
    public static String obtenerUserJson(String userId) {
        verificarTokens();
        Supplier<HttpRequest> req = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/users/" + userId))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();
        HttpResponse<String> response = retryHandler.sendWithRetry(req);
        if (response == null) return null;
        if (response.statusCode() != 200) {
            return "HTTP " + response.statusCode() + " — " + response.body();
        }
        return response.body();
    }

    private static HttpResponse<String> ejecutarPut(String itemId, String body) {
        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/items/" + itemId))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .header("Content-Type", "application/json")
                .header("Accept", "application/json")
                .PUT(HttpRequest.BodyPublishers.ofString(body, StandardCharsets.UTF_8))
                .build();
        return retryHandler.sendWithRetry(requestBuilder);
    }

    private static String bodyAtributosNivelItem(int anchoCm, int altoCm, int largoCm, int pesoGramos) {
        ObjectNode root = mapper.createObjectNode();
        ArrayNode attrs = root.putArray("attributes");
        attrs.addObject().put("id", "SELLER_PACKAGE_WIDTH").put("value_name", anchoCm + " cm");
        attrs.addObject().put("id", "SELLER_PACKAGE_HEIGHT").put("value_name", altoCm + " cm");
        attrs.addObject().put("id", "SELLER_PACKAGE_LENGTH").put("value_name", largoCm + " cm");
        attrs.addObject().put("id", "SELLER_PACKAGE_WEIGHT").put("value_name", pesoGramos + " g");
        return root.toString();
    }

    /**
     * Convierte la respuesta de error de ML en un mensaje legible y corto.
     * Ejemplo de entrada:
     *   {"message":"Validation error","error":"validation_error","status":400,
     *    "cause":[{"cause_id":5401,"code":"item.attribute.invalid...",
     *              "message":"The packaging attributes [seller_package_height] are too small..."}]}
     * Salida: "HTTP 400 · 5401 · The packaging attributes [seller_package_height] are too small..."
     */
    private static String extraerMensajeError(int status, String body) {
        String prefix = "HTTP " + status;
        if (body == null || body.isBlank()) return prefix;
        try {
            JsonNode root = mapper.readTree(body);
            JsonNode cause = root.path("cause");
            if (cause.isArray() && !cause.isEmpty()) {
                JsonNode first = cause.get(0);
                String causeId = first.path("cause_id").asString("");
                String causeMsg = first.path("message").asString("");
                if (!causeMsg.isBlank()) {
                    return causeId.isBlank()
                            ? prefix + " · " + causeMsg
                            : prefix + " · " + causeId + " · " + causeMsg;
                }
            }
            String topMsg = root.path("message").asString("");
            if (!topMsg.isBlank()) return prefix + " · " + topMsg;
        } catch (Exception ignored) {
            // Cuerpo no-JSON: caemos al fallback.
        }
        return prefix + ": " + body;
    }

    // -----------------------------------------------------------------------------------------------------------------
    // SLA
    // -----------------------------------------------------------------------------------------------------------------

    public static SlaInfo obtenerSla(long shipmentId) {
        verificarTokens();

        // 1) Obtener SLA
        String slaUrl = "https://api.mercadolibre.com/shipments/" + shipmentId + "/sla";
        Supplier<HttpRequest> slaReq = () -> HttpRequest.newBuilder()
                .uri(URI.create(slaUrl))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        OffsetDateTime expectedDate = null;
        String status = "";
        HttpResponse<String> slaResponse = retryHandler.sendWithRetry(slaReq);
        if (slaResponse != null && slaResponse.statusCode() == 200) {
            try {
                JsonNode root = mapper.readTree(slaResponse.body());
                status = root.path("status").asString("");
                String expectedDateStr = root.path("expected_date").asString("");
                if (!expectedDateStr.isBlank()) {
                    try {
                        expectedDate = OffsetDateTime.parse(expectedDateStr);
                    } catch (Exception e) {
                        AppLogger.warn("ML - Error al parsear expected_date de SLA shipment " + shipmentId + ": " + expectedDateStr);
                    }
                }
            } catch (Exception e) {
                AppLogger.warn("ML - Error al leer SLA de shipment " + shipmentId + ": " + e.getMessage());
            }
        }

        // 2) Obtener tags del shipment para detectar turbo
        boolean turbo = false;
        String shipUrl = "https://api.mercadolibre.com/shipments/" + shipmentId;
        Supplier<HttpRequest> shipReq = () -> HttpRequest.newBuilder()
                .uri(URI.create(shipUrl))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> shipResponse = retryHandler.sendWithRetry(shipReq);
        if (shipResponse != null && shipResponse.statusCode() == 200) {
            try {
                JsonNode root = mapper.readTree(shipResponse.body());
                JsonNode tags = root.path("tags");
                if (tags.isArray()) {
                    for (JsonNode tag : tags) {
                        if ("turbo".equals(tag.asString())) {
                            turbo = true;
                            break;
                        }
                    }
                }
            } catch (Exception e) {
                AppLogger.warn("ML - Error al leer tags de shipment " + shipmentId + ": " + e.getMessage());
            }
        }

        return new SlaInfo(status, expectedDate, turbo);
    }

    public static Map<Long, SlaInfo> obtenerSlasParalelo(List<Long> shipmentIds) {
        Map<Long, SlaInfo> slaMap = new LinkedHashMap<>();

        List<CompletableFuture<Void>> futures = new ArrayList<>();

        for (Long shipmentId : shipmentIds) {
            CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
                SlaInfo sla = obtenerSla(shipmentId);
                if (sla != null) {
                    synchronized (slaMap) {
                        slaMap.put(shipmentId, sla);
                    }
                }
            }, executor);
            futures.add(future);
        }

        CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();

        return slaMap;
    }

    public static Map<Long, String> obtenerShipmentSubstatuses(List<Long> shipmentIds) {
        Map<Long, String> substatusMap = new LinkedHashMap<>();

        List<CompletableFuture<Void>> futures = new ArrayList<>();

        for (Long shipmentId : shipmentIds) {
            CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
                verificarTokens();
                String url = "https://api.mercadolibre.com/shipments/" + shipmentId;

                Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                        .uri(URI.create(url))
                        .header("Authorization", "Bearer " + tokens.accessToken)
                        .GET()
                        .build();

                HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

                if (response != null && response.statusCode() == 200) {
                    JsonNode root = mapper.readTree(response.body());
                    String status = root.path("status").asString("");
                    String substatus = root.path("substatus").asString("");
                    String combined = substatus.isEmpty() ? status : substatus;
                    synchronized (substatusMap) {
                        substatusMap.put(shipmentId, combined);
                    }
                }
            }, executor);
            futures.add(future);
        }

        CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();

        return substatusMap;
    }

    /**
     * Obtiene substatus y SLA de múltiples envíos en paralelo con UNA sola llamada por envío.
     * Reemplaza las llamadas separadas a obtenerSlasParalelo + obtenerShipmentSubstatuses.
     */
    public static Map<Long, ShipmentInfo> obtenerShipmentInfoParalelo(List<Long> shipmentIds) {
        Map<Long, ShipmentInfo> result = new LinkedHashMap<>();
        List<CompletableFuture<Void>> futures = new ArrayList<>();

        for (Long shipmentId : shipmentIds) {
            futures.add(CompletableFuture.runAsync(() -> {
                verificarTokens();

                // 1) Obtener status/substatus del envío
                String shipUrl = "https://api.mercadolibre.com/shipments/" + shipmentId;
                Supplier<HttpRequest> shipReq = () -> HttpRequest.newBuilder()
                        .uri(URI.create(shipUrl))
                        .header("Authorization", "Bearer " + tokens.accessToken)
                        .header("X-Format-New", "true")
                        .GET()
                        .build();

                HttpResponse<String> shipResponse = retryHandler.sendWithRetry(shipReq);
                String combined = "";
                if (shipResponse != null && shipResponse.statusCode() == 200) {
                    JsonNode root = mapper.readTree(shipResponse.body());
                    String status = root.path("status").asString("");
                    String substatus = root.path("substatus").asString("");
                    combined = substatus.isEmpty() ? status : substatus;
                }

                // 2) Obtener SLA del endpoint dedicado (estimated_handling_limit deprecado desde mayo 2025)
                OffsetDateTime slaDate = null;
                SlaInfo sla = obtenerSla(shipmentId);
                if (sla != null && sla.expectedDate() != null) {
                    slaDate = sla.expectedDate();
                }

                synchronized (result) {
                    result.put(shipmentId, new ShipmentInfo(combined, slaDate));
                }
            }, executor));
        }

        CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();
        return result;
    }

    // -----------------------------------------------------------------------------------------------------------------
    // DIAGNOSTICO
    // -----------------------------------------------------------------------------------------------------------------

    /**
     * Consulta una orden por ID y loguea su estado y el de su envío. Útil para depuración.
     */
    public static void diagnosticarOrden(long orderId) {
        verificarTokens();
        String url = "https://api.mercadolibre.com/orders/" + orderId;

        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response == null) {
            AppLogger.info("DIAG - Orden " + orderId + ": sin respuesta");
            return;
        }
        if (response.statusCode() != 200) {
            AppLogger.info("DIAG - Orden " + orderId + ": HTTP " + response.statusCode() + " - " + response.body());
            return;
        }

        JsonNode root = mapper.readTree(response.body());
        String orderStatus = root.path("status").asString("");
        long shippingId = root.path("shipping").path("id").asLong(0);
        JsonNode tags = root.path("tags");

        AppLogger.info("DIAG - Orden " + orderId + ": status=" + orderStatus + ", shipping.id=" + shippingId + ", tags=" + tags);

        // Consultar el envío
        if (shippingId > 0) {
            String shipUrl = "https://api.mercadolibre.com/shipments/" + shippingId;
            Supplier<HttpRequest> shipReq = () -> HttpRequest.newBuilder()
                    .uri(URI.create(shipUrl))
                    .header("Authorization", "Bearer " + tokens.accessToken)
                    .header("X-Format-New", "true")
                    .GET()
                    .build();

            HttpResponse<String> shipResponse = retryHandler.sendWithRetry(shipReq);
            if (shipResponse != null && shipResponse.statusCode() == 200) {
                JsonNode shipRoot = mapper.readTree(shipResponse.body());
                String shipStatus = shipRoot.path("status").asString("");
                String shipSubstatus = shipRoot.path("substatus").asString("");
                JsonNode leadTime = shipRoot.path("lead_time");
                AppLogger.info("DIAG - Shipment " + shippingId + ": status=" + shipStatus + ", substatus=" + shipSubstatus + ", lead_time=" + leadTime);
            } else {
                AppLogger.info("DIAG - Shipment " + shippingId + ": error al consultar");
            }
        }
    }

    // -----------------------------------------------------------------------------------------------------------------
    // PEDIDOS RETIRO (sin envío, para tab Pedidos)
    // -----------------------------------------------------------------------------------------------------------------

    public static List<PedidoML> obtenerPedidosRetiro(String userId) {
        verificarTokens();

        List<PedidoML> pedidos = new ArrayList<>();
        Set<Long> orderIdsSeen = new HashSet<>();
        int offset = 0;
        final int limit = 50;
        boolean hasMore = true;
        int omitidas = 0;

        while (hasMore) {
            final int currentOffset = offset;
            String fechaDesde = OffsetDateTime.now()
                    .minusDays(7)
                    .format(java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:00:00.000XXX"));

            String url = String.format(
                    "https://api.mercadolibre.com/orders/search?seller=%s&tags=no_shipping&order.status=paid&order.date_created.from=%s&sort=date_asc&offset=%d&limit=%d",
                    userId, URLEncoder.encode(fechaDesde, StandardCharsets.UTF_8), currentOffset, limit);

            Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .header("Authorization", "Bearer " + tokens.accessToken)
                    .GET()
                    .build();

            HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

            if (response == null || response.statusCode() != 200) {
                String body = response != null ? response.body() : "sin respuesta";
                AppLogger.warn("PEDIDOS ML - Error al obtener órdenes retiro (offset " + currentOffset + "): " + body);
                break;
            }

            JsonNode root = mapper.readTree(response.body());
            JsonNode results = root.path("results");

            if (!results.isArray() || results.isEmpty()) break;

            for (JsonNode order : results) {
                long orderId = order.path("id").asLong();
                if (!orderIdsSeen.add(orderId)) continue;

                // Skip delivered
                JsonNode tagsNode = order.path("tags");
                if (tagsNode.isArray()) {
                    boolean esEntregada = false;
                    for (JsonNode tag : tagsNode) {
                        if ("delivered".equals(tag.asString())) {
                            esEntregada = true;
                            break;
                        }
                    }
                    if (esEntregada) continue;
                }

                // Skip fulfilled
                if (order.path("fulfilled").asBoolean(false)) continue;

                // Skip con nota
                if (tieneNota(orderId)) {
                    omitidas++;
                    continue;
                }

                String dateCreated = order.path("date_created").asString("");
                OffsetDateTime fecha = null;
                if (!dateCreated.isBlank()) {
                    try {
                        fecha = OffsetDateTime.parse(dateCreated);
                    } catch (Exception ignored) {
                    }
                }

                // Buyer ID
                long buyerId = order.path("buyer").path("id").asLong(0);

                // Items
                JsonNode orderItems = order.path("order_items");
                if (!orderItems.isArray()) continue;

                for (JsonNode orderItem : orderItems) {
                    JsonNode item = orderItem.path("item");
                    String rawSku = item.path("seller_sku").asString("");
                    if (rawSku.isBlank()) rawSku = item.path("seller_custom_field").asString("");
                    String sku = rawSku.isBlank() ? "" : normalizeSku(rawSku);
                    if (sku == null) sku = "";
                    String detalle = item.path("title").asString("");
                    double quantity = orderItem.path("quantity").asDouble(0);

                    pedidos.add(new PedidoML(orderId, fecha, "", "", sku, quantity, detalle, buyerId));
                }
            }

            JsonNode paging = root.path("paging");
            int total = paging.path("total").asInt(0);
            offset += limit;
            hasMore = offset < total;
            AppLogger.info(String.format("PEDIDOS ML - Obtenidas %d/%d órdenes retiro", Math.min(offset, total), total));
        }

        AppLogger.info("PEDIDOS ML - Pedidos retiro: " + pedidos.size() + " (omitidas con nota: " + omitidas + ")");

        // Obtener datos de compradores via GET /orders/{id} en paralelo
        Set<Long> orderIdsUnicos = new LinkedHashSet<>();
        for (PedidoML p : pedidos) orderIdsUnicos.add(p.orderId());

        if (!orderIdsUnicos.isEmpty()) {
            AppLogger.info("PEDIDOS ML - Obteniendo datos de compradores (" + orderIdsUnicos.size() + " órdenes)...");
            Map<Long, String[]> buyerMap = obtenerBuyersPorOrden(new ArrayList<>(orderIdsUnicos));

            for (int i = 0; i < pedidos.size(); i++) {
                PedidoML p = pedidos.get(i);
                String[] info = buyerMap.get(p.orderId());
                if (info != null) {
                    pedidos.set(i, new PedidoML(p.orderId(), p.fecha(), info[0], info[1], p.sku(), p.cantidad(), p.detalle(), p.buyerId()));
                }
            }
        }

        return pedidos;
    }

    private static Map<Long, String[]> obtenerBuyersPorOrden(List<Long> orderIds) {
        Map<Long, String[]> buyerMap = new LinkedHashMap<>();
        List<CompletableFuture<Void>> futures = new ArrayList<>();

        for (Long orderId : orderIds) {
            CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
                verificarTokens();
                String url = "https://api.mercadolibre.com/orders/" + orderId;
                Supplier<HttpRequest> req = () -> HttpRequest.newBuilder()
                        .uri(URI.create(url))
                        .header("Authorization", "Bearer " + tokens.accessToken)
                        .GET()
                        .build();

                HttpResponse<String> response = retryHandler.sendWithRetry(req);
                if (response != null && response.statusCode() == 200) {
                    try {
                        JsonNode buyer = mapper.readTree(response.body()).path("buyer");
                        String nickname = buyer.path("nickname").asString("");
                        String firstName = buyer.path("first_name").asString("");
                        String lastName = buyer.path("last_name").asString("");
                        String nombreApellido = (firstName + " " + lastName).trim();
                        synchronized (buyerMap) {
                            buyerMap.put(orderId, new String[]{nickname, nombreApellido});
                        }
                    } catch (Exception e) {
                        AppLogger.warn("ML - Error al leer buyer de orden " + orderId + ": " + e.getMessage());
                    }
                }
            }, executor);
            futures.add(future);
        }

        CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();
        return buyerMap;
    }

    // -----------------------------------------------------------------------------------------------------------------
    // NOTAS
    // -----------------------------------------------------------------------------------------------------------------

    private static boolean tieneNota(long orderId) {
        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/orders/" + orderId + "/notes"))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response == null || response.statusCode() != 200) {
            return false;
        }

        try {
            JsonNode root = mapper.readTree(response.body());
            if (!root.isArray() || root.isEmpty()) return false;

            JsonNode results = root.get(0).path("results");
            if (!results.isArray()) return false;

            for (JsonNode note : results) {
                String texto = note.path("note").asString("").trim();
                if (!texto.isEmpty()) {
                    return true;
                }
            }
        } catch (Exception e) {
            AppLogger.warn("ML - Error al leer notas de orden " + orderId + ": " + e.getMessage());
        }

        return false;
    }

    // -----------------------------------------------------------------------------------------------------------------
    // TOKENS
    // -----------------------------------------------------------------------------------------------------------------

    public static boolean inicializar() {
        mlCredentials = cargarMLCredentials();
        if (mlCredentials == null) {
            AppLogger.warn("ML - No se encontró el archivo de credenciales.");
            return false;
        }

        tokens = cargarTokens();
        if (tokens == null) {
            AppLogger.info("ML - No hay tokens de ML, solicitando autorización...");
            final String code = pedirCodeManual();
            tokens = obtenerAccessToken(code);
            guardarTokens(tokens);
        }

        return true;
    }

    public static void verificarTokens() {
        if (tokens == null) {
            AppLogger.warn("ML - Tokens no inicializados. Intentando inicializar...");
            if (!inicializar()) {
                throw new IllegalStateException("ML - No se pudieron inicializar los tokens.");
            }
            return;
        }

        if (!tokens.isExpired()) {
            return;
        }

        synchronized (TOKEN_LOCK) {
            if (tokens == null || !tokens.isExpired()) {
                return;
            }

            AppLogger.info("ML - Access token expirado, renovando...");
            try {
                tokens = refreshAccessToken(tokens.refreshToken);
                tokens.issuedAt = System.currentTimeMillis();
                guardarTokens(tokens);
                AppLogger.info("ML - Token renovado correctamente.");
            } catch (Exception e) {
                AppLogger.warn("ML - Error al renovar token: " + e.getMessage());
                throw new RuntimeException("No se pudo renovar el token de ML", e);
            }
        }
    }

    private static MLCredentials cargarMLCredentials() {
        try {
            File f = MERCADOLIBRE_FILE.toFile();
            return f.exists() ? mapper.readValue(f, MLCredentials.class) : null;
        } catch (Exception e) {
            AppLogger.warn("Error cargando credenciales ML: " + e.getMessage());
            return null;
        }
    }

    private static TokensML cargarTokens() {
        try {
            File f = TOKEN_FILE.toFile();
            return f.exists() ? mapper.readValue(f, TokensML.class) : null;
        } catch (Exception e) {
            AppLogger.warn("Error cargando tokens ML: " + e.getMessage());
            return null;
        }
    }

    private static void guardarTokens(TokensML tokens) {
        try {
            mapper.writerWithDefaultPrettyPrinter().writeValue(TOKEN_FILE.toFile(), tokens);
            AppLogger.info("ML - Tokens guardados en " + TOKEN_FILE);
        } catch (Exception e) {
            AppLogger.warn("Error guardando tokens ML: " + e.getMessage());
        }
    }

    private static String pedirCodeManual() {
        String authURL = "https://auth.mercadolibre.com.ar/authorization?response_type=code"
                + "&client_id=" + mlCredentials.clientId
                + "&redirect_uri=" + mlCredentials.redirectUri;

        AppLogger.info("ML - Se necesita autorización manual. Abriendo diálogo...");

        CompletableFuture<String> future = new CompletableFuture<>();
        Platform.runLater(() -> {
            TextInputDialog dialog = new TextInputDialog();
            dialog.setTitle("Autorización MercadoLibre");
            dialog.setHeaderText("Autorizá la app en esta URL y pegá el code:");
            dialog.setContentText(authURL);
            dialog.getDialogPane().setPrefWidth(600);
            dialog.showAndWait().ifPresentOrElse(
                    code -> future.complete(code.trim()),
                    () -> future.complete(null)
            );
        });

        try {
            String code = future.get();
            if (code == null || code.isBlank()) {
                throw new RuntimeException("ML - Autorización cancelada por el usuario.");
            }
            return code;
        } catch (Exception e) {
            throw new RuntimeException("ML - Error al obtener code de autorización.", e);
        }
    }

    private static TokensML obtenerAccessToken(String code) {

        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/oauth/token"))
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(
                        "grant_type=authorization_code" +
                                "&client_id=" + mlCredentials.clientId +
                                "&client_secret=" + mlCredentials.clientSecret +
                                "&code=" + code +
                                "&redirect_uri=" + mlCredentials.redirectUri))
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);
        if (response.statusCode() != 200) {
            throw new RuntimeException("Error al obtener access_token: " + response.body());
        }

        TokensML tokens = mapper.readValue(response.body(), TokensML.class);
        tokens.issuedAt = System.currentTimeMillis();
        return tokens;
    }

    private static TokensML refreshAccessToken(String refreshToken) {
        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/oauth/token"))
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(
                        "grant_type=refresh_token" +
                                "&client_id=" + mlCredentials.clientId +
                                "&client_secret=" + mlCredentials.clientSecret +
                                "&refresh_token=" + refreshToken))
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);
        if (response.statusCode() != 200) {
            throw new RuntimeException("Error al refrescar access_token: " + response.body());
        }

        TokensML tokens = mapper.readValue(response.body(), TokensML.class);
        tokens.issuedAt = System.currentTimeMillis();
        return tokens;
    }

}
