package ar.com.leo.pickit.service;

import ar.com.leo.AppLogger;
import ar.com.leo.api.ml.MercadoLibreAPI;
import ar.com.leo.api.ml.MercadoLibreAPI.MLOrderResult;
import ar.com.leo.api.ml.MercadoLibreAPI.SlaInfo;
import ar.com.leo.api.ml.model.OrdenML;
import ar.com.leo.api.ml.model.Venta;
import ar.com.leo.api.nube.TiendaNubeApi;
import ar.com.leo.pickit.excel.ExcelManager;
import ar.com.leo.pickit.excel.ExcelManager.ComboEntry;
import ar.com.leo.pickit.excel.ExcelManager.ProductoStock;
import ar.com.leo.pickit.excel.PickitExcelWriter;
import ar.com.leo.pickit.model.*;

import java.io.File;
import java.time.*;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class PickitGenerator {

    public record SlaOrden(String numeroVenta, int cantidadItems, String slaStatus, OffsetDateTime slaExpectedDate) {}

    private static final ExecutorService executor = Executors.newFixedThreadPool(4);

    public static void shutdownExecutors() {
        executor.shutdown();
        try {
            if (!executor.awaitTermination(5, TimeUnit.SECONDS)) {
                executor.shutdownNow();
            }
        } catch (InterruptedException e) {
            executor.shutdownNow();
            Thread.currentThread().interrupt();
        }
    }

    public static File generarPickit(File stockExcel, File combosExcel, List<ProductoManual> productosManuales, boolean soloHoy) throws Exception {

        AppLogger.info("PICKIT - Despacho ML: " + (soloHoy ? "Hasta hoy 23:59:59" : "Sin límite"));

        // Paso 1: Inicializar APIs
        AppLogger.info("PICKIT - Paso 1: Inicializando APIs (MercadoLibre + Tienda Nube)...");
        if (!MercadoLibreAPI.inicializar()) {
            throw new RuntimeException("No se pudieron inicializar los tokens de MercadoLibre.");
        }
        final String userId = MercadoLibreAPI.getUserId();

        if (!TiendaNubeApi.inicializar()) {
            throw new RuntimeException("No se pudieron inicializar las credenciales de Tienda Nube.");
        }

        final List<Venta> todasLasVentas = Collections.synchronizedList(new ArrayList<>());
        final List<OrdenML> todasLasOrdenesML = Collections.synchronizedList(new ArrayList<>());

        // Pasos 2-5: Obtener ventas en paralelo
        var futureMLPrint = executor.submit(() -> {
            AppLogger.info("PICKIT - Paso 2: Obteniendo ventas ML ready_to_print...");
            MLOrderResult result = MercadoLibreAPI.obtenerVentasReadyToPrint(userId, false);
            todasLasVentas.addAll(result.ventas());
            todasLasOrdenesML.addAll(result.ordenes());
            return result.ventas().size();
        });

        var futureMLAgreement = executor.submit(() -> {
            AppLogger.info("PICKIT - Paso 3: Obteniendo ventas ML acuerdo con el vendedor...");
            MLOrderResult result = MercadoLibreAPI.obtenerVentasSellerAgreement(userId);
            todasLasVentas.addAll(result.ventas());
            todasLasOrdenesML.addAll(result.ordenes());
            return result.ventas().size();
        });

        var futureNube = executor.submit(() -> {
            AppLogger.info("PICKIT - Paso 4: Obteniendo ventas KT HOGAR (Tienda Nube)...");
            var result = TiendaNubeApi.obtenerVentasHogar();
            todasLasVentas.addAll(result.ventas());
            return result;
        });

        var futureGastro = executor.submit(() -> {
            AppLogger.info("PICKIT - Paso 5: Obteniendo ventas KT GASTRO (Tienda Nube)...");
            var result = TiendaNubeApi.obtenerVentasGastro();
            todasLasVentas.addAll(result.ventas());
            return result;
        });

        int countMLPrint = futureMLPrint.get();
        int countMLAgreement = futureMLAgreement.get();
        var resultNube = futureNube.get();
        var resultGastro = futureGastro.get();
        int countNube = resultNube.totalOrdenes();
        int countGastro = resultGastro.totalOrdenes();

        // Paso 6: Consolidar ventas
        AppLogger.info("PICKIT - Paso 6: Consolidando ventas...");
        AppLogger.info(String.format("ML ready_to_print: %d | ML acuerdo: %d | KT HOGAR: %d | KT GASTRO: %d | Total: %d",
                countMLPrint, countMLAgreement, countNube, countGastro, todasLasVentas.size()));

        if (productosManuales != null && !productosManuales.isEmpty()) {
            for (ProductoManual pm : productosManuales) {
                todasLasVentas.add(new Venta(pm.getSku(), pm.getCantidad(), "MANUAL"));
            }
            AppLogger.info("PICKIT - Productos manuales agregados: " + productosManuales.size());
        }

        if (todasLasVentas.isEmpty()) {
            throw new RuntimeException("No se encontraron ventas para procesar. Verificar conexión a las APIs.");
        }

        // Obtener SLAs
        Set<Long> shipmentIdsUnicos = new LinkedHashSet<>();
        for (OrdenML orden : todasLasOrdenesML) {
            if (orden.getShipmentId() != null) shipmentIdsUnicos.add(orden.getShipmentId());
        }
        Map<Long, SlaInfo> slaMap = Collections.emptyMap();
        if (!shipmentIdsUnicos.isEmpty()) {
            AppLogger.info("PICKIT - Obteniendo SLAs para " + shipmentIdsUnicos.size() + " envíos...");
            slaMap = MercadoLibreAPI.obtenerSlasParalelo(new ArrayList<>(shipmentIdsUnicos));
            AppLogger.info("PICKIT - SLAs obtenidos: " + slaMap.size());
        }

        // Filtrar por SLA si modo "Hoy"
        if (soloHoy && !slaMap.isEmpty()) {
            ZoneId zonaArgentina = ZoneId.of("America/Argentina/Buenos_Aires");
            OffsetDateTime limiteHoy = LocalDate.now(zonaArgentina)
                    .atTime(23, 59, 59)
                    .atZone(zonaArgentina)
                    .toOffsetDateTime();

            Map<Long, List<OrdenML>> tempPorVenta = new LinkedHashMap<>();
            for (OrdenML orden : todasLasOrdenesML) {
                tempPorVenta.computeIfAbsent(orden.getVentaId(), k -> new ArrayList<>()).add(orden);
            }

            Set<Long> ventaIdsExcluir = new HashSet<>();
            for (Map.Entry<Long, List<OrdenML>> entry : tempPorVenta.entrySet()) {
                for (OrdenML orden : entry.getValue()) {
                    if (orden.getShipmentId() != null && slaMap.containsKey(orden.getShipmentId())) {
                        SlaInfo sla = slaMap.get(orden.getShipmentId());
                        if (sla.expectedDate() != null && sla.expectedDate().isAfter(limiteHoy)) {
                            ventaIdsExcluir.add(entry.getKey());
                        }
                        break;
                    }
                }
            }

            if (!ventaIdsExcluir.isEmpty()) {
                Set<Venta> ventasARemover = Collections.newSetFromMap(new IdentityHashMap<>());
                for (Long ventaId : ventaIdsExcluir) {
                    for (OrdenML orden : tempPorVenta.get(ventaId)) {
                        ventasARemover.addAll(orden.getItems());
                    }
                }
                todasLasVentas.removeAll(ventasARemover);
                todasLasOrdenesML.removeIf(o -> ventaIdsExcluir.contains(o.getVentaId()));
                AppLogger.info("PICKIT - Filtro SLA Hoy: " + ventaIdsExcluir.size() + " órdenes ML excluidas");
            }
        }

        // Paso 7: Limpiar SKUs
        AppLogger.info("PICKIT - Paso 7: Limpiando SKUs...");
        for (Venta venta : todasLasVentas) {
            String sku = venta.getSku().trim();
            if (esSkuConError(sku)) continue;
            int spaceIndex = sku.indexOf(' ');
            if (spaceIndex > 0) sku = sku.substring(0, spaceIndex);
            sku = sku.replaceAll("^[^0-9]*|[^0-9]*$", "");
            venta.setSku(sku);
        }

        for (Venta v : todasLasVentas) {
            String sku = v.getSku();
            if (esSkuConError(sku)) continue;
            if (sku.isBlank() || !sku.matches("\\d+")) {
                AppLogger.warn("PICKIT - SKU inválido: '" + sku + "'");
                v.setSku("SKU INVALIDO: " + sku);
            }
        }

        // Paso 8: Leer combos y expandir
        AppLogger.info("PICKIT - Paso 8: Leyendo combos y expandiendo...");
        Map<String, List<ComboEntry>> combos = ExcelManager.obtenerCombos(combosExcel);

        List<Venta> ventasExpandidas = new ArrayList<>();
        for (Venta venta : todasLasVentas) {
            if (esSkuConError(venta.getSku())) {
                ventasExpandidas.add(venta);
                continue;
            }
            List<ComboEntry> componentes = combos.get(venta.getSku());
            if (componentes != null && !componentes.isEmpty()) {
                for (ComboEntry comp : componentes) {
                    double cantidadExpandida = venta.getCantidad() * comp.cantidad();
                    if (cantidadExpandida <= 0) {
                        ventasExpandidas.add(new Venta("COMBO INVALIDO: " + comp.skuComponente(), cantidadExpandida, venta.getOrigen()));
                    } else {
                        ventasExpandidas.add(new Venta(comp.skuComponente(), cantidadExpandida, venta.getOrigen()));
                    }
                }
            } else {
                ventasExpandidas.add(venta);
            }
        }

        // Paso 9: Agrupar por SKU
        AppLogger.info("PICKIT - Paso 9: Agrupando por SKU...");
        Map<String, Double> skuCantidad = new LinkedHashMap<>();
        for (Venta venta : ventasExpandidas) {
            skuCantidad.merge(venta.getSku(), venta.getCantidad(), Double::sum);
        }
        AppLogger.info("PICKIT - SKUs únicos: " + skuCantidad.size());

        // Paso 10: Leer datos de productos
        AppLogger.info("PICKIT - Paso 10: Leyendo datos de productos de Stock.xlsx...");
        Map<String, ProductoStock> productosStock = ExcelManager.obtenerProductosStock(stockExcel);

        List<PickitItem> pickitItems = new ArrayList<>();
        int skusNoEncontrados = 0, skusStockInsuficiente = 0, skusConError = 0;

        for (Map.Entry<String, Double> entry : skuCantidad.entrySet()) {
            String sku = entry.getKey();
            double cantidad = entry.getValue();

            ProductoStock producto = productosStock.get(sku);
            String descripcion = "", proveedor = "", subRubro = "", unidad = "";
            int stockDisponible = 0;

            if (sku.startsWith("SIN SKU: ")) {
                descripcion = sku.substring("SIN SKU: ".length());
                sku = "SIN SKU";
                skusConError++;
            } else if (esSkuConError(sku)) {
                skusConError++;
            } else if (producto != null) {
                descripcion = producto.producto();
                proveedor = producto.proveedor();
                subRubro = producto.subRubro();
                unidad = producto.unidad();
                stockDisponible = producto.stock();
                if (stockDisponible < cantidad) {
                    AppLogger.warn("PICKIT - SKU " + sku + " stock insuficiente (pedido: " + (int) cantidad + ", disponible: " + stockDisponible + ")");
                    skusStockInsuficiente++;
                }
            } else {
                AppLogger.warn("PICKIT - SKU " + sku + " no encontrado en Stock.xlsx");
                skusNoEncontrados++;
            }

            pickitItems.add(new PickitItem(sku, cantidad, descripcion, proveedor, unidad, stockDisponible, subRubro));
        }

        // Paso 11: Ordenar
        AppLogger.info("PICKIT - Paso 11: Ordenando resultados...");
        pickitItems.sort(Comparator
                .comparing((PickitItem i) -> i.getUnidad() != null ? i.getUnidad() : "")
                .thenComparing(i -> i.getProveedor() != null ? i.getProveedor() : "")
                .thenComparing(i -> i.getSubRubro() != null ? i.getSubRubro() : "")
                .thenComparing(i -> i.getDescripcion() != null ? i.getDescripcion() : ""));

        // Construir CARROS
        AppLogger.info("PICKIT - Construyendo datos de CARROS...");
        todasLasOrdenesML.sort(Comparator
                .comparing((OrdenML o) -> o.getFechaCreacion() != null ? o.getFechaCreacion() : OffsetDateTime.MAX)
                .thenComparingLong(OrdenML::getOrderId));

        Map<Long, List<OrdenML>> ordenesPorVenta = new LinkedHashMap<>();
        for (OrdenML orden : todasLasOrdenesML) {
            ordenesPorVenta.computeIfAbsent(orden.getVentaId(), k -> new ArrayList<>()).add(orden);
        }

        // SLA ordenes
        List<SlaOrden> slaOrdenes = new ArrayList<>();
        for (Map.Entry<Long, List<OrdenML>> entry : ordenesPorVenta.entrySet()) {
            List<OrdenML> ordenesGrupo = entry.getValue();
            for (OrdenML orden : ordenesGrupo) {
                if (orden.getShipmentId() != null && slaMap.containsKey(orden.getShipmentId())) {
                    SlaInfo sla = slaMap.get(orden.getShipmentId());
                    int cantItems = ordenesGrupo.stream().mapToInt(o -> o.getItems().size()).sum();
                    slaOrdenes.add(new SlaOrden(ordenesGrupo.getFirst().getNumeroVenta(), cantItems, sla.status(), sla.expectedDate()));
                    break;
                }
            }
        }
        slaOrdenes.sort(Comparator.comparing(s -> s.slaExpectedDate() != null ? s.slaExpectedDate() : OffsetDateTime.MAX));
        AppLogger.info("PICKIT - Ordenes con SLA: " + slaOrdenes.size());

        List<CarrosOrden> carrosOrdenes = new ArrayList<>();
        int carroIndex = 0;
        for (Map.Entry<Long, List<OrdenML>> entry : ordenesPorVenta.entrySet()) {
            List<OrdenML> ordenesGrupo = entry.getValue();
            String letra = generarLetraCarro(carroIndex++);
            String numeroVenta = ordenesGrupo.getFirst().getNumeroVenta();
            OffsetDateTime fechaCreacion = ordenesGrupo.getFirst().getFechaCreacion();

            Set<String> skusOriginales = new HashSet<>();
            for (OrdenML orden : ordenesGrupo) {
                for (Venta v : orden.getItems()) {
                    String skuV = v.getSku();
                    if (skuV.startsWith("SIN SKU: ")) skusOriginales.add("SIN SKU");
                    else if (esSkuConError(skuV)) skusOriginales.add(skuV);
                    else if (skuV.isBlank() || !skuV.matches("\\d+")) skusOriginales.add("SKU INVALIDO: " + skuV);
                    else skusOriginales.add(skuV);
                }
            }

            if (skusOriginales.size() < 3) { carroIndex--; continue; }

            List<CarrosItem> carrosItems = new ArrayList<>();
            for (OrdenML orden : ordenesGrupo) {
                for (Venta v : orden.getItems()) {
                    String skuItem = v.getSku();
                    if (skuItem.startsWith("SIN SKU: ")) {
                        carrosItems.add(new CarrosItem("SIN SKU", v.getCantidad(), skuItem.substring("SIN SKU: ".length()), ""));
                    } else if (esSkuConError(skuItem)) {
                        carrosItems.add(new CarrosItem(skuItem, v.getCantidad(), "", ""));
                    } else if (skuItem.isBlank() || !skuItem.matches("\\d+")) {
                        carrosItems.add(new CarrosItem("SKU INVALIDO: " + skuItem, v.getCantidad(), "", ""));
                    } else {
                        List<ComboEntry> componentes = combos.get(skuItem);
                        if (componentes != null && !componentes.isEmpty()) {
                            for (ComboEntry comp : componentes) {
                                double cantExpand = v.getCantidad() * comp.cantidad();
                                ProductoStock p = productosStock.get(comp.skuComponente());
                                carrosItems.add(new CarrosItem(comp.skuComponente(), cantExpand,
                                        p != null ? p.producto() : "", p != null ? p.unidad() : ""));
                            }
                        } else {
                            ProductoStock p = productosStock.get(skuItem);
                            carrosItems.add(new CarrosItem(skuItem, v.getCantidad(),
                                    p != null ? p.producto() : "", p != null ? p.unidad() : ""));
                        }
                    }
                }
            }
            carrosOrdenes.add(new CarrosOrden(numeroVenta, fechaCreacion, letra, carrosItems));
        }
        AppLogger.info("PICKIT - Ordenes CARROS: " + carrosOrdenes.size());

        // Paso 12: Generar Excel
        AppLogger.info("PICKIT - Paso 12: Generando Excel Pickit...");
        File resultado = PickitExcelWriter.generar(pickitItems, carrosOrdenes, slaOrdenes, soloHoy);

        int skusOk = skuCantidad.size() - skusNoEncontrados - skusStockInsuficiente - skusConError;
        int countManuales = (productosManuales != null) ? productosManuales.size() : 0;

        int totalOrdenes = countMLPrint + countMLAgreement + countNube + countGastro + countManuales;

        AppLogger.success("PICKIT - ========== RESUMEN ==========");
        AppLogger.success("PICKIT -   ML ready_to_print: " + countMLPrint + " órdenes");
        AppLogger.success("PICKIT -   ML acuerdo: " + countMLAgreement + " órdenes");
        AppLogger.success("PICKIT -   KT HOGAR: " + countNube + " órdenes");
        AppLogger.success("PICKIT -   KT GASTRO: " + countGastro + " órdenes");
        if (countManuales > 0) AppLogger.success("PICKIT -   Manuales: " + countManuales);
        AppLogger.success("PICKIT -   Total: " + totalOrdenes + " órdenes | " + todasLasVentas.size() + " productos");
        AppLogger.success("PICKIT -   SKUs únicos: " + skuCantidad.size() + " | OK: " + skusOk);
        AppLogger.success("PICKIT -   Ordenes CARROS: " + carrosOrdenes.size());
        if (skusNoEncontrados > 0) AppLogger.warn("PICKIT -   SKUs no encontrados en Stock: " + skusNoEncontrados);
        if (skusStockInsuficiente > 0) AppLogger.warn("PICKIT -   SKUs con stock insuficiente: " + skusStockInsuficiente);
        if (skusConError > 0) AppLogger.warn("PICKIT -   SKUs con error: " + skusConError);
        AppLogger.success("PICKIT - ==============================");

        AppLogger.success("PICKIT - Proceso completado. Archivo: " + resultado.getAbsolutePath());
        return resultado;
    }

    private static String generarLetraCarro(int index) {
        StringBuilder sb = new StringBuilder();
        index++;
        while (index > 0) {
            index--;
            sb.insert(0, (char) ('A' + index % 26));
            index /= 26;
        }
        return sb.toString();
    }

    public static boolean esSkuConError(String sku) {
        return sku.equals("SIN SKU") ||
                sku.startsWith("SIN SKU:") ||
                sku.startsWith("SKU INVALIDO:") ||
                sku.startsWith("CANT INVALIDA:") ||
                sku.startsWith("COMBO INVALIDO:");
    }
}
