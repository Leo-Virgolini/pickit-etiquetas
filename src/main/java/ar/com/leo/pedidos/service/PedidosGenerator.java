package ar.com.leo.pedidos.service;

import ar.com.leo.AppLogger;
import ar.com.leo.api.nube.TiendaNubeApi;
import ar.com.leo.api.nube.TiendaNubeApi.TiendaNubeOrderResult;
import ar.com.leo.api.ml.MercadoLibreAPI;
import ar.com.leo.pedidos.excel.PedidosExcelWriter;
import ar.com.leo.pedidos.model.*;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class PedidosGenerator {

    private static final ExecutorService executor = Executors.newFixedThreadPool(3);

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

    public static File generarPedidos() throws Exception {

        // Paso 1: Inicializar APIs
        AppLogger.info("PEDIDOS - Paso 1: Inicializando APIs (MercadoLibre + Tienda Nube)...");
        if (!MercadoLibreAPI.inicializar()) {
            throw new RuntimeException("No se pudieron inicializar los tokens de MercadoLibre.");
        }
        final String userId = MercadoLibreAPI.getUserId();

        if (!TiendaNubeApi.inicializar()) {
            throw new RuntimeException("No se pudieron inicializar las credenciales de Tienda Nube.");
        }

        // Paso 2-4: Obtener datos en paralelo
        var futureML = executor.submit(() -> {
            AppLogger.info("PEDIDOS - Paso 2: Obteniendo pedidos ML retiro...");
            List<PedidoML> result = MercadoLibreAPI.obtenerPedidosRetiro(userId);
            long ordenes = result.stream().map(PedidoML::orderId).distinct().count();
            AppLogger.info("PEDIDOS - ML retiro: " + ordenes + " órdenes");
            return result;
        });

        var futureTNHogar = executor.submit(() -> {
            AppLogger.info("PEDIDOS - Paso 3: Obteniendo pedidos KT HOGAR (Tienda Nube)...");
            TiendaNubeOrderResult result = TiendaNubeApi.obtenerPedidosCompletosHogar();
            AppLogger.info("PEDIDOS - KT HOGAR: " + result.totalOrdenes() + " órdenes, " + result.etiquetas().size() + " etiquetas");
            return result;
        });

        var futureTNGastro = executor.submit(() -> {
            AppLogger.info("PEDIDOS - Paso 4: Obteniendo pedidos KT GASTRO (Tienda Nube)...");
            TiendaNubeOrderResult result = TiendaNubeApi.obtenerPedidosCompletosGastro();
            AppLogger.info("PEDIDOS - KT GASTRO: " + result.totalOrdenes() + " órdenes, " + result.etiquetas().size() + " etiquetas");
            return result;
        });

        List<PedidoML> pedidosML = futureML.get();
        TiendaNubeOrderResult tnHogar = futureTNHogar.get();
        TiendaNubeOrderResult tnGastro = futureTNGastro.get();

        // Paso 5: Consolidar resultados TN
        AppLogger.info("PEDIDOS - Paso 5: Consolidando resultados...");
        List<PedidoTN> pedidosTN = new ArrayList<>(tnHogar.pedidos());
        pedidosTN.addAll(tnGastro.pedidos());

        List<EtiquetaTN> etiquetasTN = new ArrayList<>(tnHogar.etiquetas());
        etiquetasTN.addAll(tnGastro.etiquetas());

        int totalOrdenesTN = tnHogar.totalOrdenes() + tnGastro.totalOrdenes();
        long ordenesMLRetiro = pedidosML.stream().map(PedidoML::orderId).distinct().count();
        AppLogger.info(String.format("PEDIDOS - Totales: ML retiro=%d | TN órdenes=%d | TN etiquetas=%d",
                ordenesMLRetiro, totalOrdenesTN, etiquetasTN.size()));

        if (pedidosML.isEmpty() && pedidosTN.isEmpty() && etiquetasTN.isEmpty()) {
            AppLogger.warn("PEDIDOS - No se encontraron pedidos para procesar.");
            return null;
        }

        // Paso 6: Generar Excel
        AppLogger.info("PEDIDOS - Paso 6: Generando Excel...");
        PedidosResult result = new PedidosResult(pedidosML, pedidosTN, etiquetasTN);
        File archivo = PedidosExcelWriter.generar(result);

        AppLogger.success("PEDIDOS - ========== RESUMEN ==========");
        AppLogger.success("PEDIDOS -   ML retiro: " + ordenesMLRetiro + " órdenes");
        AppLogger.success("PEDIDOS -   KT HOGAR: " + tnHogar.totalOrdenes() + " órdenes");
        AppLogger.success("PEDIDOS -   KT GASTRO: " + tnGastro.totalOrdenes() + " órdenes");
        AppLogger.success("PEDIDOS -   Etiquetas LLEGA HOY: " + etiquetasTN.size());
        AppLogger.success("PEDIDOS - ==============================");
        AppLogger.success("PEDIDOS - Proceso completado. Archivo: " + archivo.getAbsolutePath());
        return archivo;
    }
}
