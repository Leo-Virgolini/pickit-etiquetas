package ar.com.leo.pedidos.excel;

import ar.com.leo.pedidos.model.EtiquetaTN;
import ar.com.leo.pedidos.model.PedidoML;
import ar.com.leo.pedidos.model.PedidoTN;
import ar.com.leo.pedidos.model.PedidosResult;
import ar.com.leo.util.Util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.ToIntFunction;
import java.util.stream.Collectors;

public class PedidosExcelWriter {

    private static final DateTimeFormatter FECHA_EXCEL = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm");
    private static final ZoneId ZONA_AR = ZoneId.of("America/Argentina/Buenos_Aires");

    private static final int CARD_COLS = 6;
    private static final int RIGHT_START = 7;

    private static final int ETIQUETA_CARD_ROWS = 8;
    private static final int PEDIDO_MIN_PRODUCTS = 1;

    // Alturas para tarjetas de pedido (filas header fijas + productos dinámicos + padding)
    private static final float[] PEDIDO_HEADER_HEIGHTS = {26, 26, 26, 20};
    private static final float PEDIDO_PRODUCT_HEIGHT = 33f;
    private static final float PEDIDO_PADDING_HEIGHT = 4f;

    private static final float[] ETIQUETA_ROW_HEIGHTS = {18, 32, 28, 24, 22, 22, 8, 8};

    // Anchos de columna en 1/256 de carácter
    private static final int COL_SKU_WIDTH = 10 * 256;
    private static final int COL_CANT_WIDTH = 7 * 256;
    private static final int COL_DETALLE_WIDTH = 12 * 256;
    private static final int COL_GAP_WIDTH = (int) (1.5 * 256);

    // Altura máxima de página en puntos (A4 portrait - márgenes)
    private static final float MAX_PAGE_HEIGHT = 810f;

    // ── Records para agrupación ──

    private record ProductLine(String sku, double cantidad, String detalle) {}

    private record PedidoMLGroup(long orderId, OffsetDateTime fecha, String usuario,
                                  String nombreApellido, List<ProductLine> productos) {}

    private record PedidoTNGroup(long orderId, OffsetDateTime fecha, String nombreApellido,
                                  String tienda, String tipoEnvio, List<ProductLine> productos) {}

    // ── Contenedor de estilos ──

    private static class CardStyles {
        final XSSFCellStyle bigNumber;
        final XSSFCellStyle fecha;
        final XSSFCellStyle nombre;
        final XSSFCellStyle tiendaBadge;
        final XSSFCellStyle productHeader;
        final XSSFCellStyle productNormal;
        final XSSFCellStyle productHighlight;
        final XSSFCellStyle productWrap;
        final XSSFCellStyle productHighlightWrap;
        final XSSFCellStyle empty;
        final XSSFCellStyle smallLabel;
        final XSSFCellStyle smallLabelRight;
        final XSSFCellStyle bigNombre;
        final XSSFCellStyle domicilio;
        final XSSFCellStyle localidad;
        final XSSFCellStyle cpTelefono;
        final XSSFCellStyle observaciones;

        CardStyles(XSSFWorkbook wb, XSSFColor accentColor) {
            // N° venta grande (18pt bold)
            bigNumber = wb.createCellStyle();
            bigNumber.setFont(createXFont(wb, 18, true, false));
            bigNumber.setAlignment(HorizontalAlignment.CENTER);
            bigNumber.setVerticalAlignment(VerticalAlignment.CENTER);
            setBordersOn(bigNumber, BorderStyle.THIN);

            // Fecha (11pt, fondo gris claro)
            fecha = wb.createCellStyle();
            fecha.setFont(createXFont(wb, 11, false, false));
            fecha.setAlignment(HorizontalAlignment.CENTER);
            fecha.setVerticalAlignment(VerticalAlignment.CENTER);
            fecha.setFillForegroundColor(xcolor(230, 230, 230));
            fecha.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            setBordersOn(fecha, BorderStyle.THIN);

            // Nombre (14pt bold)
            nombre = wb.createCellStyle();
            nombre.setFont(createXFont(wb, 14, true, false));
            nombre.setAlignment(HorizontalAlignment.LEFT);
            nombre.setVerticalAlignment(VerticalAlignment.CENTER);
            nombre.setIndention((short) 1);
            setBordersOn(nombre, BorderStyle.THIN);

            // Tienda badge (12pt bold, fondo oscuro, texto blanco)
            tiendaBadge = wb.createCellStyle();
            XSSFFont fontTienda = createXFont(wb, 12, true, false);
            fontTienda.setColor(xcolor(255, 255, 255));
            tiendaBadge.setFont(fontTienda);
            tiendaBadge.setAlignment(HorizontalAlignment.CENTER);
            tiendaBadge.setVerticalAlignment(VerticalAlignment.CENTER);
            tiendaBadge.setFillForegroundColor(xcolor(60, 60, 60));
            tiendaBadge.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            tiendaBadge.setWrapText(true);
            setBordersOn(tiendaBadge, BorderStyle.THIN);

            // Headers de producto (10pt bold, fondo accent)
            productHeader = wb.createCellStyle();
            productHeader.setFont(createXFont(wb, 10, true, false));
            productHeader.setAlignment(HorizontalAlignment.CENTER);
            productHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            productHeader.setFillForegroundColor(accentColor);
            productHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            setBordersOn(productHeader, BorderStyle.THIN);

            // Producto normal (11pt)
            productNormal = wb.createCellStyle();
            productNormal.setFont(createXFont(wb, 11, false, false));
            productNormal.setAlignment(HorizontalAlignment.CENTER);
            productNormal.setVerticalAlignment(VerticalAlignment.CENTER);
            setBordersOn(productNormal, BorderStyle.THIN);

            // Producto highlight (11pt bold, fondo amarillo)
            productHighlight = wb.createCellStyle();
            productHighlight.setFont(createXFont(wb, 11, true, false));
            productHighlight.setAlignment(HorizontalAlignment.CENTER);
            productHighlight.setVerticalAlignment(VerticalAlignment.CENTER);
            productHighlight.setFillForegroundColor(xcolor(255, 255, 153));
            productHighlight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            setBordersOn(productHighlight, BorderStyle.THIN);

            // Producto wrap (11pt, wrap text, alineado izq)
            productWrap = wb.createCellStyle();
            productWrap.cloneStyleFrom(productNormal);
            productWrap.setWrapText(true);
            productWrap.setAlignment(HorizontalAlignment.LEFT);

            // Producto highlight wrap
            productHighlightWrap = wb.createCellStyle();
            productHighlightWrap.cloneStyleFrom(productHighlight);
            productHighlightWrap.setWrapText(true);
            productHighlightWrap.setAlignment(HorizontalAlignment.LEFT);

            // Vacío
            empty = wb.createCellStyle();

            // ── Estilos para etiquetas ──

            // Label chico (9pt gris)
            smallLabel = wb.createCellStyle();
            XSSFFont fontSmall = createXFont(wb, 9, false, false);
            fontSmall.setColor(xcolor(128, 128, 128));
            smallLabel.setFont(fontSmall);
            smallLabel.setAlignment(HorizontalAlignment.LEFT);
            smallLabel.setVerticalAlignment(VerticalAlignment.CENTER);
            smallLabel.setIndention((short) 1);

            // Label chico alineado a la derecha
            smallLabelRight = wb.createCellStyle();
            smallLabelRight.cloneStyleFrom(smallLabel);
            smallLabelRight.setAlignment(HorizontalAlignment.RIGHT);
            smallLabelRight.setIndention((short) 1);

            // Nombre grande (16pt bold centrado)
            bigNombre = wb.createCellStyle();
            bigNombre.setFont(createXFont(wb, 16, true, false));
            bigNombre.setAlignment(HorizontalAlignment.CENTER);
            bigNombre.setVerticalAlignment(VerticalAlignment.CENTER);

            // Domicilio (13pt bold)
            domicilio = wb.createCellStyle();
            domicilio.setFont(createXFont(wb, 13, true, false));
            domicilio.setAlignment(HorizontalAlignment.CENTER);
            domicilio.setVerticalAlignment(VerticalAlignment.CENTER);
            domicilio.setWrapText(true);

            // Localidad (13pt)
            localidad = wb.createCellStyle();
            localidad.setFont(createXFont(wb, 13, false, false));
            localidad.setAlignment(HorizontalAlignment.CENTER);
            localidad.setVerticalAlignment(VerticalAlignment.CENTER);

            // CP/Teléfono (12pt bold)
            cpTelefono = wb.createCellStyle();
            cpTelefono.setFont(createXFont(wb, 12, true, false));
            cpTelefono.setAlignment(HorizontalAlignment.CENTER);
            cpTelefono.setVerticalAlignment(VerticalAlignment.CENTER);

            // Observaciones (10pt italic)
            observaciones = wb.createCellStyle();
            observaciones.setFont(createXFont(wb, 10, false, true));
            observaciones.setAlignment(HorizontalAlignment.LEFT);
            observaciones.setVerticalAlignment(VerticalAlignment.CENTER);
            observaciones.setWrapText(true);
            observaciones.setIndention((short) 1);
        }

        private static XSSFFont createXFont(XSSFWorkbook wb, int size, boolean bold, boolean italic) {
            XSSFFont f = wb.createFont();
            f.setFontName("Calibri");
            f.setFontHeightInPoints((short) size);
            f.setBold(bold);
            f.setItalic(italic);
            return f;
        }

        private static XSSFColor xcolor(int r, int g, int b) {
            return new XSSFColor(new byte[]{(byte) r, (byte) g, (byte) b}, null);
        }

        private static void setBordersOn(CellStyle style, BorderStyle border) {
            style.setBorderTop(border);
            style.setBorderBottom(border);
            style.setBorderLeft(border);
            style.setBorderRight(border);
        }
    }

    // ── Punto de entrada ──

    public static File generar(PedidosResult result) throws Exception {
        Path excelDir = Paths.get(Util.getJarFolder(), "Pedidos");
        Files.createDirectories(excelDir);

        LocalDateTime ahora = LocalDateTime.now();
        String fecha = ahora.format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"));
        File outputFile = excelDir.resolve("PEDIDOS_" + fecha + ".xlsx").toFile();

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            if (!result.pedidosML().isEmpty())
                generarHojaMLRetiro(workbook, result.pedidosML());
            if (!result.pedidosTN().isEmpty())
                generarHojaTNPedidos(workbook, result.pedidosTN());
            if (!result.etiquetasTN().isEmpty())
                generarHojaTNEtiquetas(workbook, result.etiquetasTN());

            if (workbook.getNumberOfSheets() == 0) {
                Sheet empty = workbook.createSheet("SIN DATOS");
                empty.createRow(0).createCell(0).setCellValue("No se encontraron pedidos para procesar.");
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }
        }

        return outputFile;
    }

    // ── Agrupación de pedidos ──

    private static List<PedidoMLGroup> agruparPedidosML(List<PedidoML> pedidos) {
        Map<Long, List<PedidoML>> grouped = pedidos.stream()
                .collect(Collectors.groupingBy(PedidoML::orderId, LinkedHashMap::new, Collectors.toList()));

        return grouped.values().stream().map(list -> {
            PedidoML first = list.getFirst();
            List<ProductLine> productos = list.stream()
                    .map(p -> new ProductLine(p.sku(), p.cantidad(), p.detalle()))
                    .toList();
            return new PedidoMLGroup(first.orderId(), first.fecha(), first.usuario(),
                    first.nombreApellido(), productos);
        }).toList();
    }

    private static List<PedidoTNGroup> agruparPedidosTN(List<PedidoTN> pedidos) {
        Map<Long, List<PedidoTN>> grouped = pedidos.stream()
                .collect(Collectors.groupingBy(PedidoTN::orderId, LinkedHashMap::new, Collectors.toList()));

        return grouped.values().stream().map(list -> {
            PedidoTN first = list.getFirst();
            List<ProductLine> productos = list.stream()
                    .map(p -> new ProductLine(p.sku(), p.cantidad(), p.detalle()))
                    .toList();
            return new PedidoTNGroup(first.orderId(), first.fecha(), first.nombreApellido(),
                    first.tienda(), first.tipoEnvio(), productos);
        }).toList();
    }

    // ── Generadores de hoja ──

    private static void generarHojaMLRetiro(XSSFWorkbook wb, List<PedidoML> pedidos) {
        Sheet sheet = wb.createSheet("ML PEDIDOS RETIRO");
        CardStyles styles = new CardStyles(wb, CardStyles.xcolor(186, 154, 215));
        configurarColumnas(sheet);

        List<PedidoMLGroup> groups = agruparPedidosML(pedidos);
        colocarTarjetasPedidos(groups, sheet,
                g -> g.productos().size(),
                (s, row, col, g, slots) -> renderCardMLRetiro(s, row, col, g, styles, slots));

        configurarImpresion(sheet);
    }

    private static void generarHojaTNPedidos(XSSFWorkbook wb, List<PedidoTN> pedidos) {
        Sheet sheet = wb.createSheet("TN PEDIDOS");
        CardStyles styles = new CardStyles(wb, CardStyles.xcolor(146, 208, 80));
        configurarColumnas(sheet);

        List<PedidoTNGroup> groups = agruparPedidosTN(pedidos);
        colocarTarjetasPedidos(groups, sheet,
                g -> g.productos().size(),
                (s, row, col, g, slots) -> renderCardTNPedido(s, row, col, g, styles, slots));

        configurarImpresion(sheet);
    }

    private static void generarHojaTNEtiquetas(XSSFWorkbook wb, List<EtiquetaTN> etiquetas) {
        Sheet sheet = wb.createSheet("TN ETIQUETAS");
        CardStyles styles = new CardStyles(wb, CardStyles.xcolor(255, 165, 0));
        configurarColumnas(sheet);

        colocarTarjetasEtiquetas(etiquetas, sheet,
                (s, row, col, e) -> renderCardTNEtiqueta(s, row, col, e, styles));

        configurarImpresion(sheet);
    }

    // ── Orquestación de layout ──

    @FunctionalInterface
    private interface PedidoCardRenderer<T> {
        void render(Sheet sheet, int startRow, int startCol, T item, int productSlots);
    }

    @FunctionalInterface
    private interface CardRenderer<T> {
        void render(Sheet sheet, int startRow, int startCol, T item);
    }

    /**
     * Coloca tarjetas de pedido con altura dinámica según la cantidad de productos.
     * Los pares de tarjetas (izq/der) comparten la misma altura = max de ambas.
     * Page breaks inteligentes basados en altura acumulada.
     */
    private static <T> void colocarTarjetasPedidos(List<T> items, Sheet sheet,
                                                    ToIntFunction<T> productCountFn,
                                                    PedidoCardRenderer<T> renderer) {
        int currentRow = 0;
        float currentPageHeight = 0;

        for (int i = 0; i < items.size(); i += 2) {
            T left = items.get(i);
            T right = (i + 1 < items.size()) ? items.get(i + 1) : null;

            int maxProducts = productCountFn.applyAsInt(left);
            if (right != null) maxProducts = Math.max(maxProducts, productCountFn.applyAsInt(right));
            int slots = Math.max(maxProducts, PEDIDO_MIN_PRODUCTS);
            int cardRows = PEDIDO_HEADER_HEIGHTS.length + slots + 1;

            // Calcular altura del par
            float pairHeight = PEDIDO_PADDING_HEIGHT;
            for (float h : PEDIDO_HEADER_HEIGHTS) pairHeight += h;
            pairHeight += slots * PEDIDO_PRODUCT_HEIGHT;

            // Page break si no cabe en la página actual
            if (currentPageHeight > 0 && currentPageHeight + pairHeight > MAX_PAGE_HEIGHT) {
                sheet.setRowBreak(currentRow - 1);
                currentPageHeight = 0;
            }

            // Establecer alturas de fila
            for (int r = 0; r < cardRows; r++) {
                Row row = getOrCreateRow(sheet, currentRow + r);
                if (r < PEDIDO_HEADER_HEIGHTS.length)
                    row.setHeightInPoints(PEDIDO_HEADER_HEIGHTS[r]);
                else if (r < PEDIDO_HEADER_HEIGHTS.length + slots)
                    row.setHeightInPoints(PEDIDO_PRODUCT_HEIGHT);
                else
                    row.setHeightInPoints(PEDIDO_PADDING_HEIGHT);
            }

            // Renderizar tarjetas
            renderer.render(sheet, currentRow, 0, left, slots);
            if (right != null) {
                renderer.render(sheet, currentRow, RIGHT_START, right, slots);
            }

            currentRow += cardRows;
            currentPageHeight += pairHeight;
        }
    }

    /**
     * Coloca tarjetas de etiqueta con altura fija.
     * Page breaks inteligentes basados en altura acumulada.
     */
    private static <T> void colocarTarjetasEtiquetas(List<T> items, Sheet sheet,
                                                      CardRenderer<T> renderer) {
        float etiquetaHeight = 0;
        for (float h : ETIQUETA_ROW_HEIGHTS) etiquetaHeight += h;

        int currentRow = 0;
        float currentPageHeight = 0;

        for (int i = 0; i < items.size(); i += 2) {
            // Page break si no cabe otro par
            if (currentPageHeight > 0 && currentPageHeight + etiquetaHeight > MAX_PAGE_HEIGHT) {
                sheet.setRowBreak(currentRow - 1);
                currentPageHeight = 0;
            }

            // Establecer alturas de fila
            for (int r = 0; r < ETIQUETA_CARD_ROWS; r++) {
                Row row = getOrCreateRow(sheet, currentRow + r);
                row.setHeightInPoints(ETIQUETA_ROW_HEIGHTS[r]);
            }

            // Renderizar izquierda
            renderer.render(sheet, currentRow, 0, items.get(i));

            // Renderizar derecha si hay
            if (i + 1 < items.size()) {
                renderer.render(sheet, currentRow, RIGHT_START, items.get(i + 1));
            }

            currentRow += ETIQUETA_CARD_ROWS;
            currentPageHeight += etiquetaHeight;
        }
    }

    // ── Renderizado de tarjetas ──

    private static void renderCardMLRetiro(Sheet sheet, int startRow, int startCol,
                                            PedidoMLGroup group, CardStyles styles, int productSlots) {
        int endCol = startCol + CARD_COLS - 1;

        // Filas 0-1: N° VENTA (2×3) + FECHA (2×3) — sin celdas vacías
        mergeAndSet(sheet, startRow, startRow + 1, startCol, startCol + 2,
                String.valueOf(group.orderId()), styles.bigNumber);
        mergeAndSet(sheet, startRow, startRow + 1, startCol + 3, endCol,
                formatFecha(group.fecha()), styles.fecha);

        // Fila 2: Nombre (usuario)
        String nombreTexto;
        if (group.nombreApellido() != null && !group.nombreApellido().isBlank()) {
            nombreTexto = group.nombreApellido();
            if (group.usuario() != null && !group.usuario().isBlank())
                nombreTexto += "  (" + group.usuario() + ")";
        } else {
            nombreTexto = group.usuario() != null ? group.usuario() : "";
        }
        mergeAndSet(sheet, startRow + 2, startRow + 2, startCol, endCol, nombreTexto, styles.nombre);

        // Fila 3: Headers de producto
        renderProductHeaders(sheet, startRow + 3, startCol, endCol, styles);

        // Filas 4..4+slots-1: Productos
        renderProductRows(sheet, startRow, startCol, endCol, group.productos(), styles, productSlots);

        // Última fila: padding
        int paddingRow = startRow + 4 + productSlots;
        fillCells(sheet, paddingRow, paddingRow, startCol, endCol, styles.empty);

        // Bordes gruesos (filas 0 a 3+productSlots, sin incluir padding)
        aplicarBordesTarjeta(sheet, startRow, startRow + 3 + productSlots, startCol, endCol);
    }

    private static void renderCardTNPedido(Sheet sheet, int startRow, int startCol,
                                            PedidoTNGroup group, CardStyles styles, int productSlots) {
        int endCol = startCol + CARD_COLS - 1;

        // Filas 0-1: N° VENTA (2×2) + TIENDA (2×2) + FECHA (2×2)
        mergeAndSet(sheet, startRow, startRow + 1, startCol, startCol + 1,
                String.valueOf(group.orderId()), styles.bigNumber);
        String badgeText = safe(group.tienda());
        if (group.tipoEnvio() != null && !group.tipoEnvio().isBlank())
            badgeText += "\n(" + group.tipoEnvio() + ")";
        mergeAndSet(sheet, startRow, startRow + 1, startCol + 2, startCol + 3,
                badgeText, styles.tiendaBadge);
        mergeAndSet(sheet, startRow, startRow + 1, startCol + 4, endCol,
                formatFecha(group.fecha()), styles.fecha);

        // Fila 2: Nombre
        mergeAndSet(sheet, startRow + 2, startRow + 2, startCol, endCol,
                group.nombreApellido(), styles.nombre);

        // Fila 3: Headers de producto
        renderProductHeaders(sheet, startRow + 3, startCol, endCol, styles);

        // Filas 4..4+slots-1: Productos
        renderProductRows(sheet, startRow, startCol, endCol, group.productos(), styles, productSlots);

        // Última fila: padding
        int paddingRow = startRow + 4 + productSlots;
        fillCells(sheet, paddingRow, paddingRow, startCol, endCol, styles.empty);

        // Bordes gruesos
        aplicarBordesTarjeta(sheet, startRow, startRow + 3 + productSlots, startCol, endCol);
    }

    private static void renderCardTNEtiqueta(Sheet sheet, int startRow, int startCol,
                                              EtiquetaTN etq, CardStyles styles) {
        int endCol = startCol + CARD_COLS - 1;

        // Fila 0: Venta + Fecha (9pt gris)
        mergeAndSet(sheet, startRow, startRow, startCol, startCol + 2,
                "Venta: " + etq.orderId(), styles.smallLabel);
        mergeAndSet(sheet, startRow, startRow, startCol + 3, endCol,
                formatFecha(etq.fecha()), styles.smallLabelRight);

        // Fila 1: NOMBRE Y APELLIDO (16pt bold centrado)
        mergeAndSet(sheet, startRow + 1, startRow + 1, startCol, endCol,
                etq.nombreApellido(), styles.bigNombre);

        // Fila 2: DOMICILIO (13pt bold)
        mergeAndSet(sheet, startRow + 2, startRow + 2, startCol, endCol,
                etq.domicilio(), styles.domicilio);

        // Fila 3: LOCALIDAD (13pt)
        mergeAndSet(sheet, startRow + 3, startRow + 3, startCol, endCol,
                etq.localidad(), styles.localidad);

        // Fila 4: CP + Teléfono (12pt bold)
        mergeAndSet(sheet, startRow + 4, startRow + 4, startCol, startCol + 2,
                "CP: " + safe(etq.codigoPostal()), styles.cpTelefono);
        mergeAndSet(sheet, startRow + 4, startRow + 4, startCol + 3, endCol,
                "Tel: " + safe(etq.telefono()), styles.cpTelefono);

        // Fila 5: Observaciones (10pt italic)
        mergeAndSet(sheet, startRow + 5, startRow + 5, startCol, endCol,
                safe(etq.observaciones()), styles.observaciones);

        // Filas 6-7: Padding
        fillCells(sheet, startRow + 6, startRow + 7, startCol, endCol, styles.empty);

        // Bordes gruesos (filas 0-5)
        aplicarBordesTarjeta(sheet, startRow, startRow + 5, startCol, endCol);
    }

    // ── Helpers de renderizado de productos ──

    private static void renderProductHeaders(Sheet sheet, int row, int startCol, int endCol,
                                              CardStyles styles) {
        mergeAndSet(sheet, row, row, startCol, startCol + 1, "SKU", styles.productHeader);
        setCellValue(sheet, row, startCol + 2, "CANT", styles.productHeader);
        mergeAndSet(sheet, row, row, startCol + 3, endCol, "DETALLE DEL PRODUCTO", styles.productHeader);
    }

    private static void renderProductRows(Sheet sheet, int startRow, int startCol, int endCol,
                                           List<ProductLine> productos, CardStyles styles,
                                           int productSlots) {
        for (int i = 0; i < productSlots; i++) {
            int row = startRow + 4 + i;
            if (i < productos.size()) {
                ProductLine p = productos.get(i);
                boolean hl = p.cantidad() > 1;
                CellStyle skuStyle = hl ? styles.productHighlight : styles.productNormal;
                CellStyle cantStyle = hl ? styles.productHighlight : styles.productNormal;
                CellStyle detStyle = hl ? styles.productHighlightWrap : styles.productWrap;

                mergeAndSet(sheet, row, row, startCol, startCol + 1, p.sku(), skuStyle);
                setCantidadCell(sheet, row, startCol + 2, p.cantidad(), cantStyle);
                mergeAndSet(sheet, row, row, startCol + 3, endCol, p.detalle(), detStyle);
            } else {
                mergeAndSet(sheet, row, row, startCol, startCol + 1, "", styles.productNormal);
                setCellValue(sheet, row, startCol + 2, "", styles.productNormal);
                mergeAndSet(sheet, row, row, startCol + 3, endCol, "", styles.productNormal);
            }
        }
    }

    // ── Configuración de columnas e impresión ──

    private static void configurarColumnas(Sheet sheet) {
        sheet.setColumnWidth(0, COL_SKU_WIDTH);
        sheet.setColumnWidth(1, COL_SKU_WIDTH);
        sheet.setColumnWidth(2, COL_CANT_WIDTH);
        sheet.setColumnWidth(3, COL_DETALLE_WIDTH);
        sheet.setColumnWidth(4, COL_DETALLE_WIDTH);
        sheet.setColumnWidth(5, COL_DETALLE_WIDTH);
        sheet.setColumnWidth(6, COL_GAP_WIDTH);
        sheet.setColumnWidth(7, COL_SKU_WIDTH);
        sheet.setColumnWidth(8, COL_SKU_WIDTH);
        sheet.setColumnWidth(9, COL_CANT_WIDTH);
        sheet.setColumnWidth(10, COL_DETALLE_WIDTH);
        sheet.setColumnWidth(11, COL_DETALLE_WIDTH);
        sheet.setColumnWidth(12, COL_DETALLE_WIDTH);
    }

    private static void configurarImpresion(Sheet sheet) {
        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(false);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setMargin(Sheet.LeftMargin, 0.25);
        sheet.setMargin(Sheet.RightMargin, 0.25);
        sheet.setMargin(Sheet.TopMargin, 0.2);
        sheet.setMargin(Sheet.BottomMargin, 0.2);
        sheet.setHorizontallyCenter(true);
        sheet.getFooter().setCenter("Página &P de &N");
    }

    // ── Helpers de bajo nivel ──

    private static void mergeAndSet(Sheet sheet, int firstRow, int lastRow,
                                     int firstCol, int lastCol, String value, CellStyle style) {
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = getOrCreateRow(sheet, r);
            for (int c = firstCol; c <= lastCol; c++) {
                Cell cell = row.createCell(c);
                cell.setCellStyle(style);
            }
        }
        getOrCreateRow(sheet, firstRow).getCell(firstCol).setCellValue(value != null ? value : "");
        if (firstRow != lastRow || firstCol != lastCol) {
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }
    }

    private static void setCellValue(Sheet sheet, int rowIndex, int col, String value, CellStyle style) {
        Cell cell = getOrCreateRow(sheet, rowIndex).createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private static void setCantidadCell(Sheet sheet, int rowIndex, int col, double cantidad, CellStyle style) {
        Cell cell = getOrCreateRow(sheet, rowIndex).createCell(col);
        if (cantidad == Math.floor(cantidad)) cell.setCellValue((int) cantidad);
        else cell.setCellValue(cantidad);
        cell.setCellStyle(style);
    }

    private static Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row != null ? row : sheet.createRow(rowIndex);
    }

    private static void fillCells(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol,
                                   CellStyle style) {
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = getOrCreateRow(sheet, r);
            for (int c = firstCol; c <= lastCol; c++) {
                row.createCell(c).setCellStyle(style);
            }
        }
    }

    private static void aplicarBordesTarjeta(Sheet sheet, int firstRow, int lastRow,
                                              int firstCol, int lastCol) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        RegionUtil.setBorderTop(BorderStyle.THICK, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THICK, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THICK, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THICK, region, sheet);
    }

    private static String formatFecha(OffsetDateTime fecha) {
        if (fecha == null) return "";
        return fecha.atZoneSameInstant(ZONA_AR).format(FECHA_EXCEL);
    }

    private static String safe(String value) {
        return value != null ? value : "";
    }
}
