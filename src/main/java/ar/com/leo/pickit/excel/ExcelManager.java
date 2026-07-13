package ar.com.leo.pickit.excel;

import ar.com.leo.AppLogger;
import ar.com.leo.pickit.model.ProductoManual;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.*;

public class ExcelManager {

    public record ComboEntry(String skuComponente, double cantidad) {}
    public record ProductoStock(String producto, String subRubro, String unidad, String proveedor, int stock) {}

    private static final int FILA_HEADERS = 2;

    public static Map<String, List<ComboEntry>> obtenerCombos(File combosExcel) throws Exception {
        Map<String, List<ComboEntry>> combos = new LinkedHashMap<>();

        // readOnly=true: los .xls (HSSF) se abren con RandomAccessFile, que por
        // default es modo read-write y falla con "Permission denied" cuando el
        // archivo está en un volumen montado :ro (caso del showroom en el VPS,
        // donde Combos.xls llega read-only vía rclone). Con readOnly se abre en
        // modo "r". Inocuo para el uso manual (solo se lee el archivo).
        try (Workbook workbook = WorkbookFactory.create(combosExcel, null, true)) {
            Sheet hoja = workbook.getSheetAt(0);
            if (hoja == null) throw new Exception("No se encontró ninguna hoja en " + combosExcel.getName());

            Map<String, Integer> headers = obtenerIndicesHeaders(hoja, FILA_HEADERS);
            int colSkuCombo = obtenerIndiceHeader(headers, "Código Compuesto", combosExcel.getName());
            int colSkuComponente = obtenerIndiceHeader(headers, "Código Componente", combosExcel.getName());
            int colCantidad = obtenerIndiceHeader(headers, "Cantidad", combosExcel.getName());

            for (int i = FILA_HEADERS + 1; i <= hoja.getLastRowNum(); i++) {
                Row fila = hoja.getRow(i);
                if (isEmptyRow(fila)) continue;

                Cell celdaSkuCombo = fila.getCell(colSkuCombo);
                Cell celdaSkuComponente = fila.getCell(colSkuComponente);
                Cell celdaCantidad = fila.getCell(colCantidad);

                if (celdaSkuCombo == null || celdaSkuComponente == null || celdaCantidad == null) continue;

                String skuCombo = getCellValue(celdaSkuCombo).trim();
                String skuComponente = getCellValue(celdaSkuComponente).trim();
                String cantidadStr = getCellValue(celdaCantidad).trim();

                if (skuCombo.isEmpty() || skuComponente.isEmpty() || cantidadStr.isEmpty()) continue;

                double cantidad;
                try {
                    cantidad = Double.parseDouble(cantidadStr);
                } catch (NumberFormatException e) {
                    AppLogger.warn("EXCEL - Error al parsear cantidad combo en fila " + (i + 1) + ": " + cantidadStr);
                    continue;
                }

                if (cantidad <= 0) {
                    AppLogger.warn("EXCEL - Cantidad inválida en combo fila " + (i + 1) + ": " + cantidadStr);
                    continue;
                }

                combos.computeIfAbsent(skuCombo, k -> new ArrayList<>())
                        .add(new ComboEntry(skuComponente, cantidad));
            }
        }

        AppLogger.info("EXCEL - Combos leídos: " + combos.size());
        return combos;
    }

    public static Map<String, ProductoStock> obtenerProductosStock(File stockExcel) throws Exception {
        Map<String, ProductoStock> productos = new LinkedHashMap<>();

        try (OPCPackage opcPackage = OPCPackage.open(stockExcel, PackageAccess.READ);
             Workbook workbook = new XSSFWorkbook(opcPackage)) {

            Sheet hoja = workbook.getSheetAt(0);
            if (hoja == null) throw new Exception("No se encontró ninguna hoja en " + stockExcel.getName());

            Map<String, Integer> headers = obtenerIndicesHeaders(hoja, FILA_HEADERS);
            int colSku = obtenerIndiceHeader(headers, "Código Producto", stockExcel.getName());
            int colProducto = obtenerIndiceHeader(headers, "Producto", stockExcel.getName());
            int colSubRubro = obtenerIndiceHeader(headers, "Sub Rubro", stockExcel.getName());
            int colStock = obtenerIndiceHeader(headers, "Stock Disponible", stockExcel.getName());
            int colUnidad = obtenerIndiceHeader(headers, "Unidad", stockExcel.getName());
            int colProveedor = obtenerIndiceHeader(headers, "Proveedor", stockExcel.getName());

            for (int i = FILA_HEADERS + 1; i <= hoja.getLastRowNum(); i++) {
                Row fila = hoja.getRow(i);
                if (isEmptyRow(fila)) continue;

                Cell celdaSku = fila.getCell(colSku);
                if (celdaSku == null) continue;

                String sku = getCellValue(celdaSku).trim();
                if (sku.isEmpty()) continue;

                String producto = fila.getCell(colProducto) != null ? getCellValue(fila.getCell(colProducto)).trim() : "";
                String subRubro = fila.getCell(colSubRubro) != null ? getCellValue(fila.getCell(colSubRubro)).trim() : "";
                String unidad = fila.getCell(colUnidad) != null ? getCellValue(fila.getCell(colUnidad)).trim() : "";
                String proveedor = fila.getCell(colProveedor) != null ? getCellValue(fila.getCell(colProveedor)).trim() : "";

                int stock = 0;
                Cell celdaStock = fila.getCell(colStock);
                if (celdaStock != null) {
                    String stockStr = getCellValue(celdaStock).trim();
                    if (!stockStr.isEmpty()) {
                        try { stock = (int) Double.parseDouble(stockStr); } catch (NumberFormatException ignored) {}
                    }
                }

                productos.put(sku, new ProductoStock(producto, subRubro, unidad, proveedor, stock));
            }
        }

        AppLogger.info("EXCEL - Productos leídos: " + productos.size());
        return productos;
    }

    public static List<ProductoManual> obtenerProductosManualesDesdeExcel(File excel) throws Exception {
        Map<String, Double> productosMap = new LinkedHashMap<>();

        try (OPCPackage opcPackage = OPCPackage.open(excel, PackageAccess.READ);
             Workbook workbook = new XSSFWorkbook(opcPackage)) {
            Sheet hoja = workbook.getSheetAt(0);
            if (hoja == null) throw new Exception("No se encontró ninguna hoja en " + excel.getName());

            int filaHeaders = buscarFilaHeaders(hoja, "SKU");
            Map<String, Integer> headers = obtenerIndicesHeaders(hoja, filaHeaders);
            int colSku = obtenerIndiceHeader(headers, "SKU", excel.getName());
            int colCantidad = obtenerIndiceHeader(headers, "CANTIDAD", excel.getName());

            for (int i = filaHeaders + 1; i <= hoja.getLastRowNum(); i++) {
                Row fila = hoja.getRow(i);
                if (isEmptyRow(fila)) continue;

                Cell celdaSku = fila.getCell(colSku);
                Cell celdaCantidad = fila.getCell(colCantidad);
                if (celdaSku == null || celdaCantidad == null) continue;

                String sku = getCellValue(celdaSku).trim();
                String cantidadStr = getCellValue(celdaCantidad).trim();
                if (sku.isEmpty() || cantidadStr.isEmpty()) continue;

                if (!sku.matches("\\d+")) {
                    AppLogger.warn("EXCEL IMPORT - SKU no numérico en fila " + (i + 1) + ": " + sku + " (omitido)");
                    continue;
                }

                double cantidad;
                try {
                    cantidad = Double.parseDouble(cantidadStr);
                } catch (NumberFormatException e) {
                    AppLogger.warn("EXCEL IMPORT - Error al parsear cantidad en fila " + (i + 1) + ": " + cantidadStr);
                    continue;
                }

                if (cantidad <= 0) {
                    AppLogger.warn("EXCEL IMPORT - Cantidad inválida en fila " + (i + 1) + ": " + cantidadStr);
                    continue;
                }

                productosMap.merge(sku, cantidad, Double::sum);
            }
        }

        List<ProductoManual> productos = new ArrayList<>();
        for (Map.Entry<String, Double> entry : productosMap.entrySet()) {
            productos.add(new ProductoManual(entry.getKey(), entry.getValue()));
        }

        AppLogger.info("EXCEL IMPORT - Productos importados: " + productos.size());
        return productos;
    }

    private static int buscarFilaHeaders(Sheet hoja, String headerRequerido) throws Exception {
        int maxFilas = Math.min(hoja.getLastRowNum(), 10);
        for (int i = 0; i <= maxFilas; i++) {
            Row row = hoja.getRow(i);
            if (row == null) continue;
            for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    if (cell.getStringCellValue().trim().equalsIgnoreCase(headerRequerido)) {
                        return i;
                    }
                }
            }
        }
        throw new Exception("No se encontró el header '" + headerRequerido + "' en las primeras filas");
    }

    private static Map<String, Integer> obtenerIndicesHeaders(Sheet hoja, int filaHeaders) throws Exception {
        Map<String, Integer> indices = new HashMap<>();
        Row row = hoja.getRow(filaHeaders);
        if (row == null) throw new Exception("No se encontró la fila de headers (fila " + (filaHeaders + 1) + ")");
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String valor = cell.getStringCellValue().trim();
                if (!valor.isEmpty()) indices.put(valor, i);
            }
        }
        return indices;
    }

    private static int obtenerIndiceHeader(Map<String, Integer> indices, String nombreHeader, String archivo) throws Exception {
        for (Map.Entry<String, Integer> entry : indices.entrySet()) {
            if (entry.getKey().equalsIgnoreCase(nombreHeader)) return entry.getValue();
        }
        throw new Exception("No se encontró el header '" + nombreHeader + "' en " + archivo + ". Headers disponibles: " + indices.keySet());
    }

    private static boolean isEmptyRow(Row row) {
        if (row == null) return true;
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().isBlank()) continue;
                return false;
            }
        }
        return true;
    }

    private static String getCellValue(Cell cell) throws Exception {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) yield cell.getDateCellValue().toString();
                double num = cell.getNumericCellValue();
                yield num == Math.floor(num) ? String.valueOf((long) num) : String.valueOf(num);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> getCellFormulaValue(cell);
            case ERROR -> throw new Exception("EXCEL - Error en la celda fila: " + (cell.getAddress().getRow() + 1) + " columna: " + (cell.getAddress().getColumn() + 1));
            default -> "";
        };
    }

    private static String getCellFormulaValue(Cell cell) {
        return switch (cell.getCachedFormulaResultType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) yield cell.getDateCellValue().toString();
                double num = cell.getNumericCellValue();
                yield num == Math.floor(num) ? String.valueOf((long) num) : String.valueOf(num);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case ERROR -> {
                AppLogger.warn("EXCEL - Fórmula con error en celda " + cell.getAddress());
                yield "0";
            }
            default -> "";
        };
    }
}
