package ar.com.leo.etiquetas.parser;

import ar.com.leo.AppLogger;
import ar.com.leo.etiquetas.model.MedidaSku;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Gestiona el Excel "madre" de medidas de embalaje por SKU.
 * Columnas esperadas:
 *   SKU | PRODUCTO | Ancho cm | Alto cm | Profundidad cm | Peso físico (empaque + producto) kg
 *       | Ancho +20% | Alto +20% | Profunidad +20% | Peso físico (empaque + producto) +20% | SUBIDO
 * Las +20% son las que se suben a la API de ML (dims en cm, peso en kg).
 */
public class MedidasExcelManager {

    public static final String[] HEADERS = {
            "SKU",
            "PRODUCTO",
            "Ancho\ncm",
            "Alto\ncm",
            "Profundidad\ncm",
            "Peso físico\n(empaque + producto)\nkg",
            "Ancho +20%",
            "Alto +20%",
            "Profunidad +20%",
            "Peso físico (empaque + producto) +20%",
            "SUBIDO",
            "ERROR"
    };

    private static final int COL_SKU = 0;
    private static final int COL_PRODUCTO = 1;
    private static final int COL_ANCHO = 2;
    private static final int COL_ALTO = 3;
    private static final int COL_PROFUNDIDAD = 4;
    private static final int COL_PESO = 5;
    private static final int COL_ANCHO_MAS = 6;
    private static final int COL_ALTO_MAS = 7;
    private static final int COL_PROFUNDIDAD_MAS = 8;
    private static final int COL_PESO_MAS = 9;
    private static final int COL_SUBIDO = 10;
    private static final int COL_ERROR = 11;

    // Retry para sharing violation (Excel abierto por el usuario).
    private static final int MAX_WRITE_RETRIES = 5;
    private static final long WRITE_RETRY_BASE_MS = 500;

    // Serializa agregarPendientes, marcarResultados y leerMedidas para evitar condiciones de carrera
    // cuando el procesamiento de un lote y la subida asincrónica del anterior se solapan.
    private final Object fileLock = new Object();

    public Map<String, MedidaSku> leerMedidas(Path excelPath) throws Exception {
        synchronized (fileLock) {
            return leerMedidasInterno(excelPath);
        }
    }

    private Map<String, MedidaSku> leerMedidasInterno(Path excelPath) throws Exception {
        if (!Files.exists(excelPath)) {
            crearArchivoVacio(excelPath);
            return new LinkedHashMap<>();
        }

        Map<String, MedidaSku> medidas = new LinkedHashMap<>();
        try (OPCPackage pkg = OPCPackage.open(excelPath.toFile(), PackageAccess.READ);
             Workbook workbook = new XSSFWorkbook(pkg)) {

            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) return medidas;

            Row header = sheet.getRow(0);
            if (header == null) return medidas;

            int skuCol = -1, productoCol = -1;
            int anchoCol = -1, altoCol = -1, profundidadCol = -1, pesoCol = -1;
            int anchoMasCol = -1, altoMasCol = -1, profundidadMasCol = -1, pesoMasCol = -1;
            int subidoCol = -1, errorCol = -1;

            for (int i = 0; i < header.getLastCellNum(); i++) {
                Cell cell = header.getCell(i);
                if (cell == null) continue;
                String h = normalizarHeader(getCellString(cell));
                boolean mas20 = h.contains("+20");

                if (h.equals("SKU")) skuCol = i;
                else if (h.startsWith("PRODUCTO")) productoCol = i;
                else if (h.equals("SUBIDO")) subidoCol = i;
                else if (h.equals("ERROR")) errorCol = i;
                else if (h.startsWith("ANCHO")) {
                    if (mas20) anchoMasCol = i; else anchoCol = i;
                }
                else if (h.startsWith("ALTO")) {
                    if (mas20) altoMasCol = i; else altoCol = i;
                }
                else if (h.startsWith("PROFUN") || h.startsWith("LARGO")) {
                    if (mas20) profundidadMasCol = i; else profundidadCol = i;
                }
                else if (h.startsWith("PESO")) {
                    if (mas20) pesoMasCol = i; else pesoCol = i;
                }
            }

            if (skuCol == -1) {
                throw new IllegalArgumentException(
                        "El Excel de medidas no tiene columna 'SKU'. Revise el archivo: " + excelPath);
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                Cell skuCell = row.getCell(skuCol);
                if (skuCell == null) continue;
                String sku = getCellString(skuCell).trim();
                if (sku.isEmpty()) continue;

                String producto = productoCol != -1 ? getCellString(row.getCell(productoCol)).trim() : "";
                Double ancho = anchoCol != -1 ? getCellDouble(row.getCell(anchoCol)) : null;
                Double alto = altoCol != -1 ? getCellDouble(row.getCell(altoCol)) : null;
                Double profundidad = profundidadCol != -1 ? getCellDouble(row.getCell(profundidadCol)) : null;
                Double peso = pesoCol != -1 ? getCellDouble(row.getCell(pesoCol)) : null;
                Double anchoMas = anchoMasCol != -1 ? getCellDouble(row.getCell(anchoMasCol)) : null;
                Double altoMas = altoMasCol != -1 ? getCellDouble(row.getCell(altoMasCol)) : null;
                Double profundidadMas = profundidadMasCol != -1 ? getCellDouble(row.getCell(profundidadMasCol)) : null;
                Double pesoMas = pesoMasCol != -1 ? getCellDouble(row.getCell(pesoMasCol)) : null;
                boolean subido = subidoCol != -1 && esSubido(getCellString(row.getCell(subidoCol)));
                String error = errorCol != -1 ? getCellString(row.getCell(errorCol)).trim() : "";

                medidas.put(sku, new MedidaSku(sku, producto,
                        ancho, alto, profundidad, peso,
                        anchoMas, altoMas, profundidadMas, pesoMas,
                        subido, error));
            }
        }
        return medidas;
    }

    /**
     * Inserta los SKUs pendientes que aún no figuran. Reusa primero filas existentes con SKU vacío
     * (típicamente pre-cargadas con fórmulas tipo =BUSCARX en PRODUCTO o base*1.2 en las +20%) y
     * si se acaban, appendea al final. Preserva todas las fórmulas existentes. SUBIDO se inicializa a "NO".
     * No escribe la columna PRODUCTO: queda delegada a la fórmula que el usuario tenga configurada.
     * Devuelve la cantidad de SKUs nuevos agregados.
     */
    public int agregarPendientes(Path excelPath, Collection<String> skusNuevos) throws Exception {
        if (skusNuevos == null || skusNuevos.isEmpty()) return 0;

        synchronized (fileLock) {
            Map<String, MedidaSku> existentes = leerMedidasInterno(excelPath);
            Collection<String> aAgregar = skusNuevos.stream()
                    .filter(s -> s != null && !s.isBlank())
                    .map(String::trim)
                    .filter(s -> !existentes.containsKey(s))
                    .distinct()
                    .toList();

            if (aAgregar.isEmpty()) return 0;

            try (XSSFWorkbook workbook = abrirOCrear(excelPath)) {
                Sheet sheet = workbook.getSheetAt(0);
                if (sheet == null) sheet = workbook.createSheet("MEDIDAS");

                asegurarHeaders(workbook, sheet);

                CellStyle skuPendienteStyle = crearEstiloPendiente(workbook);
                CellStyle celdaFaltanteStyle = crearEstiloCeldaFaltante(workbook);
                CellStyle subidoNoStyle = crearEstiloSubidoNo(workbook);

                // Buscar slots reutilizables: filas existentes con SKU vacío. Permite reusar filas
                // pre-cargadas con fórmulas (ej: PRODUCTO con =BUSCARX(...)) en lugar de appendear al final.
                java.util.Deque<Integer> slots = new java.util.ArrayDeque<>();
                int lastRowNum = sheet.getLastRowNum();
                for (int r = 1; r <= lastRowNum; r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;
                    Cell skuCell = row.getCell(COL_SKU);
                    if (skuCell == null || getCellString(skuCell).trim().isEmpty()) {
                        slots.offer(r);
                    }
                }

                int nextAppendRow = lastRowNum + 1;
                if (nextAppendRow < 1) nextAppendRow = 1;

                for (String sku : aAgregar) {
                    Row row;
                    if (!slots.isEmpty()) {
                        row = sheet.getRow(slots.poll());
                    } else {
                        row = sheet.createRow(nextAppendRow++);
                    }

                    Cell skuCell = row.getCell(COL_SKU);
                    if (skuCell == null) skuCell = row.createCell(COL_SKU, CellType.STRING);
                    skuCell.setCellValue(sku);
                    skuCell.setCellStyle(skuPendienteStyle);

                    // PRODUCTO: NO tocar. El usuario usa una fórmula (ej: =BUSCARX) para resolver la descripción.
                    // Al escribir el SKU arriba, la fórmula se recalcula sola al abrir el Excel.

                    // Celdas de medidas: solo resetear si no tienen fórmula ni valor cargado.
                    for (int c = COL_ANCHO; c <= COL_PESO_MAS; c++) {
                        Cell cell = row.getCell(c);
                        if (cell != null && cell.getCellType() == CellType.FORMULA) continue;
                        if (cell == null) cell = row.createCell(c, CellType.BLANK);
                        else cell.setBlank();
                        cell.setCellStyle(celdaFaltanteStyle);
                    }

                    Cell subidoCell = row.getCell(COL_SUBIDO);
                    if (subidoCell == null) subidoCell = row.createCell(COL_SUBIDO, CellType.STRING);
                    subidoCell.setCellValue("NO");
                    subidoCell.setCellStyle(subidoNoStyle);

                    // ERROR vacío — se rellena si una subida falla.
                    Cell errorCell = row.getCell(COL_ERROR);
                    if (errorCell == null) row.createCell(COL_ERROR, CellType.BLANK);
                    else errorCell.setBlank();
                }

                // Forzar que Excel recalcule las fórmulas al abrir el archivo (nuevos SKU pueden disparar
                // fórmulas tipo BUSCARX/XLOOKUP en PRODUCTO u otras columnas).
                workbook.setForceFormulaRecalculation(true);

                asegurarColumnaError(workbook, sheet);
                autoSizeColumns(sheet);
                escribirWorkbook(excelPath, workbook);
            }

            AppLogger.info("MEDIDAS - Agregados " + aAgregar.size() + " SKU(s) pendientes al Excel de medidas.");
            return aAgregar.size();
        }
    }

    /**
     * Escribe en una sola pasada los resultados de una subida a ML:
     *   - SKUs en skusOk: SUBIDO=SI (verde) y se limpia la celda ERROR.
     *   - SKUs en erroresPorSku: SUBIDO queda NO y se escribe el mensaje en la columna ERROR (rojo).
     * Devuelve la cantidad total de filas modificadas.
     */
    public int marcarResultados(Path excelPath, Collection<String> skusOk,
                                Map<String, String> erroresPorSku) throws Exception {
        boolean hayOk = skusOk != null && !skusOk.isEmpty();
        boolean hayErr = erroresPorSku != null && !erroresPorSku.isEmpty();
        if (!hayOk && !hayErr) return 0;
        if (!Files.exists(excelPath)) return 0;

        synchronized (fileLock) {
            int actualizados = 0;
            try (XSSFWorkbook workbook = abrirOCrear(excelPath)) {
                Sheet sheet = workbook.getSheetAt(0);
                if (sheet == null) return 0;

                asegurarColumnaError(workbook, sheet);

                int skuCol = -1, subidoCol = -1, errorCol = -1;
                Row header = sheet.getRow(0);
                if (header == null) return 0;
                for (int i = 0; i < header.getLastCellNum(); i++) {
                    Cell cell = header.getCell(i);
                    if (cell == null) continue;
                    String h = normalizarHeader(getCellString(cell));
                    if (h.equals("SKU")) skuCol = i;
                    else if (h.equals("SUBIDO")) subidoCol = i;
                    else if (h.equals("ERROR")) errorCol = i;
                }
                if (skuCol == -1 || subidoCol == -1) return 0;

                CellStyle subidoSiStyle = crearEstiloSubidoSi(workbook);
                CellStyle errorStyle = crearEstiloError(workbook);
                CellStyle clearStyle = crearEstiloErrorVacio(workbook);

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;
                    String sku = getCellString(row.getCell(skuCol)).trim();
                    if (sku.isEmpty()) continue;

                    if (hayOk && skusOk.contains(sku)) {
                        Cell cell = row.getCell(subidoCol);
                        if (cell == null) cell = row.createCell(subidoCol, CellType.STRING);
                        cell.setCellValue("SI");
                        cell.setCellStyle(subidoSiStyle);
                        // Limpiar ERROR si había uno previo.
                        if (errorCol != -1) {
                            Cell ec = row.getCell(errorCol);
                            if (ec == null) ec = row.createCell(errorCol, CellType.BLANK);
                            ec.setBlank();
                            ec.setCellStyle(clearStyle);
                        }
                        actualizados++;
                    } else if (hayErr && erroresPorSku.containsKey(sku) && errorCol != -1) {
                        Cell ec = row.getCell(errorCol);
                        if (ec == null) ec = row.createCell(errorCol, CellType.STRING);
                        ec.setCellValue(erroresPorSku.get(sku));
                        ec.setCellStyle(errorStyle);
                        actualizados++;
                    }
                }

                if (actualizados > 0) {
                    escribirWorkbook(excelPath, workbook);
                }
            }

            if (actualizados > 0) {
                AppLogger.info("MEDIDAS - Resultados aplicados en " + actualizados + " fila(s).");
            }
            return actualizados;
        }
    }

    private XSSFWorkbook abrirOCrear(Path excelPath) throws Exception {
        if (Files.exists(excelPath)) {
            try (FileInputStream fis = new FileInputStream(excelPath.toFile())) {
                return new XSSFWorkbook(fis);
            }
        }
        if (excelPath.getParent() != null) {
            Files.createDirectories(excelPath.getParent());
        }
        XSSFWorkbook wb = new XSSFWorkbook();
        wb.createSheet("MEDIDAS");
        return wb;
    }

    private void asegurarHeaders(Workbook workbook, Sheet sheet) {
        Row header = sheet.getRow(0);
        if (header != null && header.getCell(0) != null
                && "SKU".equalsIgnoreCase(getCellString(header.getCell(0)).trim())) {
            return;
        }
        if (header == null) header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
        aplicarBordes(headerStyle);

        for (int i = 0; i < HEADERS.length; i++) {
            Cell c = header.createCell(i, CellType.STRING);
            c.setCellValue(HEADERS[i]);
            c.setCellStyle(headerStyle);
        }
    }

    private CellStyle crearEstiloPendiente(Workbook workbook) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        XSSFColor amarillo = new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xE0, (byte) 0x82}, null);
        style.setFillForegroundColor(amarillo);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        aplicarBordes(style);
        return style;
    }

    private CellStyle crearEstiloCeldaFaltante(Workbook workbook) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        XSSFColor amarilloTenue = new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xF3, (byte) 0xCD}, null);
        style.setFillForegroundColor(amarilloTenue);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        aplicarBordes(style);
        return style;
    }

    private CellStyle crearEstiloSubidoNo(Workbook workbook) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        XSSFColor rojoTenue = new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xD6, (byte) 0xD6}, null);
        style.setFillForegroundColor(rojoTenue);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        aplicarBordes(style);
        return style;
    }

    private CellStyle crearEstiloSubidoSi(Workbook workbook) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        XSSFColor verdeTenue = new XSSFColor(new byte[]{(byte) 0xD4, (byte) 0xED, (byte) 0xDA}, null);
        style.setFillForegroundColor(verdeTenue);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        aplicarBordes(style);
        return style;
    }

    private CellStyle crearEstiloError(Workbook workbook) {
        // Error de ML: texto rojo oscuro sobre fondo rosa pálido, con wrap para mensajes largos.
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        XSSFColor rosaPalido = new XSSFColor(new byte[]{(byte) 0xFE, (byte) 0xE2, (byte) 0xE2}, null);
        style.setFillForegroundColor(rosaPalido);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        Font font = workbook.createFont();
        font.setBold(true);
        XSSFColor rojoOscuro = new XSSFColor(new byte[]{(byte) 0x99, (byte) 0x1B, (byte) 0x1B}, null);
        ((org.apache.poi.xssf.usermodel.XSSFFont) font).setColor(rojoOscuro);
        style.setFont(font);
        aplicarBordes(style);
        return style;
    }

    private CellStyle crearEstiloErrorVacio(Workbook workbook) {
        // Estilo limpio para la celda ERROR cuando el SKU pasó a SI.
        CellStyle style = workbook.createCellStyle();
        aplicarBordes(style);
        return style;
    }

    private void asegurarColumnaError(Workbook workbook, Sheet sheet) {
        Row header = sheet.getRow(0);
        if (header == null) return;
        Cell errorHeaderCell = header.getCell(COL_ERROR);
        String h = errorHeaderCell == null ? "" : normalizarHeader(getCellString(errorHeaderCell));
        if (h.equals("ERROR")) return;

        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
        aplicarBordes(headerStyle);

        if (errorHeaderCell == null) errorHeaderCell = header.createCell(COL_ERROR, CellType.STRING);
        errorHeaderCell.setCellValue("ERROR");
        errorHeaderCell.setCellStyle(headerStyle);
    }

    private void aplicarBordes(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    private void autoSizeColumns(Sheet sheet) {
        for (int i = 0; i < HEADERS.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * Escribe el workbook con reintentos. En Windows el archivo puede estar bloqueado si el usuario
     * lo tiene abierto en Excel (sharing violation). Reintenta con espera progresiva antes de rendirse.
     */
    private void escribirWorkbook(Path excelPath, Workbook workbook) throws IOException {
        IOException ultimoError = null;
        for (int intento = 1; intento <= MAX_WRITE_RETRIES; intento++) {
            try (FileOutputStream fos = new FileOutputStream(excelPath.toFile())) {
                workbook.write(fos);
                return;
            } catch (IOException e) {
                ultimoError = e;
                if (intento < MAX_WRITE_RETRIES) {
                    long espera = WRITE_RETRY_BASE_MS * intento;
                    AppLogger.warn("MEDIDAS - Archivo bloqueado (¿abierto en Excel?). Reintento "
                            + intento + "/" + MAX_WRITE_RETRIES + " en " + espera + "ms");
                    try {
                        Thread.sleep(espera);
                    } catch (InterruptedException ie) {
                        Thread.currentThread().interrupt();
                        throw e;
                    }
                }
            }
        }
        throw new IOException("No se pudo escribir el Excel de medidas después de "
                + MAX_WRITE_RETRIES + " intentos. ¿El archivo está abierto en Excel?", ultimoError);
    }

    private void crearArchivoVacio(Path excelPath) throws Exception {
        if (excelPath.getParent() != null) {
            Files.createDirectories(excelPath.getParent());
        }
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("MEDIDAS");
            asegurarHeaders(workbook, sheet);
            autoSizeColumns(sheet);
            escribirWorkbook(excelPath, workbook);
        }
        AppLogger.info("MEDIDAS - Archivo creado: " + excelPath);
    }

    private static String normalizarHeader(String raw) {
        if (raw == null) return "";
        return raw.replace('\n', ' ')
                .replaceAll("\\s+", " ")
                .trim()
                .toUpperCase();
    }

    private static boolean esSubido(String raw) {
        if (raw == null) return false;
        String v = raw.trim().toUpperCase();
        return v.equals("SI") || v.equals("SÍ") || v.equals("YES") || v.equals("TRUE") || v.equals("1");
    }

    private static String getCellString(Cell cell) {
        if (cell == null) return "";
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) type = cell.getCachedFormulaResultType();
        return switch (type) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
                yield v == Math.floor(v) && !Double.isInfinite(v)
                        ? String.valueOf((long) v)
                        : String.valueOf(v);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    private static Double getCellDouble(Cell cell) {
        if (cell == null) return null;
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) type = cell.getCachedFormulaResultType();
        return switch (type) {
            case NUMERIC -> cell.getNumericCellValue();
            case STRING -> {
                String s = cell.getStringCellValue().trim().replace(",", ".");
                if (s.isEmpty()) yield null;
                try { yield Double.parseDouble(s); }
                catch (NumberFormatException e) { yield null; }
            }
            default -> null;
        };
    }
}
