package ar.com.leo.etiquetas.parser;

import ar.com.leo.AppLogger;
import ar.com.leo.etiquetas.model.DatosEmbalaje;
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
 * Gestiona el Excel "madre" de medidas de embalaje por SKU:
 *
 *   SKU | PRODUCTO | Largo cm | Ancho cm | Alto cm | Peso físico (empaque + producto) kg
 *       | Largo +20% | Ancho +20% | Alto +20% | Peso físico (empaque + producto) +5%
 *       | SUBIDO | ESTANDARIZADO | ENVASE | TIPO DE ROLLO | CANT PAÑOS | OBSERVACIONES | ERROR
 *
 * Las columnas con porcentaje son las que se suben a la API de ML (dims en cm, peso en kg). Todas
 * se ubican por su encabezado, no por posición: el usuario reordena y agrega columnas propias.
 */
public class MedidasExcelManager {

    public static final String[] HEADERS = {
            "SKU",
            "PRODUCTO",
            "Largo\ncm",
            "Ancho\ncm",
            "Alto\ncm",
            "Peso físico\n(empaque + producto)\nkg",
            "Largo +20%",
            "Ancho +20%",
            "Alto +20%",
            "Peso físico (empaque + producto) +5%",
            "SUBIDO",
            "ESTANDARIZADO",
            "ENVASE",
            "TIPO DE ROLLO",
            "CANT PAÑOS",
            "OBSERVACIONES",
            "ERROR"
    };

    /** Marca una columna de margen: "+20%" en las dimensiones, "+5%" en el peso. */
    private static final java.util.regex.Pattern PORCENTAJE =
            java.util.regex.Pattern.compile("\\+\\s*\\d+\\s*%");

    /** Hoja con el catálogo de envases: código en "N°" e inscripción en "INSCRIPCION". */
    public static final String HOJA_ESTANDARIZACION = "ESTANDARIZACION";

    // Retry para sharing violation (Excel abierto por el usuario).
    private static final int MAX_WRITE_RETRIES = 5;
    private static final long WRITE_RETRY_BASE_MS = 500;

    // Serializa agregarPendientes, marcarResultados y leerMedidas para evitar condiciones de carrera
    // cuando el procesamiento de un lote y la subida asincrónica del anterior se solapan.
    private final Object fileLock = new Object();

    /**
     * Índices de todas las columnas de una hoja, resueltos por encabezado. Se usa tanto para leer
     * como para escribir: la app no puede asumir posiciones fijas porque el usuario agrega y
     * reordena columnas propias (las de embalaje son suyas).
     */
    private record Columnas(int sku, int producto,
                            int ancho, int alto, int profundidad, int peso,
                            int anchoMas, int altoMas, int profundidadMas, int pesoMas,
                            int subido, int error,
                            int estandarizado, int envase,
                            int rollo, int panos, int observaciones) {

        /** Columnas de medidas que agregarPendientes limpia al reusar una fila. */
        int[] deMedidas() {
            return new int[]{ancho, alto, profundidad, peso, anchoMas, altoMas, profundidadMas, pesoMas};
        }

        /** Distingue la hoja de medidas de cualquier otra tabla del usuario que tenga un SKU. */
        boolean esDeLaApp() {
            return subido != -1 || error != -1 || ancho != -1 || alto != -1
                    || profundidad != -1 || peso != -1 || estandarizado != -1;
        }
    }

    /**
     * Resuelve las columnas por su encabezado normalizado.
     *
     * Las de medidas se evalúan antes que las de embalaje: un encabezado como "Ancho caja cm" es
     * natural —lo que se mide es la caja— y con los patrones de embalaje primero terminaría
     * asignado a la columna de caja, dejando la dimensión sin leer.
     */
    private Columnas resolverColumnas(Sheet sheet) {
        int sku = -1, producto = -1;
        int ancho = -1, alto = -1, profundidad = -1, peso = -1;
        int anchoMas = -1, altoMas = -1, profundidadMas = -1, pesoMas = -1;
        int subido = -1, error = -1;
        int estandarizado = -1, envase = -1;
        int rollo = -1, panos = -1, observaciones = -1;

        Row header = sheet == null ? null : sheet.getRow(0);
        if (header == null) {
            return new Columnas(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
                    -1, -1, -1, -1, -1);
        }

        for (int i = 0; i < header.getLastCellNum(); i++) {
            Cell cell = header.getCell(i);
            if (cell == null) continue;
            String h = normalizarHeader(getCellString(cell));
            // Las columnas de margen son las que llevan un porcentaje: +20% en las dimensiones y
            // +5% en el peso. No alcanza con buscar "+": el encabezado del peso base es
            // "Peso físico (empaque + producto) kg" y también lo tiene.
            boolean mas20 = PORCENTAJE.matcher(h).find();

            if (h.equals("SKU")) sku = i;
            else if (h.startsWith("PRODUCTO")) producto = i;
            else if (h.equals("SUBIDO")) subido = i;
            else if (h.equals("ERROR")) error = i;
            else if (h.startsWith("ESTANDARIZ")) estandarizado = i;
            else if (h.startsWith("ANCHO")) {
                if (mas20) anchoMas = i; else ancho = i;
            }
            else if (h.startsWith("ALTO")) {
                if (mas20) altoMas = i; else alto = i;
            }
            else if (h.startsWith("PROFUN") || h.startsWith("LARGO")) {
                if (mas20) profundidadMas = i; else profundidad = i;
            }
            else if (h.startsWith("PESO")) {
                if (mas20) pesoMas = i; else peso = i;
            }
            // Columnas de embalaje.
            else if (h.contains("ENVASE")) envase = i;
            else if (h.contains("ROLLO")) rollo = i;
            // El encabezado puede venir con o sin eñe.
            else if (h.contains("PAÑO") || h.contains("PANO")) panos = i;
            else if (h.contains("OBSERV")) observaciones = i;
        }

        return new Columnas(sku, producto, ancho, alto, profundidad, peso,
                anchoMas, altoMas, profundidadMas, pesoMas, subido, error,
                estandarizado, envase, rollo, panos, observaciones);
    }

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

        try (OPCPackage pkg = OPCPackage.open(excelPath.toFile(), PackageAccess.READ);
             Workbook workbook = new XSSFWorkbook(pkg)) {
            return leerMedidasDe(workbook);
        }
    }

    private Map<String, MedidaSku> leerMedidasDe(Workbook workbook) {
        Map<String, MedidaSku> medidas = new LinkedHashMap<>();
        Sheet sheet = hojaMedidas(workbook);
        if (sheet == null) return medidas;

        // Una hoja sin ninguna fila es un archivo recién creado: se devuelve vacío para que
        // agregarPendientes lo inicialice. Si tiene filas pero no encabezados, en cambio, es un
        // archivo equivocado y el error tiene que llegarle al usuario.
        if (sheet.getPhysicalNumberOfRows() == 0) return medidas;

        Columnas cols = resolverColumnas(sheet);
        if (cols.sku() == -1) {
            throw new IllegalArgumentException(
                    "El Excel de medidas no tiene columna 'SKU'. Revise el archivo.");
        }

        Map<String, String> inscripciones = leerInscripciones(workbook);

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            Cell skuCell = row.getCell(cols.sku());
            if (skuCell == null) continue;
            String sku = getCellString(skuCell).trim();
            if (sku.isEmpty()) continue;

            String envase = celda(row, cols.envase());
            DatosEmbalaje embalaje = new DatosEmbalaje(
                    envase,
                    inscripciones.getOrDefault(normalizarHeader(envase), ""),
                    celda(row, cols.rollo()),
                    celda(row, cols.panos()),
                    celda(row, cols.observaciones()),
                    esSubido(celda(row, cols.estandarizado())),
                    cols.estandarizado() != -1);
            if (embalaje.equals(DatosEmbalaje.VACIO)) embalaje = DatosEmbalaje.VACIO;

            medidas.put(sku, new MedidaSku(sku,
                    celda(row, cols.producto()),
                    numero(row, cols.ancho()),
                    numero(row, cols.alto()),
                    numero(row, cols.profundidad()),
                    numero(row, cols.peso()),
                    numero(row, cols.anchoMas()),
                    numero(row, cols.altoMas()),
                    numero(row, cols.profundidadMas()),
                    numero(row, cols.pesoMas()),
                    cols.subido() != -1 && esSubido(getCellString(row.getCell(cols.subido()))),
                    celda(row, cols.error()),
                    embalaje));
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
                Sheet sheet = hojaMedidasParaEscribir(workbook);
                if (sheet == null) {
                    if (workbook.getNumberOfSheets() > 0) return 0;
                    sheet = workbook.createSheet("MEDIDAS");
                }

                asegurarHeaders(workbook, sheet);
                // Todo se escribe por índice resuelto, no por posición fija: el usuario puede
                // haber intercalado sus columnas entre las de la app. Las columnas que la app
                // administra se crean si faltan, así nunca hay que caer a un índice adivinado.
                int colSubido = asegurarColumna(workbook, sheet, "SUBIDO");
                int colError = asegurarColumna(workbook, sheet, "ERROR");
                Columnas cols = resolverColumnas(sheet);
                int colSku = cols.sku();

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
                    Cell skuCell = row.getCell(colSku);
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

                    Cell skuCell = row.getCell(colSku);
                    if (skuCell == null) skuCell = row.createCell(colSku, CellType.STRING);
                    skuCell.setCellValue(sku);
                    skuCell.setCellStyle(skuPendienteStyle);

                    // PRODUCTO: NO tocar. El usuario usa una fórmula (ej: =BUSCARX) para resolver la descripción.
                    // Al escribir el SKU arriba, la fórmula se recalcula sola al abrir el Excel.

                    // Celdas de medidas: solo resetear si no tienen fórmula ni valor cargado.
                    // Se recorren las columnas resueltas y no un rango de índices: entre medio
                    // pueden estar las columnas de embalaje, que son del usuario y no se tocan.
                    for (int c : cols.deMedidas()) {
                        if (c == -1) continue;
                        Cell cell = row.getCell(c);
                        if (cell != null && cell.getCellType() == CellType.FORMULA) continue;
                        if (cell == null) cell = row.createCell(c, CellType.BLANK);
                        else cell.setBlank();
                        cell.setCellStyle(celdaFaltanteStyle);
                    }

                    Cell subidoCell = row.getCell(colSubido);
                    if (subidoCell == null) subidoCell = row.createCell(colSubido, CellType.STRING);
                    subidoCell.setCellValue("NO");
                    subidoCell.setCellStyle(subidoNoStyle);

                    // ERROR vacío — se rellena si una subida falla.
                    Cell errorCell = row.getCell(colError);
                    if (errorCell == null) row.createCell(colError, CellType.BLANK);
                    else errorCell.setBlank();

                }

                // Forzar que Excel recalcule las fórmulas al abrir el archivo (nuevos SKU pueden disparar
                // fórmulas tipo BUSCARX/XLOOKUP en PRODUCTO u otras columnas).
                workbook.setForceFormulaRecalculation(true);

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
                Sheet sheet = hojaMedidasParaEscribir(workbook);
                if (sheet == null) return 0;

                // SUBIDO y ERROR se aseguran igual que en agregarPendientes: si el archivo no las
                // tiene, esta subida no tendría dónde dejar el resultado y se perdería en silencio.
                // Se hace después de comprobar el SKU para no tocar el archivo si no es el correcto.
                if (resolverColumnas(sheet).sku() == -1) return 0;
                asegurarColumna(workbook, sheet, "SUBIDO");
                int errorCol = asegurarColumna(workbook, sheet, "ERROR");

                Columnas cols = resolverColumnas(sheet);
                int skuCol = cols.sku(), subidoCol = cols.subido();

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

    /**
     * Escribe los encabezados por defecto solo en una hoja que todavía no los tiene. La condición
     * es que no haya columna SKU en ninguna posición, no que SKU esté en la A: el usuario puede
     * tener columnas propias a la izquierda, y reescribir la fila le borraría sus encabezados
     * —incluidos los de embalaje— dejando cada valor bajo el título equivocado.
     */
    private void asegurarHeaders(Workbook workbook, Sheet sheet) {
        if (buscarColumna(sheet, "SKU") != -1) return;

        Row header = sheet.getRow(0);
        if (header == null) header = sheet.createRow(0);

        CellStyle headerStyle = crearEstiloHeader(workbook);

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

    /**
     * Índice de una columna que la app administra, creándola al final si el archivo no la tiene.
     *
     * Siempre al final de lo que haya: en una posición fija se arriesga a pisar una columna del
     * usuario, y saltando hasta la posición nominal quedarían columnas en blanco sin encabezado en
     * el medio. La posición nominal solo aplica a un archivo nuevo, que nace con todos los headers.
     */
    /**
     * Cuántas columnas ocupa la hoja mirando todas las filas, no solo el encabezado: el usuario
     * puede tener una columna con datos y sin título, y escribirle encima la destruiría.
     */
    private int anchoUsado(Sheet sheet) {
        int ancho = 0;
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row != null) ancho = Math.max(ancho, row.getLastCellNum());
        }
        return Math.max(0, ancho);
    }

    private int asegurarColumna(Workbook workbook, Sheet sheet, String nombre) {
        int existente = buscarColumna(sheet, nombre);
        if (existente != -1) return existente;

        Row header = sheet.getRow(0);
        if (header == null) header = sheet.createRow(0);

        int destino = anchoUsado(sheet);
        Cell headerCell = header.getCell(destino);
        if (headerCell == null) headerCell = header.createCell(destino, CellType.STRING);
        headerCell.setCellValue(nombre);
        headerCell.setCellStyle(crearEstiloHeader(workbook));
        return destino;
    }

    /**
     * Hoja de medidas: la que tenga columna SKU y además alguna columna propia de la app (SUBIDO,
     * ERROR o una dimensión). Se busca así y no por posición porque el usuario puede tener hojas
     * propias antes de ella.
     *
     * No alcanza con la columna SKU: una tabla de búsqueda del usuario —de las que alimentan sus
     * BUSCARX— también la tiene, y escribirle los SKU pendientes la destruiría. Si ninguna hoja
     * califica se cae a la primera, para que el error de "falta la columna SKU" siga saliendo con
     * un mensaje entendible en vez de devolver vacío en silencio.
     */
    /**
     * Hoja donde la app puede escribir: la que tiene SKU y alguna columna propia. Si hay una hoja
     * con SKU que no califica —típicamente la tabla de búsqueda del usuario— se devuelve null antes
     * que arriesgarse a escribirle encima. Solo cuando no hay ninguna hoja con SKU se asume que el
     * archivo está por inicializarse y se usa la primera.
     */
    private Sheet hojaMedidasParaEscribir(Workbook workbook) {
        Sheet conSku = null;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (esCatalogo(sheet)) continue;
            Columnas cols = resolverColumnas(sheet);
            if (cols.sku() == -1) continue;
            if (cols.esDeLaApp()) return sheet;
            if (conSku == null) conSku = sheet;
        }
        if (conSku != null) {
            AppLogger.warn("MEDIDAS - La única hoja con columna SKU no parece la de medidas "
                    + "(no tiene SUBIDO, ERROR ni dimensiones). No se escribe nada para no dañarla.");
            return null;
        }
        return workbook.getNumberOfSheets() == 0 ? null : workbook.getSheetAt(0);
    }

    private Sheet hojaMedidas(Workbook workbook) {
        Sheet soloConSku = null;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (esCatalogo(sheet)) continue;
            Columnas cols = resolverColumnas(sheet);
            if (cols.sku() == -1) continue;
            if (cols.esDeLaApp()) return sheet;
            if (soloConSku == null) soloConSku = sheet;
        }
        // Sin ninguna que tenga columnas de la app, la que tenga SKU es la mejor candidata: puede
        // ser una hoja de medidas todavía incompleta. Recién si no hay ninguna se cae a la primera,
        // para que el error de "falta la columna SKU" salga con un mensaje entendible.
        if (soloConSku != null) return soloConSku;
        return workbook.getNumberOfSheets() == 0 ? null : workbook.getSheetAt(0);
    }

    /**
     * La hoja de envases no es candidata a hoja de medidas: sus columnas LARGO/ANCHO/ALTO ya la
     * harían pasar por una, y basta con que algún día gane una columna SKU para que la app le
     * empiece a escribir encima.
     */
    private boolean esCatalogo(Sheet sheet) {
        return sheet == null || HOJA_ESTANDARIZACION.equalsIgnoreCase(sheet.getSheetName());
    }

    /** Índice de una columna buscándola por su header normalizado, o -1 si no existe. */
    private int buscarColumna(Sheet sheet, String headerBuscado) {
        Row header = sheet.getRow(0);
        if (header == null) return -1;
        for (int i = 0; i < header.getLastCellNum(); i++) {
            Cell cell = header.getCell(i);
            if (cell == null) continue;
            if (normalizarHeader(getCellString(cell)).equals(headerBuscado)) return i;
        }
        return -1;
    }

    private CellStyle crearEstiloHeader(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
        aplicarBordes(headerStyle);
        return headerStyle;
    }

    private void aplicarBordes(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    /** Dimensiona las columnas que la hoja tiene de verdad, no solo las que la app conoce. */
    private void autoSizeColumns(Sheet sheet) {
        Row header = sheet.getRow(0);
        int columnas = header == null ? HEADERS.length : Math.max(HEADERS.length, header.getLastCellNum());
        for (int i = 0; i < columnas; i++) {
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

    /** Mayúsculas, sin espacios en los extremos y con los espacios internos colapsados. */
    private static String normalizarHeader(String raw) {
        if (raw == null) return "";
        return raw.replace('\n', ' ')
                .replaceAll("\\s+", " ")
                .trim()
                .toUpperCase();
    }

    /**
     * Inscripción de cada envase, indexada por su código normalizado. Sale de la hoja
     * ESTANDARIZACION: la columna "N°" trae el código (BOL-1, CAJ-1) y "INSCRIPCION" el texto que
     * está escrito en el envase físico.
     *
     * Si la hoja no está, se devuelve vacío: la etiqueta muestra el código solo. Es un dato de
     * referencia, no una validación, y no vale la pena frenar el lote por eso.
     */
    private Map<String, String> leerInscripciones(Workbook workbook) {
        Map<String, String> inscripciones = new LinkedHashMap<>();
        Sheet sheet = workbook.getSheet(HOJA_ESTANDARIZACION);
        if (sheet == null) return inscripciones;

        Row header = sheet.getRow(0);
        if (header == null) return inscripciones;

        int codigoCol = -1, inscripcionCol = -1;
        for (int i = 0; i < header.getLastCellNum(); i++) {
            Cell cell = header.getCell(i);
            if (cell == null) continue;
            String h = normalizarHeader(getCellString(cell));
            // Primera coincidencia: un segundo encabezado que empiece igual (p. ej. "N° ROLLOS")
            // no debe robarle la columna al código.
            if (codigoCol == -1 && (h.startsWith("N°") || h.equals("N") || h.startsWith("COD"))) codigoCol = i;
            else if (inscripcionCol == -1 && h.startsWith("INSCRIP")) inscripcionCol = i;
        }
        if (codigoCol == -1 || inscripcionCol == -1) return inscripciones;

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String codigo = celda(row, codigoCol);
            if (codigo.isEmpty()) continue;
            inscripciones.putIfAbsent(normalizarHeader(codigo), celda(row, inscripcionCol));
        }
        return inscripciones;
    }

    /** Valor de una celda de la fila, o cadena vacía si la columna no existe en este archivo. */
    private static String celda(Row row, int col) {
        return col == -1 ? "" : getCellString(row.getCell(col)).trim();
    }

    /** Valor numérico de una celda, o null si la columna no existe o no tiene número. */
    private static Double numero(Row row, int col) {
        return col == -1 ? null : getCellDouble(row.getCell(col));
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

    /**
     * Valor numérico de una celda, o null si no lo es.
     *
     * Se exige que la celda sea numérica de verdad —o una fórmula con resultado numérico, que es
     * como están cargadas las columnas de margen— y se rechaza el texto aunque se pueda parsear:
     * un "38,4" pegado como texto suele ser un dato mal cargado, y de ahí sale la medida que se
     * publica en ML. Las fórmulas con error también quedan afuera.
     */
    private static Double getCellDouble(Cell cell) {
        if (cell == null) return null;
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) type = cell.getCachedFormulaResultType();
        return type == CellType.NUMERIC ? cell.getNumericCellValue() : null;
    }
}
