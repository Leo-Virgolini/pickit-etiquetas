package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.Embalaje;
import ar.com.leo.etiquetas.model.MedidaSku;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class MedidasExcelManagerTest {

    @TempDir
    Path tempDir;

    private final MedidasExcelManager manager = new MedidasExcelManager();

    /** Headers del Excel de medidas tal como existían antes de agregar la columna EMBALAJE. */
    private static final String[] HEADERS_VIEJOS = {
            "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
            "Peso físico\n(empaque + producto)\nkg",
            "Ancho +20%", "Alto +20%", "Profunidad +20%",
            "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR"
    };

    // -------------------------------------------------------------------------------------------
    // Catálogo de embalajes
    // -------------------------------------------------------------------------------------------

    @Test
    void catalogoDeUnArchivoSinHojaEmbalajesEsVacio() throws Exception {
        Path excel = crearExcelViejo("medidas.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));

        Map<String, Embalaje> catalogo = manager.leerCatalogoEmbalajes(excel);

        assertTrue(catalogo.isEmpty());
    }

    @Test
    void catalogoDeUnArchivoInexistenteEsVacio() throws Exception {
        Map<String, Embalaje> catalogo = manager.leerCatalogoEmbalajes(tempDir.resolve("no-existe.xlsx"));

        assertTrue(catalogo.isEmpty());
    }

    @Test
    void catalogoLeeCodigoTipoYMedidas() throws Exception {
        Path excel = crearExcelConCatalogo("con-catalogo.xlsx", List.<Object[]>of(
                new Object[]{"CAJA 3", "CAJA", 30.0, 20.0, 15.0}
        ));

        Map<String, Embalaje> catalogo = manager.leerCatalogoEmbalajes(excel);

        Embalaje caja = catalogo.get("CAJA 3");
        assertNotNull(caja);
        assertEquals("CAJA 3", caja.codigo());
        assertEquals("CAJA", caja.tipo());
        assertEquals(30.0, caja.anchoCm());
        assertEquals(20.0, caja.altoCm());
        assertEquals(15.0, caja.profundidadCm());
    }

    @Test
    void catalogoIndexaPorCodigoNormalizadoPreservandoElOriginal() throws Exception {
        Path excel = crearExcelConCatalogo("normalizado.xlsx", List.<Object[]>of(
                new Object[]{"  Bolsa   Chica ", "BOLSA", null, null, null}
        ));

        Map<String, Embalaje> catalogo = manager.leerCatalogoEmbalajes(excel);

        assertEquals("Bolsa Chica", catalogo.get("BOLSA CHICA").codigo());
    }

    @Test
    void catalogoIgnoraFilasSinCodigo() throws Exception {
        Path excel = crearExcelConCatalogo("con-vacias.xlsx", List.<Object[]>of(
                new Object[]{"CAJA 1", "CAJA", 10.0, 10.0, 10.0},
                new Object[]{"   ", "CAJA", 20.0, 20.0, 20.0},
                new Object[]{"CAJA 2", "CAJA", 30.0, 30.0, 30.0}
        ));

        Map<String, Embalaje> catalogo = manager.leerCatalogoEmbalajes(excel);

        assertEquals(2, catalogo.size());
    }

    // -------------------------------------------------------------------------------------------
    // Columna EMBALAJE en la hoja de SKUs
    // -------------------------------------------------------------------------------------------

    @Test
    void leerMedidasDeUnArchivoViejoDejaElEmbalajeVacio() throws Exception {
        Path excel = crearExcelViejo("viejo.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));

        MedidaSku medida = manager.leerMedidas(excel).get("1241212");

        assertEquals("", medida.embalaje());
        assertEquals("Producto A", medida.producto());
    }

    @Test
    void leerMedidasDevuelveElCodigoDeEmbalajeCargado() throws Exception {
        Path excel = crearExcelViejo("con-embalaje.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));
        escribirCeldaEnFilaDeSku(excel, 12, "CAJA 3");

        MedidaSku medida = manager.leerMedidas(excel).get("1241212");

        assertEquals("CAJA 3", medida.embalaje());
    }

    @Test
    void agregarPendientesCreaElHeaderEmbalajeEnUnArchivoViejo() throws Exception {
        Path excel = crearExcelViejo("sin-header.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("EMBALAJE", leerHeader(excel, 11));
        assertEquals("ERROR", leerHeader(excel, 12));
    }

    @Test
    void agregarPendientesCreaLaHojaCatalogoConHeaders() throws Exception {
        Path excel = crearExcelViejo("sin-hoja.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Sheet catalogo = wb.getSheet("EMBALAJES");
            assertNotNull(catalogo);
            assertEquals("CÓDIGO", catalogo.getRow(0).getCell(0).getStringCellValue());
            // Solo headers: sin filas de ejemplo.
            assertEquals(0, catalogo.getLastRowNum());
        }
    }

    @Test
    void agregarPendientesNoPisaLasFilasDelCatalogoExistente() throws Exception {
        Path excel = crearExcelConCatalogo("catalogo-existente.xlsx", List.<Object[]>of(
                new Object[]{"CAJA 3", "CAJA", 30.0, 20.0, 15.0}
        ));

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("CAJA 3", manager.leerCatalogoEmbalajes(excel).get("CAJA 3").codigo());
    }

    @Test
    void agregarPendientesPreservaLasFormulasExistentes() throws Exception {
        Path excel = crearExcelViejo("con-formula.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));
        escribirFormulaEnFilaDeSku(excel, 6, "C2*1.2");

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheetAt(0).getRow(1).getCell(6);
            assertEquals(CellType.FORMULA, cell.getCellType());
            assertEquals("C2*1.2", cell.getCellFormula());
        }
    }

    @Test
    void marcarResultadosNoBorraElEmbalajeCargado() throws Exception {
        Path excel = crearExcelViejo("marcar.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));
        escribirCeldaEnFilaDeSku(excel, 12, "CAJA 3");

        manager.marcarResultados(excel, List.of("1241212"), Map.of());

        assertEquals("CAJA 3", manager.leerMedidas(excel).get("1241212").embalaje());
    }

    // -------------------------------------------------------------------------------------------
    // Migración de la estructura (columna + hoja + desplegable)
    // -------------------------------------------------------------------------------------------

    @Test
    void sinSkusNuevosIgualSeCreaLaEstructura() throws Exception {
        // Estado estacionario: todos los SKU del lote ya figuran en el Excel. Es el camino habitual,
        // así que la migración no puede depender de que aparezca un SKU nuevo.
        Path excel = crearExcelViejo("estacionario.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));

        manager.asegurarEstructuraEmbalajes(excel);

        assertEquals("EMBALAJE", leerHeader(excel, 11));
        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            assertNotNull(wb.getSheet("EMBALAJES"));
        }
    }

    @Test
    void asegurarEstructuraEsIdempotenteYNoAcumulaValidaciones() throws Exception {
        Path excel = crearExcelViejo("idempotente.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));

        manager.asegurarEstructuraEmbalajes(excel);
        manager.asegurarEstructuraEmbalajes(excel);
        manager.asegurarEstructuraEmbalajes(excel);

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            assertEquals(1, wb.getSheetAt(0).getDataValidations().size());
        }
    }

    @Test
    void elDesplegableQuedaVisible() throws Exception {
        Path excel = crearExcelViejo("desplegable.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));

        manager.asegurarEstructuraEmbalajes(excel);

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            assertFalse(wb.getSheetAt(0).getDataValidations().get(0).getSuppressDropDownArrow());
        }
    }

    @Test
    void noPisaUnaColumnaPropiaDelUsuario() throws Exception {
        Path excel = crearExcelViejo("columna-propia.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));
        escribirHeader(excel, 12, "OBSERVACIONES");

        manager.asegurarEstructuraEmbalajes(excel);

        // EMBALAJE se inserta antes de ERROR: todo lo que estaba a la derecha corre un lugar.
        assertEquals("EMBALAJE", leerHeader(excel, 11));
        assertEquals("ERROR", leerHeader(excel, 12));
        assertEquals("OBSERVACIONES", leerHeader(excel, 13));
    }

    @Test
    void insertarLaColumnaNoPierdeElContenidoDeError() throws Exception {
        Path excel = crearExcelViejo("con-error.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));
        escribirCelda(excel, 1, 11, "Item no encontrado");

        manager.asegurarEstructuraEmbalajes(excel);

        assertEquals("Item no encontrado", leerCelda(excel, 1, 12));
        assertEquals("", leerCelda(excel, 1, 11));
    }

    @Test
    void elHeaderDelCatalogoTieneFondoGrisYBordes() throws Exception {
        Path excel = crearExcelViejo("estilo-catalogo.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));

        manager.asegurarEstructuraEmbalajes(excel);

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            CellStyle estilo = wb.getSheet("EMBALAJES").getRow(0).getCell(0).getCellStyle();
            assertEquals(FillPatternType.SOLID_FOREGROUND, estilo.getFillPattern());
            assertNotNull(((XSSFCellStyle) estilo).getFillForegroundColorColor());
            assertEquals(BorderStyle.THIN, estilo.getBorderTop());
            assertEquals(BorderStyle.THIN, estilo.getBorderBottom());
        }
    }

    @Test
    void reusaLaColumnaEmbalajeAunqueNoEsteEnElIndiceEsperado() throws Exception {
        Path excel = crearExcelViejo("columna-corrida.xlsx", List.<String[]>of(new String[]{"1241212", "Producto A"}));
        escribirHeader(excel, 12, "OBSERVACIONES");
        escribirHeader(excel, 13, "EMBALAJE");

        manager.asegurarEstructuraEmbalajes(excel);

        // Ya existe: se reusa donde está, sin mover nada.
        assertEquals("EMBALAJE", leerHeader(excel, 13));
        assertEquals("OBSERVACIONES", leerHeader(excel, 12));
        assertNull(leerHeader(excel, 14));
    }

    @Test
    void laHojaDeMedidasSeResuelvePorNombreNoPorPosicion() throws Exception {
        // El usuario puede arrastrar EMBALAJES al primer lugar en Excel.
        Path excel = crearExcelConCatalogo("hojas-invertidas.xlsx", List.<Object[]>of(
                new Object[]{"CAJA 3", "CAJA", 30.0, 20.0, 15.0}
        ));
        moverHojaAlPrincipio(excel, "EMBALAJES");

        Map<String, MedidaSku> medidas = manager.leerMedidas(excel);

        assertNotNull(medidas.get("1241212"), "la hoja de medidas debe seguir siendo la de SKUs");
    }

    @Test
    void leerMedidasYCatalogoDevuelveAmbosEnUnaSolaLectura() throws Exception {
        Path excel = crearExcelConCatalogo("una-lectura.xlsx", List.<Object[]>of(
                new Object[]{"CAJA 3", "CAJA", 30.0, 20.0, 15.0}
        ));

        MedidasExcelManager.DatosMedidas datos = manager.leerMedidasYCatalogo(excel);

        assertNotNull(datos.medidas().get("1241212"));
        assertNotNull(datos.catalogo().get("CAJA 3"));
    }

    // -------------------------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------------------------

    private Path crearExcelViejo(String nombre, List<String[]> filas) throws Exception {
        Path excel = tempDir.resolve(nombre);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("MEDIDAS");
            Row header = sheet.createRow(0);
            for (int i = 0; i < HEADERS_VIEJOS.length; i++) {
                header.createCell(i, CellType.STRING).setCellValue(HEADERS_VIEJOS[i]);
            }
            int r = 1;
            for (String[] fila : filas) {
                Row row = sheet.createRow(r++);
                for (int i = 0; i < fila.length; i++) {
                    row.createCell(i, CellType.STRING).setCellValue(fila[i]);
                }
            }
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
        return excel;
    }

    private Path crearExcelConCatalogo(String nombre, List<Object[]> embalajes) throws Exception {
        Path excel = crearExcelViejo(nombre, List.<String[]>of(new String[]{"1241212", "Producto A"}));
        try (Workbook wb = abrirParaEditar(excel)) {
            Sheet sheet = wb.createSheet("EMBALAJES");
            Row header = sheet.createRow(0);
            String[] headers = {"CÓDIGO", "TIPO", "Ancho cm", "Alto cm", "Profundidad cm"};
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i, CellType.STRING).setCellValue(headers[i]);
            }
            int r = 1;
            for (Object[] e : embalajes) {
                Row row = sheet.createRow(r++);
                row.createCell(0, CellType.STRING).setCellValue((String) e[0]);
                row.createCell(1, CellType.STRING).setCellValue((String) e[1]);
                for (int i = 2; i < e.length; i++) {
                    if (e[i] == null) continue;
                    row.createCell(i, CellType.NUMERIC).setCellValue((Double) e[i]);
                }
            }
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
        return excel;
    }

    private void escribirCeldaEnFilaDeSku(Path excel, int col, String valor) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            Sheet sheet = wb.getSheetAt(0);
            sheet.getRow(0).createCell(col, CellType.STRING).setCellValue("EMBALAJE");
            sheet.getRow(1).createCell(col, CellType.STRING).setCellValue(valor);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void escribirFormulaEnFilaDeSku(Path excel, int col, String formula) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.getSheetAt(0).getRow(1).createCell(col, CellType.FORMULA).setCellFormula(formula);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    /**
     * Carga el workbook completo en memoria y suelta el archivo, para poder reescribirlo encima.
     * WorkbookFactory.create(file, ..., false) mantiene el archivo tomado y lo corrompe al guardar.
     */
    private Workbook abrirParaEditar(Path excel) throws Exception {
        try (java.io.InputStream in = java.nio.file.Files.newInputStream(excel)) {
            return new XSSFWorkbook(in);
        }
    }

    private void escribirHeader(Path excel, int col, String valor) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.getSheetAt(0).getRow(0).createCell(col, CellType.STRING).setCellValue(valor);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void moverHojaAlPrincipio(Path excel, String nombre) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.setSheetOrder(nombre, 0);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void escribirCelda(Path excel, int fila, int col, String valor) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.getSheetAt(0).getRow(fila).createCell(col, CellType.STRING).setCellValue(valor);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private String leerCelda(Path excel, int fila, int col) throws Exception {
        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheetAt(0).getRow(fila).getCell(col);
            if (cell == null || cell.getCellType() == CellType.BLANK) return "";
            return cell.getStringCellValue();
        }
    }

    private String leerHeader(Path excel, int col) throws Exception {
        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheetAt(0).getRow(0).getCell(col);
            return cell == null ? null : cell.getStringCellValue();
        }
    }
}
