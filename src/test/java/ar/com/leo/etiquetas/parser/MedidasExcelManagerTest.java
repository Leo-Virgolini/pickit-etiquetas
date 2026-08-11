package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import ar.com.leo.etiquetas.model.MedidaSku;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;

class MedidasExcelManagerTest {

    @TempDir
    Path tempDir;

    private final MedidasExcelManager manager = new MedidasExcelManager();

    /** Columnas base, sin ninguna de las de embalaje. */
    private static final String[] HEADERS_BASE = {
            "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
            "Peso físico\n(empaque + producto)\nkg",
            "Ancho +20%", "Alto +20%", "Profunidad +20%",
            "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR"
    };

    /** Columnas base más las ocho de embalaje entre SUBIDO y ERROR, como las carga el usuario. */
    private static final String[] HEADERS_CON_EMBALAJE = {
            "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
            "Peso físico\n(empaque + producto)\nkg",
            "Ancho +20%", "Alto +20%", "Profunidad +20%",
            "Peso físico (empaque + producto) +20%", "SUBIDO",
            "N° Bolsa", "Nombre Caja", "N° Caja", "PLURIBOL", "CANT PLURIBOL",
            "ROLLO INFLABLE", "CANT PAÑOS", "OBSERVACIONES",
            "ERROR"
    };

    // -------------------------------------------------------------------------------------------
    // Lectura de las columnas de embalaje
    // -------------------------------------------------------------------------------------------

    @Test
    void leeLasOchoColumnasDeEmbalaje() throws Exception {
        Path excel = crearExcel("completo.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "5", "GRANDE", "3", "SI", "2", "DIAMANTE", "4", "Colchon + Tapa"});

        DatosEmbalaje datos = manager.leerMedidas(excel).get("1241212").embalaje();

        assertEquals("5", datos.nroBolsa());
        assertEquals("GRANDE", datos.nombreCaja());
        assertEquals("3", datos.nroCaja());
        assertEquals("SI", datos.pluribol());
        assertEquals("2", datos.cantPluribol());
        assertEquals("DIAMANTE", datos.rollo());
        assertEquals("4", datos.cantPanos());
        assertEquals("Colchon + Tapa", datos.observaciones());
    }

    @Test
    void unArchivoSinLasColumnasDeEmbalajeDevuelveDatosVacios() throws Exception {
        Path excel = crearExcel("sin-embalaje.xlsx", HEADERS_BASE,
                new String[]{"1241212", "Producto A"});

        MedidaSku medida = manager.leerMedidas(excel).get("1241212");

        assertSame(DatosEmbalaje.VACIO, medida.embalaje());
        assertEquals("Producto A", medida.producto(), "el resto de la fila se sigue leyendo");
    }

    @Test
    void lasColumnasSeUbicanPorHeaderNoPorPosicion() throws Exception {
        // Mismo contenido pero con las columnas de embalaje en otro orden y al final.
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR",
                "OBSERVACIONES", "N° Caja", "Nombre Caja"
        };
        Path excel = crearExcel("desordenado.xlsx", headers,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO", "",
                        "Colchon", "3", "GRANDE"});

        DatosEmbalaje datos = manager.leerMedidas(excel).get("1241212").embalaje();

        assertEquals("3", datos.nroCaja());
        assertEquals("GRANDE", datos.nombreCaja());
        assertEquals("Colchon", datos.observaciones());
    }

    @Test
    void aceptaPanosEscritoSinEnie() throws Exception {
        String[] headers = {"SKU", "PRODUCTO", "SUBIDO", "ROLLO INFLABLE", "CANT PANOS"};
        Path excel = crearExcel("sin-enie.xlsx", headers,
                new String[]{"1241212", "Producto A", "NO", "DIAMANTE", "4"});

        DatosEmbalaje datos = manager.leerMedidas(excel).get("1241212").embalaje();

        assertEquals("4", datos.cantPanos());
    }

    @Test
    void distingueCantPluribolDePluribol() throws Exception {
        String[] headers = {"SKU", "PRODUCTO", "SUBIDO", "PLURIBOL", "CANT PLURIBOL"};
        Path excel = crearExcel("pluribol.xlsx", headers,
                new String[]{"1241212", "Producto A", "NO", "SI", "2"});

        DatosEmbalaje datos = manager.leerMedidas(excel).get("1241212").embalaje();

        assertEquals("SI", datos.pluribol());
        assertEquals("2", datos.cantPluribol());
    }

    @Test
    void distingueNombreCajaDeNumeroDeCaja() throws Exception {
        String[] headers = {"SKU", "PRODUCTO", "SUBIDO", "Nombre Caja", "N° Caja"};
        Path excel = crearExcel("cajas.xlsx", headers,
                new String[]{"1241212", "Producto A", "NO", "GRANDE", "3"});

        DatosEmbalaje datos = manager.leerMedidas(excel).get("1241212").embalaje();

        assertEquals("GRANDE", datos.nombreCaja());
        assertEquals("3", datos.nroCaja());
    }

    @Test
    void unNumeroDeCajaNumericoSeLeeSinDecimales() throws Exception {
        Path excel = crearExcel("numerico.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});
        escribirNumero(excel, 1, 13, 3);

        assertEquals("3", manager.leerMedidas(excel).get("1241212").embalaje().nroCaja());
    }

    // -------------------------------------------------------------------------------------------
    // El resto del Excel sigue funcionando
    // -------------------------------------------------------------------------------------------

    @Test
    void agregarPendientesNoEscribeEnLasColumnasDeEmbalaje() throws Exception {
        Path excel = crearExcel("pendientes.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});

        manager.agregarPendientes(excel, List.of("999999"));

        DatosEmbalaje datos = manager.leerMedidas(excel).get("999999").embalaje();
        assertSame(DatosEmbalaje.VACIO, datos);
    }

    @Test
    void agregarPendientesPreservaLasFormulasExistentes() throws Exception {
        Path excel = crearExcel("con-formula.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});
        escribirFormula(excel, 1, 6, "C2*1.2");

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheetAt(0).getRow(1).getCell(6);
            assertEquals(CellType.FORMULA, cell.getCellType());
            assertEquals("C2*1.2", cell.getCellFormula());
        }
    }

    @Test
    void marcarResultadosNoBorraLosDatosDeEmbalaje() throws Exception {
        Path excel = crearExcel("marcar.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "", "GRANDE", "3"});

        manager.marcarResultados(excel, List.of("1241212"), Map.of());

        assertEquals("GRANDE", manager.leerMedidas(excel).get("1241212").embalaje().nombreCaja());
    }

    @Test
    void unArchivoNuevoSeCreaConLasColumnasDeEmbalaje() throws Exception {
        Path excel = tempDir.resolve("nuevo.xlsx");

        manager.leerMedidas(excel);

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertEquals("SUBIDO", header.getCell(10).getStringCellValue());
            assertEquals("N° Bolsa", header.getCell(11).getStringCellValue());
            assertEquals("OBSERVACIONES", header.getCell(18).getStringCellValue());
            assertEquals("ERROR", header.getCell(19).getStringCellValue());
        }
    }

    @Test
    void noSeCreaLaHojaCatalogoDeEmbalajes() throws Exception {
        Path excel = tempDir.resolve("sin-catalogo.xlsx");

        manager.leerMedidas(excel);

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            assertEquals(1, wb.getNumberOfSheets());
            assertNotNull(wb.getSheet("MEDIDAS"));
        }
    }

    // -------------------------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------------------------

    private Path crearExcel(String nombre, String[] headers, String[] fila) throws Exception {
        Path excel = tempDir.resolve(nombre);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("MEDIDAS");
            Row header = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i, CellType.STRING).setCellValue(headers[i]);
            }
            Row row = sheet.createRow(1);
            for (int i = 0; i < fila.length; i++) {
                row.createCell(i, CellType.STRING).setCellValue(fila[i]);
            }
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
        return excel;
    }

    private void escribirNumero(Path excel, int fila, int col, double valor) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.getSheetAt(0).getRow(fila).createCell(col, CellType.NUMERIC).setCellValue(valor);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void escribirFormula(Path excel, int fila, int col, String formula) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.getSheetAt(0).getRow(fila).createCell(col, CellType.FORMULA).setCellFormula(formula);
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
        try (InputStream in = Files.newInputStream(excel)) {
            return new XSSFWorkbook(in);
        }
    }
}
