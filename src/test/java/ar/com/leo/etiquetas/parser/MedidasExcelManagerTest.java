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

    @Test
    void laHojaDeMedidasSeUbicaPorSuColumnaSku() throws Exception {
        // El usuario puede tener hojas propias ("Resumen", "Notas") antes de la de medidas.
        Path excel = crearExcel("hoja-ajena.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});
        agregarHojaAjenaAlPrincipio(excel, "Resumen");

        MedidaSku medida = manager.leerMedidas(excel).get("1241212");

        assertNotNull(medida, "debe encontrar la hoja de medidas aunque no sea la primera");
    }

    @Test
    void unaMedidaQueMencionaLaCajaNoSeConfundeConLaColumnaDeCaja() throws Exception {
        // "Ancho caja cm" es un nombre natural: lo que se mide es la caja.
        String[] headers = {"SKU", "PRODUCTO", "Ancho caja cm", "Alto de la caja", "SUBIDO", "N° Caja"};
        Path excel = crearExcel("medida-caja.xlsx", headers,
                new String[]{"1241212", "Producto A", "30", "20", "NO", "3"});

        MedidaSku medida = manager.leerMedidas(excel).get("1241212");

        assertEquals(30.0, medida.anchoCm(), "la medida no debe ir a parar a la columna de caja");
        assertEquals(20.0, medida.altoCm());
        assertEquals("3", medida.embalaje().nroCaja());
    }

    @Test
    void agregarPendientesUbicaLasColumnasPorHeader() throws Exception {
        // Columnas de embalaje intercaladas antes de las +20%, como habilita el README.
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "N° Bolsa", "Nombre Caja", "N° Caja",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR"
        };
        Path excel = crearExcel("intercaladas.xlsx", headers,
                new String[]{"", "", "", "", "", "", "5", "GRANDE", "3"});

        manager.agregarPendientes(excel, List.of("999999"));

        // La fila con SKU vacío se reusa: los datos de embalaje que el usuario cargó ahí no se pisan.
        MedidaSku medida = manager.leerMedidas(excel).get("999999");
        assertEquals("GRANDE", medida.embalaje().nombreCaja());
        assertEquals("3", medida.embalaje().nroCaja());
        assertEquals("5", medida.embalaje().nroBolsa());
    }

    @Test
    void agregarPendientesEscribeElNoEnLaColumnaSubido() throws Exception {
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "N° Bolsa", "Nombre Caja", "N° Caja",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR"
        };
        Path excel = crearExcel("subido.xlsx", headers, new String[]{"1241212", "Producto A"});

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Row fila = wb.getSheetAt(0).getRow(2);
            assertEquals("NO", fila.getCell(13).getStringCellValue(), "el NO va en SUBIDO");
            assertEquals(CellType.BLANK, fila.getCell(9).getCellType(), "no en Ancho +20%");
        }
    }

    @Test
    void noPisaLosEncabezadosCuandoSkuNoEstaEnLaPrimeraColumna() throws Exception {
        // El usuario agregó una columna propia a la izquierda, así que SKU quedó en la B.
        String[] headers = {"ID interno", "SKU", "PRODUCTO", "SUBIDO", "N° Caja", "ERROR"};
        Path excel = crearExcel("sku-corrido.xlsx", headers,
                new String[]{"X-1", "1241212", "Producto A", "NO", "3", ""});

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("ID interno", leerHeader(excel, 0), "no se deben reescribir los encabezados");
        assertEquals("SKU", leerHeader(excel, 1));
        assertEquals("3", manager.leerMedidas(excel).get("1241212").embalaje().nroCaja());
    }

    @Test
    void noEligeUnaHojaAuxiliarDelUsuarioComoHojaDeMedidas() throws Exception {
        // Una tabla de búsqueda con columna SKU no es la hoja de medidas: no tiene ni SUBIDO ni
        // las columnas de dimensiones, y escribir ahí destruiría los datos del usuario.
        Path excel = crearExcel("medidas.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});
        agregarHojaAuxiliarAlPrincipio(excel, "Catalogo", new String[]{"SKU", "DESCRIPCION"});

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Sheet auxiliar = wb.getSheet("Catalogo");
            assertEquals(2, auxiliar.getRow(0).getLastCellNum(), "la hoja auxiliar no se toca");
            assertEquals(0, auxiliar.getLastRowNum());
        }
        assertNotNull(manager.leerMedidas(excel).get("999999"), "el SKU va a la hoja de medidas");
    }

    @Test
    void laColumnaErrorSeCreaAlFinalSinPisarColumnasPropias() throws Exception {
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%", "SUBIDO",
                "N° Bolsa", "Nombre Caja", "N° Caja", "PLURIBOL", "CANT PLURIBOL",
                "ROLLO INFLABLE", "CANT PAÑOS", "OBSERVACIONES",
                "MI COLUMNA"
        };
        Path excel = crearExcel("sin-error.xlsx", headers,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "", "", "", "", "", "", "", "", "dato mio"});

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("MI COLUMNA", leerHeader(excel, 19), "no se pisa la columna del usuario");
        assertEquals("ERROR", leerHeader(excel, 20));
        assertEquals("dato mio", leerCelda(excel, 1, 19));
    }

    @Test
    void unaHojaSinEncabezadosNoRompeLaLectura() throws Exception {
        // Un .xlsx recién creado a mano: la app tiene que poder inicializarlo.
        Path excel = tempDir.resolve("hoja-vacia.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("MEDIDAS");
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }

        assertEquals(Map.of(), manager.leerMedidas(excel));

        manager.agregarPendientes(excel, List.of("999999"));
        assertEquals("SKU", leerHeader(excel, 0));
        assertNotNull(manager.leerMedidas(excel).get("999999"));
    }

    @Test
    void siFaltaLaColumnaSubidoSeCreaEnVezDeEscribirPorPosicion() throws Exception {
        // El usuario renombró SUBIDO, así que la app no la reconoce. Escribir el "NO" en el índice
        // 10 caería sobre una de sus columnas de embalaje.
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%",
                "N° Bolsa", "Nombre Caja", "N° Caja"
        };
        Path excel = crearExcel("sin-subido.xlsx", headers,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "",
                        "5", "GRANDE", "3"});

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("N° Bolsa", leerHeader(excel, 10), "no se pisa la columna de embalaje");
        assertEquals("5", manager.leerMedidas(excel).get("1241212").embalaje().nroBolsa());
        assertEquals("SUBIDO", leerHeader(excel, 13));
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

    private void agregarHojaAuxiliarAlPrincipio(Path excel, String nombre, String[] headers) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            Sheet aux = wb.createSheet(nombre);
            Row header = aux.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i, CellType.STRING).setCellValue(headers[i]);
            }
            wb.setSheetOrder(nombre, 0);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void agregarHojaAjenaAlPrincipio(Path excel, String nombre) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            Sheet ajena = wb.createSheet(nombre);
            ajena.createRow(0).createCell(0, CellType.STRING).setCellValue("dato propio");
            wb.setSheetOrder(nombre, 0);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private String leerHeader(Path excel, int col) throws Exception {
        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheet("MEDIDAS").getRow(0).getCell(col);
            return cell == null ? null : cell.getStringCellValue();
        }
    }

    private String leerCelda(Path excel, int fila, int col) throws Exception {
        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheet("MEDIDAS").getRow(fila).getCell(col);
            if (cell == null || cell.getCellType() == CellType.BLANK) return "";
            return cell.getStringCellValue();
        }
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
